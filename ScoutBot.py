from warnings import warn
from time import sleep, time
from datetime import datetime, timedelta
from ConfigParser import SafeConfigParser
import dateutil.parser
import os
from pytz import timezone
import re
import json
import signal
import sys
import shelve

# HelpScout API
import helpscout

# Google Calendar API modules
import httplib2
from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

# slackbot stuff
from slackclient import SlackClient

CALENDAR_REFRESH_INTERVAL = timedelta(minutes=10)
TZ = timezone('US/Pacific')

HELPSCOUT_SCAN_INTERVAL = timedelta(minutes=1)
HELPSCOUT_TIMEOUT = 30

CALENDAR_SCAN_INTERVAL = timedelta(minutes=5)

ANNOYANCE_FREQUENCY = timedelta(minutes=10)

def text2int(textnum, numwords={}):
    if not numwords:
      units = [
        "zero", "one", "two", "three", "four", "five", "six", "seven", "eight",
        "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen",
        "sixteen", "seventeen", "eighteen", "nineteen",
      ]

      tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]

      scales = ["hundred", "thousand", "million", "billion", "trillion"]

      numwords["and"] = (1, 0)
      for idx, word in enumerate(units):    numwords[word] = (1, idx)
      for idx, word in enumerate(tens):     numwords[word] = (1, idx * 10)
      for idx, word in enumerate(scales):   numwords[word] = (10 ** (idx * 3 or 2), 0)

    current = result = 0
    for word in textnum.split():
        if word not in numwords:
          raise Exception("Illegal word: " + word)

        scale, increment = numwords[word]
        current = current * scale + increment
        if scale > 100:
            result += current
            current = 0

    return result + current


# helper to deal with bizarre helpscout paging interface - you have to
# call each method multiple times until it returns nothing.
def helpscout_pager(meth, *args, **kwargs):
    while True:
        page = meth(*args, **kwargs)
        if page is None or page.items is None or len(page.items) == 0:
            return
        for item in page.items:
            yield item


# make a human-readable timedelta, from
# http://stackoverflow.com/a/13756038/153053.  Tweaked to not show
# seconds, not important for this use.
def td_format(td_object):
    seconds = int(td_object.total_seconds())
    periods = [
        ('year',        60*60*24*365),
        ('month',       60*60*24*30),
        ('day',         60*60*24),
        ('hour',        60*60),
        ('minute',      60)
                ]

    if not seconds:
        return "0 seconds"
    
    strings=[]
    for period_name,period_seconds in periods:
        if seconds >= period_seconds:
            period_value , seconds = divmod(seconds,period_seconds)
            if period_value == 1:
                strings.append("%s %s" % (period_value, period_name))
            else:
                strings.append("%s %ss" % (period_value, period_name))
                
    return ", ".join(strings)

def human_time(dt):
    time = dt.strftime('%I:%M%P').lstrip('0')
    time = re.sub(r':00', '', time)
    return time

# a generic timeout wrapper
def timeout(func, args=(), kwargs={}, timeout_duration=1, default=None):
    import signal

    class TimeoutError(Exception):
        pass

    def handler(signum, frame):
        raise TimeoutError()

    # set the timeout handler
    signal.signal(signal.SIGALRM, handler) 
    signal.alarm(timeout_duration)
    try:
        result = func(*args, **kwargs)
    except TimeoutError as exc:
        result = default
    finally:
        signal.alarm(0)

    return result

def _body_len_minus_sig(body):
    if not re.search(r'---?\r?\n', body, re.S):
        return len(body)

    match = re.search(r'^(.*)---?\r?\n', body, re.S)
    return len(match.group(1))
    
class ScoutBot:
    def __init__(self, config_file='scoutbot.cfg'):
        config = SafeConfigParser()
        config.read([config_file, os.path.expanduser('~/.' + config_file)])

        self.hs_api_key                 = config.get('helpscout',
                                                     'api_key')
        self.helpscout_current_tickets  = None
        self.last_helpscout_scan        = datetime.utcnow() - \
                                          HELPSCOUT_SCAN_INTERVAL
        self.support_domain             = config.get('scoutbot',
                                                     'support_domain')
        self.max_wait_new_ticket        = int(config.get('scoutbot',
                                                         'max_wait_new_ticket'))
        self.max_wait_response_or_close = int(config.get('scoutbot',
                                                         'max_wait_response_or_close'))
        self.support_calendar_id        = config.get('scoutbot',
                                                     'support_calendar_id')
        self.last_calender_scan        = datetime.utcnow() - \
                                         CALENDAR_SCAN_INTERVAL
        self.bugzilla_url              = config.get('bugzilla', 'url')

        self.slack_api_key              = config.get('slack', 'api_key')
        self.slack_bot_name             = config.get('slack', 'bot_name')
        self.slack_last_ping            = 0
        self.slack_stack                = []
        self.slack_connected            = False

        self.support_open_at   = dateutil.parser.parse(
            config.get('scoutbot', 'support_open_at'))
        self.support_close_at  = dateutil.parser.parse(
            config.get('scoutbot', 'support_close_at'))

        support_open_days = json.loads(
            config.get('scoutbot', 'support_open_days'))
        self.support_open_days = support_open_days

        channels = json.loads(config.get('slack', 'channels'))
        self.slack_channels = channels

        log_channels = json.loads(config.get('slack', 'log_channels'))
        self.slack_log_channels = log_channels

        if self.hs_api_key is None:
            raise Exception("Missing api_key config value!")

        self.client = helpscout.Client()
        self.client.api_key = self.hs_api_key

        self.calendar = []
        self.calendar_refreshed_at = None

        self.last_alert_on_ticket = dict()
        self.last_alert_everyone_on_ticket = dict()
        self.last_hs_link = dict()
        self.last_bugzilla_link = dict()
        self.initial_alert_sent = set()

        self.memory = shelve.open('memory.db')
        if "quiet_users" not in self.memory:
            self.memory['quiet_users'] = set()
        if "ignore_list" not in self.memory:
            self.memory['ignore_list'] = set()
        if "snooze" not in self.memory:
            self.memory['snooze'] = dict()
        if "unsub" not in self.memory:
            self.memory['unsub'] = set()

        # is this too rude?  Maybe weird if ScoutBot gets used by
        # other code...
        def signal_handler(signal, frame):
            print("\nExiting SlackBot.  Thank you for playing.\n")
            self.memory.sync()
            sys.exit(0)
        signal.signal(signal.SIGINT, signal_handler)

    def open_conversations(self, hours=24, status='active'):
        client = self.client

        # this API wrapper has the wackiest paging system...  Gotta
        # clear it here or subsequent calls will silently find
        # nothing!
        client.clearstate()
        
        results = []
        for mailbox in list(helpscout_pager(client.mailboxes)):
            # look back up to 6 hours by default
            start_date = datetime.utcnow() - timedelta(hours = hours)
            start_date = start_date.replace(microsecond=0).isoformat() + 'Z'

            client.clearstate()
            for conv in helpscout_pager(client.conversations_for_mailbox,
                                        mailbox.id,
                                        status=status,
                                        modifiedSince=start_date):
                results.append(self.parse_conversation(conv))
        return results

    def parse_conversation(self, conv):
        client = self.client

        data = dict(
            owner      = conv.owner,
            id         = conv.id,
            num        = conv.number,
            subject    = conv.subject,
            folder_id  = conv.folderid,
            url        = 'https://secure.helpscout.net/conversation/%s' % (conv.id,),
            created_at = dateutil.parser.parse(conv.createdat).replace(tzinfo=None)
        )

        # refetch to get thread info, needed to figure out last reply
        threads = client.conversation(conv.id).threads
        last_support_msg_at = None
        last_client_msg_at = None
        last_owner_email = None
        first = True
        last_body = None
        for thread in threads:
            created_at = dateutil.parser.parse(thread.createdat)\
                         .replace(tzinfo=None)

            email = thread.createdby['email']
            body = thread.body.lower() if thread.body else ''
            last_body = body
            
            if email.endswith(self.support_domain):
                if (last_support_msg_at is None or \
                    last_support_msg_at < created_at):
                    last_support_msg_at = created_at
                    last_owner_email = last_owner_email
            else:
                if last_client_msg_at is not None:
                    if (('thanks' in body or 'thank you' in body) and
                        _body_len_minus_sig(body) < 60):
                        self.log("Ignoring short thanks reply from client: %s"
                                 % (body))
                        continue

                if (last_client_msg_at is None or \
                    last_client_msg_at < created_at):
                    last_client_msg_at = created_at

        data['new'] = last_support_msg_at is None
        data['last_support_msg_at'] = last_support_msg_at
        data['last_client_msg_at'] = last_client_msg_at
        data['last_body'] = body

        if data['new']:
            data['needs_reply_or_close'] = True
        elif data['last_client_msg_at'] is None:
            # we're talking to ourselves now, great
            data['needs_reply_or_close'] = False
        else:
            data['needs_reply_or_close'] = last_client_msg_at > last_support_msg_at

        # if the last client contact came while support was closed and
        # support is now open, rebase last_client_msg_at at support
        # open time
        last_msg_at = (data['last_client_msg_at'] if
                       data['last_client_msg_at'] else
                       datetime.utcnow()).replace(
            tzinfo=timezone("UTC")).astimezone(tz=TZ)
        support_start_at = datetime.now(tz=TZ).replace(
            hour=self.support_open_at.hour,
            minute=0,
            second=0)

        if (last_msg_at < support_start_at):
            # ticket pre-dated the start of support, start the clock again
            data['wait_time'] = datetime.now(tz=TZ) - support_start_at
            if data['wait_time'].total_seconds() < 0:
                data['wait_time'] = datetime.utcnow() - datetime.utcnow()

            data['wait_time_human'] = td_format(data['wait_time'])

            self.log("Found ticket %s from before support started at (%d:00) - reseting wait time to %s" % (data['num'], self.support_open_at.hour, data['wait_time_human']))

        else:
            data['wait_time'] = datetime.utcnow() - (data['last_client_msg_at'] if data['last_client_msg_at'] else datetime.utcnow())
            data['wait_time_human'] = td_format(data['wait_time'])
        
        return data

    def watch(self, once=False):
        while True:
            self.scan_conversations()
            if once:
                return
            sleep(10)

    def helpscout_status(self):
        if not self.helpscout_current_tickets:
            return "Huh, not sure.  I might be having trouble reaching helpscout. Please find @sam and ask him to fix me."

        ignore_list = self.memory['ignore_list']
        snooze      = self.memory['snooze']

        summary = ["Currently active tickets modified within 24 hours:"]
        for ticket in self.helpscout_current_tickets:
            if ticket['num'] in ignore_list:
                summary.append("[<{url}|#{num}>] {subject} => *ignored*.".format(**ticket))
                continue

            if int(ticket['num']) in snooze and datetime.utcnow() < snooze[int(ticket['num'])]:
                summary.append("[<{url}|#{num}>] {subject} => *snoozed* until {time}.".format(time=snooze[int(ticket['num'])], **ticket))
                continue


            if ticket['new']:
                summary.append("[<{url}|#{num}>] {subject} => *new and unclaimed* {wait_time_human}.".format(**ticket))
            elif ticket['needs_reply_or_close']:
                summary.append("[<{url}|#{num}>] {subject} => *needs response or close* {wait_time_human}.".format(**ticket))
            else:
                summary.append("[<{url}|#{num}>] {subject} => *handled*.".format(**ticket))

        if len(summary) == 1:
            summary.append("None!")

        return "\n".join(summary)
    
    def scan_conversations(self):
        self.log("*** Scanning for conversations...")
        tickets = timeout(lambda: self.open_conversations(),
                          timeout_duration=HELPSCOUT_TIMEOUT,
                          default='TIMEOUT')
        if tickets == 'TIMEOUT':
            self.log("*** Timed out looking for conversation...")
            return
        self.helpscout_current_tickets = tickets

        for ticket in tickets:
            if ticket['new']:
                self.log("*** [%s] %s => new and unclaimed %s" % \
                         (ticket['num'],
                          ticket['subject'],
                          ticket['wait_time_human']))
                if ticket['wait_time'].total_seconds() > self.max_wait_new_ticket:
                    self.alert_support(ticket)

                if ticket['wait_time'].total_seconds() > (self.max_wait_new_ticket * 3):
                    self.alert_everyone(ticket, yell=True)
                elif ticket['wait_time'].total_seconds() > (self.max_wait_new_ticket * 2):
                    self.alert_everyone(ticket)

            elif ticket['needs_reply_or_close']:
                self.log("*** [%s] %s => needs response or close %s" % \
                         (ticket['num'],
                          ticket['subject'],
                          ticket['wait_time_human']))
                if ticket['wait_time'].total_seconds() > self.max_wait_response_or_close:
                    self.alert_support(ticket)
                if ticket['wait_time'].total_seconds() > (self.max_wait_response_or_close * 3):
                    self.alert_everyone(ticket, yell=True)
                elif ticket['wait_time'].total_seconds() > (self.max_wait_response_or_close * 2):
                    self.alert_everyone(ticket)

            else:
                self.log("+ [%s] %s => handled" % \
                         (ticket['num'],
                          ticket['subject']))

    def _support_closed(self):
        now = datetime.now(tz=TZ)
        if now.weekday() not in self.support_open_days:
            return True

        on_now = self.support_now(just_name=True).lower()
        if "closed" in on_now or "holiday" in on_now:
            return True

        if (now.hour >= self.support_open_at.hour and
            now.hour < self.support_close_at.hour):
            return False
        
        return True
                
    def alert_support(self, ticket):
        if self._support_closed():
            self.log("Ignoring [<{url}|#{num}>] for now, support is closed.".format(**ticket))
            return
        
        ignore_list = self.memory['ignore_list']
        if ticket['num'] in ignore_list:
            self.log("Ignoring [<{url}|#{num}>], it's on the ignore_list.".format(**ticket))
            return

        snooze = self.memory['snooze']
        if int(ticket['num']) in snooze and datetime.utcnow() < snooze[int(ticket['num'])]:
            self.log("Ignoring [<{url}|#{num}>], it's snoozed until {time}.".format(time=snooze[int(ticket['num'])], **ticket))
            return

        user =  self.support_now(just_name=True)
        if user:
            if ticket['num'] not in self.initial_alert_sent and user not in self.memory.quiet_users:
                self.slackbot_direct_message(user, "Ticket [<{url}|#{num}>] {subject} was opened.\nRespond 'quieter' to stop these messages (then 'louder' if you want them resumed).  Respond 'help' to see more options.".format(**ticket))
                self.initial_alert_sent.add(ticket['num'])
                return

            # don't alert too often on any given issue
            if ticket['num'] in self.last_alert_on_ticket:
                if ((datetime.utcnow() - self.last_alert_on_ticket[ticket['num']])
                    < ANNOYANCE_FREQUENCY):
                    return 

            self.log("+++ CALLING FOR HELP ON %s FROM %s +++" % (ticket['num'], user))
            self.slackbot_direct_message(user, "Ticket [<{url}|#{num}>] {subject} has been awaiting a response for {wait_time_human}.\nRespond 'ignore {num}' and I will ignore this ticket from now on, or 'snooze {num} [n]' to snooze for n minutes, 10 by default.  Respond 'help' to see more options.".format(**ticket))
            
        else:
            self.slackbot_broadcast("Ticket [<{url}|#{num}>] {subject} has been awaiting a response for {wait_time_human} and I couldn't figure out who is on support!\n(Tell me 'ignore {num}' to ignore it or 'snooze {num} [n]' to snooze for n minutes, 10 by default.)".format(**ticket))

        self.last_alert_on_ticket[ticket['num']] = datetime.utcnow()

    def alert_everyone(self, ticket, yell=False):
        if self._support_closed():
            self.log("Ignoring [<{url}|#{num}>] for now, support is closed.".format(**ticket))
            return

        ignore_list = self.memory['ignore_list']
        if ticket['num'] in ignore_list:
            self.log("Ignoring [<{url}|#{num}>], it's on the ignore_list.".format(**ticket))
            return

        snooze = self.memory['snooze']
        if int(ticket['num']) in snooze and datetime.utcnow() < snooze[int(ticket['num'])]:
            self.log("Ignoring [<{url}|#{num}>], it's snoozed until {time}.".format(time=snooze[int(ticket['num'])], **ticket))
            return

        # don't alert too often on any given issue
        if ticket['num'] in self.last_alert_everyone_on_ticket:
            if ((datetime.utcnow() - self.last_alert_everyone_on_ticket[ticket['num']]) < ANNOYANCE_FREQUENCY):
                return 

        # has it been a really long time?  Add in <!channel> for extra BOOM
        extra = "<!channel> " if yell else ""
        self.slackbot_broadcast("{extra}Ticket [<{url}|#{num}>] {subject} has been awaiting a response for {wait_time_human}!\n(Tell me 'ignore {num}' to ignore it or 'snooze {num} [n]' to snooze for n minutes, 10 by default.)".format(extra=extra, **ticket))

        self.last_alert_everyone_on_ticket[ticket['num']] = datetime.utcnow()

    def log(self, msg):
        if self.slack_connected:
            self.slackbot_log(msg)
        print "%s: %s" % (datetime.now(), msg)

    def support_now(self, just_name=False):
        cal = self.refresh_support_calendar()
        now = datetime.now(tz=TZ)
        for c in cal:
            if now >= c[0] and now <= c[1]:
                if just_name:
                    return c[2]
                return "%s is on support now." % (c[2],)
        if just_name:
            return None
        return "Nobody is on support now! :fire::fire::fire:"

    def support_day(self, offset=0):
        cal = self.refresh_support_calendar()
        now = datetime.now(tz=TZ) + timedelta(days=offset)

        today = []
        for c in cal:
            if now.date() == c[0].date():
                today.append("%s is on from %s to %s %s %s " % \
                             (c[2],
                              human_time(c[0]),
                              human_time(c[1]),
                              c[0].tzname(),
                              "%d day(s) from now" % \
                                offset if offset else "today"))

        if len(today):
            today = [re.sub(r'1 day\(s\) from now', 'tomorrow', x) \
                     for x in today]
            return "\n".join(today)
        return "Nobody is on support %s!" % \
               ("%d day(s) from now" % offset if offset else "today")

    def slack_name_for_full_name(self, orig):
        if not self.slack_connected or not self.slack_user_names:
            return orig

        full = orig.lower()
        parts = re.split(r'\W+', full)

        # look for a match on first name
        if parts[0] in self.slack_user_names:
            return "<@%s>" % (self.slack_user_names[parts[0]],)

        # look for initials
        initials = ''.join([p[0] for p in parts])
        if initials in self.slack_user_names:
            return "<@%s>" % (self.slack_user_names[initials],)

        # look for the case where someone uses a three-initial handle
        # but only lists their first and last on the calendar...  Lame.
        if len(parts) == 2:
            import string
            for middle in string.ascii_lowercase:
                initials = parts[0][0] + middle + parts[1][0]
                if initials in self.slack_user_names:
                    return "<@%s>" % (self.slack_user_names[initials],)

        # failure!
        return orig

    # pull a fresh calendar from Google periodically
    def refresh_support_calendar(self, use_cache=True):
        if (use_cache and
            len(self.calendar) and
            self.calendar_refreshed_at and
            (datetime.utcnow() - self.calendar_refreshed_at) <
              CALENDAR_REFRESH_INTERVAL):
            return self.calendar

        self.log("*** Refreshing support calendar...")
    
        SCOPES             = 'https://www.googleapis.com/auth/calendar.readonly'
        CLIENT_SECRET_FILE = 'google_api_client_secret.json'
        APPLICATION_NAME   = 'ScoutBot'

        credential_path    = 'google_api_credentials.json'
        store              = oauth2client.file.Storage(credential_path)
        credentials        = store.get()
        if not credentials:
            raise Exception("Please setup a valid Google API credentials "
                            "file in google_api_credentials.json.")
        
        http               = credentials.authorize(httplib2.Http())
        service            = discovery.build('calendar', 'v3', http=http)

        start = datetime.utcnow() - timedelta(days = 1)
        end = start + timedelta(days = 14)
        
        eventsResult = service.events().list(
            calendarId=self.support_calendar_id,
            timeMin=start.isoformat() + "Z",
            timeMax=end.isoformat() + "Z",
            singleEvents=True,
            orderBy='startTime').execute()
        events = eventsResult.get('items', [])

        self.calendar = []
        for event in events:
            try:
                start = dateutil.parser.parse(
                    event['start'].get('dateTime',
                                       event['start'].get('date')))
                start = start.astimezone(tz=TZ)
                
                end = dateutil.parser.parse(
                    event['end'].get('dateTime',
                                     event['start'].get('date')))
                end = end.astimezone(tz=TZ)

                self.calendar.append((start, end, self.slack_name_for_full_name(event['summary'])))
            except Exception, e:
                # sometimes there's weird stuff on the calendar that
                # can't be parsed, ignore it
                pass

        self.calendar_refreshed_at = datetime.utcnow()
        return self.calendar

    def _index_slack_names(self):
        # index user IDs by name and real_name to try to match up
        # support shift names
        self.slack_user_names = {}
        for user in self.sc.server.users:
            self.slack_user_names[user.name.lower()] = user.id
            self.slack_user_names[user.real_name.lower()] = user.id

    def slackbot(self):
        while True:
            try:
                self._slackbot()
            except Exception, e:
                print "Caught error from slackbot: %r" % e
                import sys, traceback
                traceback.print_exc()
                print "Pausing and reconnecting after 10 seconds..."
            sleep(10)

    def _slackbot(self):
        self.sc = SlackClient(self.slack_api_key)

        if self.sc.rtm_connect():
            self.slack_connected = True

            # need this so I can scan for messages @me
            slack_bot_user = self.sc.server.users.find(self.slack_bot_name)
            if not slack_bot_user:
                raise Exception("Failed to load user for name %r" % (self.slack_bot_name,))
            self.slack_bot_user_id = slack_bot_user.id

            self._index_slack_names()

            while True:
                msg = self.sc.rtm_read()
                self.slackbot_input(msg)
                self.slackbot_output()
                self.slackbot_autoping()

                if ((datetime.utcnow() - self.last_helpscout_scan) >
                    HELPSCOUT_SCAN_INTERVAL):
                    self.last_helpscout_scan = datetime.utcnow()
                    self.watch(once=True)

                if ((datetime.utcnow() - self.last_calender_scan) >
                    CALENDAR_SCAN_INTERVAL):
                    self.last_calender_scan = datetime.utcnow()
                    self.refresh_support_calendar()

                sleep(1)
        else:
            print "Connection Failed, invalid token?"

    def slackbot_input(self, msgs):
        for msg in msgs:
            self.slackbot_handle(msg)

    def slackbot_unsub(self, user):
        unsub = self.memory['unsub']
        unsub.add(user)
        self.memory['unsub'] = unsub

        return "You are now unsubscribed and will no longer receive alerts.  Respond with 'resub' to undo."

    def slackbot_resub(self, user):
        unsub = self.memory['unsub']
        unsub.discard(user)
        self.memory['unsub'] = unsub

        return "You are now re-subscribed and will receive alerts.  Respond with 'unsub' to undo."

    def slackbot_ignore_ticket(self, num):
        ignore_list = self.memory['ignore_list']
        ignore_list.add(int(num))
        self.memory['ignore_list'] = ignore_list

        return "Cool, I'll stop worrying about %s.  To undo respond 'unignore %s'." % (num, num)

    def slackbot_unignore_ticket(self, num):
        ignore_list = self.memory['ignore_list']
        ignore_list.discard(int(num))
        self.memory['ignore_list'] = ignore_list

        return "Ok, I'll start worrying about %s again.  To undo respond 'ignore %s'." % (num, num)

    def slackbot_snooze_ticket(self, num, time):
        if time is None or time == "":
            time = 10
        else:
            time = int(time)
        snooze = self.memory['snooze']
        snooze[int(num)] = datetime.utcnow() + timedelta(minutes=time)
        self.memory['snooze'] = snooze

        return "Cool, snoozing %s for %d minutes.  To undo respond 'unsnooze %s'." % (num, time, num)

    def slackbot_unsnooze_ticket(self, num):
        snooze = self.memory['snooze']
        del snooze[int(num)]
        self.memory['snooze'] = snooze

        return "Ok, I'll start worrying about %s again.  To undo respond 'snooze %s'." % (num, num)

    def slackbot_handle(self, msg):
        msg_type   = msg.get("type", "")
        text       = msg.get("text", "")
        channel_id = msg.get("channel", "")
        user_id    = msg.get("user", "")

        # print repr(msg) + "\n"
    
        if msg_type == "message":
            # don't listen to yourself talk
            if user_id == self.slack_bot_user_id:
                return

            # if someone mentions a ticket number, post a helpful link
            if re.search(r'\b(?:hs|helpscout)\s?[#]?(\d+)\b', text, re.I):
                match = re.search(r'\b(?:hs|helpscout)\s?[#]?(\d+)\b',
                                  text, re.I)
                ticket_num = match.groups(1)[0];
                self.slackbot_link_hs(msg, ticket_num)

            # if someone mentions a ticket number, post a helpful link
            if re.search(r'\b(?:bug|bugzilla)\s?[#]?(\d+)\b', text, re.I):
                match = re.search(r'\b(?:bug|bugzilla)\s?[#]?(\d+)\b',
                                  text, re.I)
                bug_num = match.groups(1)[0];
                self.slackbot_bugzilla_link(msg, bug_num)

            # only look for the below commands if targeted directly
            if (not re.search(r'\b%s\b' % self.slack_bot_name, text, re.I) and
                not re.search(r'<@%s>' % self.slack_bot_user_id, text, re.I) and
                not channel_id.startswith("D")):
                return

            if re.search(r'\bhelp\b', text, re.I):
                self.slackbot_reply(msg, self.slackbot_help())
                return

            if re.search(r'\bunsub\b', text, re.I):
                self.slackbot_reply(msg, self.slackbot_unsub(user_id))
                return

            if re.search(r'\bresub\b', text, re.I):
                self.slackbot_reply(msg, self.slackbot_resub(user_id))
                return

            if re.search(r'\bignore\s+(\d+)\b', text, re.I):
                match = re.search(r'\ignore\s+(\d+)\b', text, re.I)
                self.slackbot_reply(
                    msg,
                    self.slackbot_ignore_ticket(match.groups(1)[0]))
                return

            if re.search(r'\bunignore\s+(\d+)\b', text, re.I):
                match = re.search(r'\bunignore\s+(\d+)\b', text, re.I)
                self.slackbot_reply(
                    msg,
                    self.slackbot_unignore_ticket(match.groups(1)[0]))
                return

            if re.search(r'\bsnooze\s+(\d+)\b(?:\s+(\d+))?', text, re.I):
                match = re.search(r'\bsnooze\s+(\d+)\b(?:\s+(\d+))?', text, re.I)
                self.slackbot_reply(
                    msg,
                    self.slackbot_snooze_ticket(match.group(1),
                                                match.group(2)))
                return

            if re.search(r'\bunsnooze\s+(\d+)\b', text, re.I):
                match = re.search(r'\bunsnooze\s+(\d+)\b', text, re.I)
                self.slackbot_reply(
                    msg,
                    self.slackbot_unsnooze_ticket(match.groups(1)[0]))
                return

            if re.search(r'\bsupport\b', text, re.I):
                days_since = re.search(r'\b(\w+)\s+days?\b', text, re.I)
                if days_since:
                    days = int(days_since.group(1)) if \
                           re.match(r'\d+', days_since.group(1)) else \
                           text2int(days_since.group(1))
                    if days:
                        self.slackbot_reply(msg, self.support_day(offset=days))
                        return

            if re.search(r'\b(on\s+)?support\b', text, re.I) and \
               re.search(r'\bnow\b', text, re.I):
                self.slackbot_reply(msg, self.support_now())
                return

            if re.search(r'\b(on\s+)?support\b', text, re.I) and \
               re.search(r'\btoday\b', text, re.I):
                self.slackbot_reply(msg, self.support_day())
                return

            if re.search(r'\bsupport\b', text, re.I) and \
               re.search(r'\btomorrow\b', text, re.I):
                self.slackbot_reply(msg, self.support_day(offset=1))
                return

            if re.search(r'\bhelpscout\b', text, re.I) and \
               re.search(r'\bstatus\b', text, re.I):
                self.slackbot_reply(msg, self.helpscout_status())
                return

            if re.search(r'\blouder\b', text, re.I):
                self.slackbot_reply(msg, self.set_user_loudness(user_id, "loud"))
                return

            if re.search(r'\bquieter\b', text, re.I):
                self.slackbot_reply(msg, self.set_user_loudness(user_id, "quiet"))
                return

            if re.search(r'\bhow\s+are\s+you\b', text, re.I) or \
               re.search(r'\bexcuse\b', text, re.I) or \
               re.search(r'\bjoke\b', text, re.I) or \
               re.search(r'\bwhat\'s\s+up\b', text, re.I):
                self.slackbot_reply(msg, self.joke())
                return
    
    def set_user_loudness(self, user_id, setting):
        if 'quiet_users' not in self.memory:
            self.memory.quiet_users = set()
        if setting == "loud":
            if user_id not in self.memory.quiet_users:
                self.memory.quiet_users.add(user_id)
                return "When you're on, you'll be pinged as soon as a ticket comes in. Message me 'quieter' to disable."
            else:
                return "You're already set up to be pinged as soon as a ticket comes in. Message me 'quieter' to disable."
        elif setting == "quiet":
            if user_id in self.memory.quiet_users:
                self.memory.quiet_users.remove(user_id)
                return "You'll no longer be pinged as tickets come in. Message me 'louder' to resume getting those pings."
            else:
                return "You're already set up not to be pinged as tickets come in. Message me 'louder' to resume getting those pings."
        else:
            raise Exception("set_user_loudness needs 'loud' or 'quiet'")

    def joke(self):
        try:
            h = httplib2.Http(".cache")
            (resp_headers, content) = h.request(
                "http://api.icndb.com/jokes/random?firstName=Cal&lastName=Bot",
                "GET")
            data = json.loads(content)
            joke = data['value']['joke']
            joke = re.sub(r'&amp;', '&', joke)
            joke = re.sub(r'&lt;', '<', joke)
            joke = re.sub(r'&gt;', '>', joke)
            joke = re.sub(r'&quot;', '"', joke)
            return joke
        except:
            return "Huh, I got nothin'."

    def slackbot_help(self):
        return """ScoutBot commands:

        support now
        support today
        support tomorrow
        support X days from now

        helpscout status

        ignore XXX   - ignore ticket XXX
        unignore XXX - stop ignoring ticket XXX        

        snooze XXX [YYY] - ignore ticket XXX for YYY minutes or 10 by default
        unsnooze XXX     - stop snoozing ticket XXX        

        unsub    - permanently unsubscribe you from getting annoyed by me
        resub    - go back to getting annoyed by me
        
        louder  - get pinged as tickets arrive on your shift, without delay
        quieter - stop getting pinged as tickets arrive
        """

    def slackbot_link_hs(self, msg, num):
        if num in self.last_hs_link:
            if ((datetime.utcnow() - self.last_hs_link[num])
                < ANNOYANCE_FREQUENCY):
                return

        client = self.client

        # this API wrapper has the wackiest paging system...  Gotta
        # clear it here or subsequent calls will silently find
        # nothing!
        client.clearstate()

        url = None
        subject = None
        for result in helpscout_pager(client.search, query="number:%s" % num):
            if int(result.number) == int(num):
                url = "https://secure.helpscout.net/conversation/%s" % result.id
                subject = result.subject
                break

        if url:
            self.slackbot_reply(msg,
                                "HelpScout [<{url}|#{num}>] - {subject}".format(
                                    num=num,
                                    url=url,
                                    subject=subject))
            self.last_hs_link[num] = datetime.utcnow()

    def slackbot_bugzilla_link(self, msg, num):
        if num in self.last_bugzilla_link:
            if ((datetime.utcnow() - self.last_bugzilla_link[num])
                < ANNOYANCE_FREQUENCY):
                return
        url = self.bugzilla_url + 'show_bug.cgi?id=' + num
        self.slackbot_reply(msg,
                            "Bugzilla [<{url}|#{num}>]".format(
                                num=num,
                                url=url))
        self.last_bugzilla_link[num] = datetime.utcnow()

    def slackbot_reply(self, msg, response):
        self.slack_stack.append((msg['channel'], response))

    def slackbot_broadcast(self, msg):
        for channel in self.slack_channels:
            self.slack_stack.append((channel, msg))

    def slackbot_direct_message(self, user, msg):
        # strip out slack formatting
        user = re.sub(r'[\<\>@]+', '', user)

        unsub = self.memory['unsub']
        if user in unsub:
            self.log("Suppressing send of %s to %s: user is unsubscribed." %
                     (msg, user))
            return

        dm_channel = json.loads(self.sc.server.api_call('im.open', user=user))
        if "channel" in dm_channel:
            self.slack_stack.append((dm_channel['channel']['id'], msg))
        else:
            self.log("Failed to open IM channel to %r: %r" % (user, dm_channel))

    def slackbot_log(self, msg):
        for channel in self.slack_log_channels:
            self.slack_stack.append((channel, msg))

    def slackbot_output(self):
        while len(self.slack_stack):
            msg = self.slack_stack.pop(0)
            channel = self.sc.server.channels.find(msg[0])
            if not channel:
                raise Exception("Could not find channel for msg %r" % (msg))

            # need to do this not send_message to get links to format
            # correctly, oddly enough
            self.sc.server.api_call('chat.postMessage',
                                    channel=channel.id,
                                    text=msg[1],
                                    username=self.slack_bot_name,
                                    as_user=True)

    def slackbot_autoping(self):
        #hardcode the interval to 3 seconds
        now = int(time())
        if now > self.slack_last_ping + 3:
            self.sc.server.ping()
            self.slack_last_ping = now
