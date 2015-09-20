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

CALENDAR_REFRESH_INTERVAL = timedelta(minutes=5)
TZ = timezone('US/Pacific')

HELPSCOUT_SCAN_INTERVAL = timedelta(minutes=1)
HELPSCOUT_TIMEOUT = 5

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
        items = meth(*args, **kwargs)
        if items is None:
            return
        for item in items:
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
        self.slack_api_key              = config.get('slack', 'api_key')
        self.slack_bot_name             = config.get('slack', 'bot_name')
        self.slack_last_ping            = 0
        self.slack_stack                = []
        self.slack_connected            = False

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

        # is this too rude?  Maybe weird if ScoutBot gets used by
        # other code...
        def signal_handler(signal, frame):
            print("\nExiting SlackBot.  Thank you for playing.\n")
            sys.exit(0)
        signal.signal(signal.SIGINT, signal_handler)

    def open_conversations(self, hours=6, status='active'):
        client = self.client

        # this API wrapper has the wackiest paging system...  Gotta
        # clear it here or subsequent calls will silently find
        # nothing!
        client.clearstate()
        
        results = []
        for mailbox in helpscout_pager(client.mailboxes):
            # look back up to 6 hours by default
            start_date = datetime.utcnow() - timedelta(hours = hours)
            start_date = start_date.replace(microsecond=0).isoformat() + 'Z'
            
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
        for thread in threads:
            created_at = dateutil.parser.parse(thread['createdAt'])\
                         .replace(tzinfo=None)

            email = thread['createdBy']['email']
            if email.endswith(self.support_domain):
                if (last_support_msg_at is None or \
                    last_support_msg_at < created_at):
                    last_support_msg_at = created_at
                    last_owner_email = last_owner_email
            else:
                if (last_client_msg_at is None or \
                    last_client_msg_at < created_at):
                    last_client_msg_at = created_at

        data['new'] = last_support_msg_at is None
        data['last_support_msg_at'] = last_support_msg_at
        data['last_client_msg_at'] = last_client_msg_at
        if data['new']:
            data['needs_reply_or_close'] = True
        elif data['last_client_msg_at'] is None:
            # we're talking to ourselves now, great
            data['needs_reply_or_close'] = False
        else:
            data['needs_reply_or_close'] = last_client_msg_at > last_support_msg_at
        data['wait_time'] = datetime.utcnow() - (data['last_client_msg_at'] if data['last_client_msg_at'] else datetime.utcnow())
        
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

        summary = []
        for ticket in self.helpscout_current_tickets:
            if ticket['new']:
                summary.append("[%s] %s => new and unclaimed %s." % \
                               (ticket['num'],
                                ticket['subject'],
                                td_format(ticket['wait_time'])))
            elif ticket['needs_reply_or_close']:
                summary.append("[%s] %s => needs response or close %s." % \
                               (ticket['num'],
                                ticket['subject'],
                                td_format(ticket['wait_time'])))
            else:
                summary.append("[%s] %s => handled." % \
                         (ticket['num'],
                          ticket['subject']))
        return "\n".join(summary)
    
    def scan_conversations(self):
        self.log("*** Scanning for conversations...")
        tickets = timeout(lambda: self.open_conversations(),
                          timeout_duration=HELPSCOUT_TIMEOUT,
                          default=None)
        if tickets is None:
            self.log("*** Timed out looking for conversation...")
            return
        self.helpscout_current_tickets = tickets

        for ticket in tickets:
            if ticket['new']:
                self.log("*** [%s] %s => new and unclaimed %s" % \
                         (ticket['num'],
                          ticket['subject'],
                          td_format(ticket['wait_time'])))
                if ticket['wait_time'].total_seconds() > self.max_wait_new_ticket:
                    self.alert_support(ticket)

                if ticket['wait_time'].total_seconds() > (self.max_wait_new_ticket * 2):
                    self.alert_everyone(ticket)

            elif ticket['needs_reply_or_close']:
                self.log("*** [%s] %s => needs response of close %s" % \
                         (ticket['num'],
                        ticket['subject'],
                          td_format(ticket['wait_time'])))
                if ticket['wait_time'].total_seconds() > self.max_wait_response_or_close:
                    self.alert_support(ticket)
                if ticket['wait_time'].total_seconds() > (self.max_wait_response_or_close * 2):
                    self.alert_everyone(ticket)

            else:
                self.log("+ [%s] %s => handled" % \
                         (ticket['num'],
                          ticket['subject']))


    def alert_support(self, ticket):
        self.log("+++ CALLING FOR HELP!!! +++")

    def alert_everyone(self, ticket):
        self.log("+++ CALLING FOR HELP FROM EVERYONE!!! +++")

    def log(self, msg):
        if self.slack_connected:
            self.slackbot_log(msg)
        print "%s: %s" % (datetime.now(), msg)

    def support_now(self):
        cal = self.refresh_support_calendar()
        now = datetime.now(tz=TZ)
        for c in cal:
            if now >= c[0] and now <= c[1]:
                return "%s is on support now." % (c[2],)
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

    # pull a fresh calendar from Google periodically
    def refresh_support_calendar(self, use_cache=True):
        if (use_cache and
            len(self.calendar) and
            self.calendar_refreshed_at and
            (datetime.utcnow() - self.calendar_refreshed_at) >
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
            start = dateutil.parser.parse(
                event['start'].get('dateTime',
                                   event['start'].get('date')))
            start = start.astimezone(tz=TZ)
            end = dateutil.parser.parse(
                event['end'].get('dateTime',
                                 event['start'].get('date')))
            end = end.astimezone(tz=TZ)

            self.calendar.append((start, end, event['summary']))
        self.calendar_refreshed_at = datetime.utcnow()
        return self.calendar

    def slackbot(self):
        self.sc = SlackClient(self.slack_api_key)

        if self.sc.rtm_connect():
            self.slack_connected = True

            # need this so I can scan for messages @me
            self.slack_bot_user_id = self.sc.server.users.find(
                self.slack_bot_name).id

            while True:
                msg = self.sc.rtm_read()
                self.slackbot_input(msg)
                self.slackbot_output()
                self.slackbot_autoping()

                if ((datetime.utcnow() - self.last_helpscout_scan) >
                    HELPSCOUT_SCAN_INTERVAL):
                    self.last_helpscout_scan = datetime.utcnow()
                    self.watch(once=True)
                sleep(1)
        else:
            print "Connection Failed, invalid token?"

    def slackbot_input(self, msgs):
        for msg in msgs:
            self.slackbot_handle(msg)

    def slackbot_handle(self, msg):
        msg_type = msg.get('type', "")
        text = msg.get('text', "")

        if msg_type == "message":
            if (not re.search(r'\b%s\b' % self.slack_bot_name, text, re.I) and
                not re.search(r'<@%s>' % self.slack_bot_user_id, text, re.I)):
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

            if re.search(r'\bon\s+support\b', text, re.I) and \
               re.search(r'\bnow\b', text, re.I):
                self.slackbot_reply(msg, self.support_now())
                return

            if re.search(r'\bsupport\b', text, re.I) and \
               re.search(r'\btomorrow\b', text, re.I):
                self.slackbot_reply(msg, self.support_day(offset=1))
                return

            if re.search(r'\bhelpscout\b', text, re.I) and \
               re.search(r'\bstatus\b', text, re.I):
                self.slackbot_reply(msg, self.helpscout_status())
                return

    def slackbot_reply(self, msg, response):
        self.slack_stack.append((msg['channel'], response))

    def slackbot_broadcast(self, msg):
        for channel in self.slack_channels:
            self.slack_stack.append((channel, msg))

    def slackbot_log(self, msg):
        for channel in self.slack_log_channels:
            self.slack_stack.append((channel, msg))

    def slackbot_output(self):
        while len(self.slack_stack):
            msg = self.slack_stack.pop(0)
            channel = self.sc.server.channels.find(msg[0])
            if not channel:
                raise Exception("Could not find channel for msg %r" % (msg))
            channel.send_message(msg[1])

    def slackbot_autoping(self):
        #hardcode the interval to 3 seconds
        now = int(time())
        if now > self.slack_last_ping + 3:
            self.sc.server.ping()
            self.slack_last_ping = now
