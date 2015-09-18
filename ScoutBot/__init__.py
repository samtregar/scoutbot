from warnings import warn
from time import sleep
from datetime import datetime, timedelta
from ConfigParser import SafeConfigParser
import dateutil.parser
import os
from pytz import timezone
import re

# HelpScout API
import helpscout

# Google Calendar API modules
import httplib2
from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

CALENDAR_REFRESH_INTERVAL = timedelta(minutes=5)
TZ = timezone('US/Pacific')

# helper to deal with bizarre helpscout paging interface - you have to
# call each method multiple times until it returns nothing.
def helpscount_pager(meth, *args, **kwargs):
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
    
class ScoutBot:
    def __init__(self, config_file='scoutbot.cfg'):
        config = SafeConfigParser()
        config.read([config_file, os.path.expanduser('~/.' + config_file)])

        self.hs_api_key                 = config.get('helpscout',
                                                     'api_key')
        self.support_domain             = config.get('scoutbot',
                                                     'support_domain')
        self.max_wait_new_ticket        = int(config.get('scoutbot',
                                                         'max_wait_new_ticket'))
        self.max_wait_response_or_close = int(config.get('scoutbot',
                                                         'max_wait_response_or_close'))
        self.support_calendar_id        = config.get('scoutbot',
                                                     'support_calendar_id')

        if self.hs_api_key is None:
            raise Exception("Missing api_key config value!")

        self.client = helpscout.Client()
        self.client.api_key = self.hs_api_key

        self.calendar = []
        self.calendar_refreshed_at = None

    def open_conversations(self, hours=6, status='active'):
        client = self.client

        # this API wrapper has the wackiest paging system...  Gotta
        # clear it here or subsequent calls will silently find
        # nothing!
        client.clearstate()
        
        results = []
        for mailbox in helpscount_pager(client.mailboxes):
            # look back up to 6 hours by default
            start_date = datetime.utcnow() - timedelta(hours = hours)
            start_date = start_date.replace(microsecond=0).isoformat() + 'Z'
            
            for conv in helpscount_pager(client.conversations_for_mailbox,
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
        else:
            data['needs_reply_or_close'] = last_client_msg_at > last_support_msg_at
        data['wait_time'] = datetime.utcnow() - data['last_client_msg_at']
        
        return data

    def watch(self):
        while True:
            self.log("*** Scanning for conversations...")
            tickets = self.open_conversations()
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

            sleep(10)

    def alert_support(self, ticket):
        self.log("+++ CALLING FOR HELP!!! +++")

    def alert_everyone(self, ticket):
        self.log("+++ CALLING FOR HELP FROM EVERYONE!!! +++")

    def log(self, msg):
        print "%s: %s" % (datetime.now(), msg)

    def support_now(self):
        cal = self.refresh_support_calendar()
        now = datetime.now(tz=TZ)
        for c in cal:
            if now >= c[0] and now <= c[1]:
                return "%s is on support now." % (c[2],)
        return "Nobody is on support now!"

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
