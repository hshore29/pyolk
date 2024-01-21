"""Classes corresponding to Outlook.sqlite tables"""

import json
import email
from email.message import EmailMessage
from uuid import UUID
from datetime import datetime, date, timedelta
from dataclasses import dataclass, field

import icalendar
from bs4 import BeautifulSoup

from utils import *
from datafiles import *

def append(olk, data):
    keys = list(data.keys())
    for key in keys:
        if hasattr(olk, key):
            setattr(olk, key, data.pop(key))

def export(olk, path):
    # File name
    if type(olk) is OlkFolder:
        name = '_' + str(olk.RecordID)
    elif type(olk) is OlkMain:
        name = 'Main ' + str(olk.RecordID)
    elif type(olk) is OlkAccountMail:
        name = 'AccountMail ' + str(olk.RecordID)
    elif type(olk) is OlkAccountExchange:
        name = 'AccountExch ' + str(olk.RecordID)
    elif type(olk) is OlkNote:
        name = olk.Title.replace('/', '')[:30].strip()
    else:
        if hasattr(olk, 'Name'):
            name = olk.Name
        else:
            name = str(olk.RecordID)

    # File data
    if hasattr(olk, 'to_file'):
        ext, data = olk.to_file()
    else:
        ext = 'json'
        data = json.dumps(olk.__dict__, default=json_serializer)

    # Write
    path = path + ('/' if path else '') + name + '.' + ext
    with open(path, 'w') as f:
        f.write(data)

def dataField(r=False):
    return field(default=None, init=False, repr=r)


@dataclass
class OlkAction:
    Type: str
    Date: datetime
    RecordID: int


@dataclass
class OlkRecipient:
    Type: str
    Name: str
    Address: str


def get_angle_addr(user):
    return user.Name + ' <' + user.Address + '>'


@dataclass
class OlkMessage:
    # Outlook.sqlite fields
    RecordID: int
    FolderID: int
    AccountUID: int
    ModDate: datetime = field(repr=False)
    MessageType: str = field(repr=False)
    HasAttachment: bool = field(repr=False)
    Hidden: bool = field(repr=False)
    IMAPUID: int = field(repr=False)
    IsOutgoingMessage: bool = field(repr=False)
    MarkedForDelete: bool = field(repr=False)
    MentionedMe: bool = field(repr=False)
    MessageID: int = field(repr=False)
    NormalizedSubject: str = field(repr=False)
    PartiallyDownloaded: bool = field(repr=False)
    DownloadState: int = field(repr=False)
    ReadFlag: bool = field(repr=False)
    RecipientList: str = field(repr=False)
    DisplayTo: str = field(repr=False)
    Preview: str = field(repr=False)
    SenderList: str
    Sent: bool
    Size: int = field(repr=False)
    Status: int = field(repr=False)
    SuppressAutoBackfill: bool = field(repr=False)
    ConversationID: int = field(repr=False)
    ThreadTopic: str = field(repr=False)
    TimeReceived: datetime
    TimeSent: datetime = field(repr=False)
    DueDate: datetime = field(repr=False)
    ExchangeID: str = field(repr=False)
    ExchangeChangeKey: str = field(repr=False)
    FlagStatus: str = field(repr=False)
    Priority: int = field(repr=False)
    HasReminder: bool = field(repr=False)
    InferenceClassification: int = field(repr=False)
    CategoryID: int = field(repr=False)
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    AlarmTrigger: int = dataField()
    CompletedDate: datetime = dataField()
    Reminder: datetime = dataField()
    Sensitivity: str = dataField()
    ScheduledSendDate: datetime = dataField()
    MessageClass: str = dataField()
    Subject: str = dataField(True)
    Headers: str = dataField()
    References: str = dataField()
    Body: str = dataField()
    HTMLBody: str = dataField()
    CardData: str = dataField()
    XML: dict = dataField()
    vCalendar: str = dataField()
    Actions: list = dataField()
    DidReply: bool = dataField()
    DidForward: bool = dataField()
    From: list = dataField()
    To: list = dataField()
    CC: list = dataField()
    BCC: list = dataField()
    MeetingAttendees: list = dataField()
    HasMessageSource: bool = dataField()
    AttachmentMetadata: list = dataField()
    InReplyTo: str = dataField()
    HasDownloadedExternalImages: bool = dataField()
    MessageSize: int = dataField()
    # OwnedBlock attributes
    Attachments: list = field(default_factory=list, init=False, repr=False)
    MessageSource: str = dataField()

    def add_data(self, data):
        # Grab important objects
        self.ActionsTaken = [
            OlkAction(**a) for a in data.pop('ActionsTaken', list())
            ]
        self.From = [OlkRecipient(**a) for a in data.pop('From', list())]
        self.To = [OlkRecipient(**a) for a in data.pop('To', list())]
        self.CC = [OlkRecipient(**a) for a in data.pop('CC', list())]
        self.BCC = [OlkRecipient(**a) for a in data.pop('BCC', list())]
        self.MeetingAttendees = [
            OlkRecipient(**a) for a in data.pop('MeetingAttendees', list())
            ]

        # Copy useful fields
        append(self, data)

        # Store data
        #self.data = data

    def add_blockdata(self, blocks):
        for block in blocks:
            if block['BlockType'] == 'Attc':
                self.Attachments.append(block['FileContents'])
            elif block['BlockType'] == 'MSrc':
                self.MessageSource = block['FileContents']

    def to_file(self):
        msg = EmailMessage()

        msg.add_header('Date', email.utils.format_datetime(self.TimeSent))

        msg.add_header('From', ', '.join(map(get_angle_addr, self.From)))
        msg.add_header('To', ', '.join(map(get_angle_addr, self.To)))
        msg.add_header('Cc', ', '.join(map(get_angle_addr, self.CC)))
        msg.add_header('Bcc', ', '.join(map(get_angle_addr, self.BCC)))

        msg.add_header('Message-ID', self.MessageID)
        msg.add_header('In-Reply-To', self.InReplyTo)
        msg.add_header('References', self.References)

        msg.add_header('Subject', self.Subject)

        if self.Body:
            msg.set_content(self.Body)
            msg.replace_header('Content-Type', 'text/html')
        elif self.Preview:
            msg.set_content(self.Preview)

        # Skipping Attachments and MessageSource since they aren't present
        # in this archive
        return ('eml', msg.as_string())

@dataclass
class OlkAttendee:
    RecipientType: str
    Name: str
    Address: str
    AttendeeType: str


@dataclass
class OlkRRule:
    RecurrenceType: str
    Freq: str = field(repr=False)
    Interval: int = field(repr=False)
    EndType: str
    Occurrences: int = field(repr=False)
    StartDate: date
    Until: date
    RecurrenceDates: list = field(repr=False)
    ExceptionDates: list = field(repr=False)
    Day: str = field(default=None, repr=False)
    MonthDay: int = field(default=None, repr=False)
    SetPos: int = field(default=None, repr=False)


@dataclass
class OlkTimezone:
    TZID: str
    Name: str = field(repr=False)
    MSTZID: int = field(repr=False)
    Standard: list = field(repr=False)
    Daylight: list = field(default=None, repr=False)


def attach_to_event(mime, vevent, i='1'):
    if mime.is_multipart():
        for j, submime in enumerate(mime.get_payload()):
            attach_to_event(submime, vevent, str(i) + str(j))
    else:
        params = {
            'filename': mime.get_filename(),
            #'encoding': 'base64',
            'fmttype': mime.get_content_type(),
            #'x-filename': attach.get_filename(),
            'value': 'BINARY'
            }
        vevent.add('attach', mime.get_payload().strip(), parameters=params)


@dataclass
class OlkEvent:
    # Outlook.sqlite fields
    RecordID: int
    FolderID: int
    AccountUID: int
    ModDate: datetime = field(repr=False)
    StartDateUTC: datetime
    EndDateUTC: datetime
    IsRecurring: bool
    RecurrenceID: int = field(repr=False)
    AttendeeCount: int
    AllowNewTimeProposal: int = field(repr=False)
    UUID: bytes = field(repr=False)
    HasReminder: int = field(repr=False)
    IsRecurring: int = field(repr=False)
    MasterRecordID: int = field(repr=False)
    ExchangeID: str = field(repr=False)
    ExchangeChangeKey: str = field(repr=False)
    CategoryID: int = field(repr=False)
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    StartDateOrganizer: datetime = dataField()
    EndDateOrganizer: datetime = dataField()
    OwnerCriticalChange: datetime = dataField()
    ReplyTime: datetime = dataField()
    Sensitivity: str = dataField()
    Priority: str = dataField()
    MessageClass: str = dataField()
    Subject: str = dataField(True)
    Location: str = dataField(True)
    AllDayEvent: bool = dataField(True)
    DoNotForward: bool = dataField()
    Timezone: OlkTimezone = dataField()
    Body: str = dataField()
    Topic: str = dataField()
    CreationTime: datetime = dataField()
    AlarmTrigger: int = dataField()
    Attendees: list = dataField()
    Organizer: OlkRecipient = dataField()
    BusyStatus: str = dataField(True)
    Response: str = dataField()
    XML: dict = dataField()
    IsMyMeeting: bool = dataField()
    CalendarOwnerName: str = dataField()
    ConferenceHTTPJoinLink: str = dataField()
    OrganizerIsCalendarOwner: str = dataField()
    RRule: OlkRRule = dataField()
    IsCancelled: bool = dataField()
    CanJoinOnline: bool = dataField()
    # OwnedBlock attributes
    Attachments: list = field(default_factory=list, init=False, repr=False)

    def add_data(self, data):
        # Grab important objects
        self.Attendees = [
            OlkAttendee(**a) for a in data.pop('Attendees', list())
            ]
        if 'RRule' in data:
            self.RRule = OlkRRule(**data.pop('RRule'))
        if 'Timezone' in data:
            self.Timezone = OlkTimezone(**data.pop('Timezone'))
        if 'Organizer' in data:
            self.Organizer = OlkRecipient(**data.pop('Organizer'))

        # Copy useful fields
        append(self, data)

        # Store data
        self.data = data

        # Truncate timestamps for all day events
        if self.AllDayEvent:
            self.StartDateUTC = self.StartDateUTC.date()
            self.EndDateUTC = self.EndDateUTC.date()
            self.StartDateOrganizer = self.StartDateOrganizer.date()
            self.EndDateOrganizer = self.EndDateOrganizer.date()

    def add_blockdata(self, blocks):
        # Store data
        for block in blocks:
            if block['BlockType'] == 'ClAt':
                attachment = fix_attachment_encoding(block['FileContents'])
                self.Attachments.append(attachment)

    def to_file(self):
        cal = icalendar.Calendar()
        cal.add('prodid', '-//Microsoft Corporation//Outlook for Mac MIMEDIR//EN')
        cal.add('version', '2.0')

        tz = icalendar.Timezone()
        tz.add('tzid', self.Timezone.TZID)

        for standard in self.Timezone.Standard:
            st = icalendar.TimezoneStandard()
            st.add('dtstart', standard['StartDate'])
            if standard['RRule']:
                st['rrule'] = icalendar.vText(standard['RRule'])
            st['tzoffsetfrom'] = icalendar.vText(standard['OffsetFrom'])
            st['tzoffsetto'] = icalendar.vText(standard['OffsetTo'])
            tz.add_component(st)

        if self.Timezone.Daylight:
            for daylight in self.Timezone.Daylight:
                st = icalendar.TimezoneDaylight()
                st.add('dtstart', daylight['StartDate'])
                if daylight['RRule']:
                    st['rrule'] = icalendar.vText(daylight['RRule'])
                st['tzoffsetfrom'] = icalendar.vText(daylight['OffsetFrom'])
                st['tzoffsetto'] = icalendar.vText(daylight['OffsetTo'])
                tz.add_component(st)
        cal.add_component(tz)

        event = icalendar.Event()
        # uid
        event.add('x-entourage_uuid', str(UUID(bytes=self.UUID)).upper())
        event.add('x-microsoft-exchange-id', self.ExchangeID)
        event.add('x-microsoft-exchange-changekey', self.ExchangeChangeKey)
        
        event.add('dtstamp', self.OwnerCriticalChange)
        event.add('dtstart', self.StartDateOrganizer)
        event.add('dtend', self.EndDateOrganizer)
        event.add('last-modified', self.ModDate)

        if self.Subject:
            event.add('summary', self.Subject)
        body = self.Body.replace('\r\n', '\r').replace('\r', '\r\n')
        plain = BeautifulSoup(body, features='lxml').get_text().strip()
        event.add('description', plain)
        
        if self.Organizer:
            event.add('organizer', 'mailto:' + self.Organizer.Address[:-4],
                      parameters={'cn': self.Organizer.Name})
        #sequence: PidLidAppointmentSequence or 0

        for a in self.Attendees:
            params = {'cn': a.Name, 'rsvp': icalendar.vBoolean(False)}
            address = 'mailto:' + a.Address
            if a.RecipientType == 'Resource':
                params['cutype'] = 'RESOURCE'
                params['role'] = 'NON-PARTICIPANT'
            elif a.RecipientType == 'Optional':
                params['role'] = 'OPT-PARTICIPANT'
            else:
                params['role'] = 'REQ-PARTICIPANT'
            # PARTSTAT & RESPTIME not saved in Olk cache
            event.add('attendee', 'mailto:' + address[:-4], parameters=params)
        #categories: comma-delimited list of PidNameKeywords (category names)
        event.add('class', self.Sensitivity)
        if self.CreationTime:
            event.add('created', self.CreationTime)
        #exdate: recurrence, DeletedInstanceDates
        if self.Location:
            event.add('location', self.Location)
        if self.Priority in ('High', 'HighOverride'):
            event.add('priority', 1)
        elif self.Priority in ('Low', 'LowOverride'):
            event.add('priority', 9)
        else:
            event.add('priority', 5)
        #rdate: recurrence, ModifiedInstanceDates
        #rrule
        #recurrence id
        transp = 'TRANSPARENT' if self.BusyStatus == 'FREE' else 'OPAQUE'
        event.add('transp', transp)

        #x-alt-desc: PidTagRtfCompressed with param FMTTYPE=text/HTML
        event.add('x-microsoft-cdo-busystatus', self.BusyStatus)
        event['x-microsoft-cdo-alldayevent'] = icalendar.vBoolean(self.AllDayEvent)
        #x-microsoft-cdo-importance: from priority?
        #x-microsoft-cdo-ownerapptid: PidTagOwnerAppointmentId
        #x-microsoft-cdo-owner-critical-change: owner critical change time
        event.add('x-microsoft-cdo-replytime', self.ReplyTime)
        event['x-microsoft-disallow-counter'] = icalendar.vBoolean(not self.AllowNewTimeProposal)
        if self.DoNotForward is not None:
            event['x-microsoft-donotforwardmeeting'] = icalendar.vBoolean(self.DoNotForward)
        #x-microsoft-cdo-insttype
        #x-microsoft-exdate
        #x-ms-olk-apptlastsequence: PidLidAppointmentLastSequence
        #x-ms-olk-apptseqtime: PidLidAppointmentSequenceTime
        #x-ms-olk-autostartcheck: PidLidAutoStartCheck
        #x-ms-olk-collaboratedoc: PidLidCollaborateDoc
        #x-ms-olk-confcheck: PidLidConferencingCheck
        #x-ms-olk-conftype: PidLidConferencingType
        #x-ms-olk-directory: PidLidDirectory
        #x-ms-olk-mwsurl: PidLidMeetingWorkspaceUrl
        #x-ms-olk-netshowurl: PidLidNetShowUrl
        #x-ms-olk-onlinepassword: PidLidOnlinePassword
        #x-ms-olk-orgalias: PidLidOrganizerAlias

        if self.HasReminder:
            alarm = icalendar.Alarm()
            trigger = timedelta(minutes=self.AlarmTrigger)
            alarm['trigger'] = icalendar.vDuration(trigger)
            alarm.add('action', 'DISPLAY')
            alarm.add('description', 'Reminder')
            event.add_component(alarm)

        if self.Attachments:
            for i, a in enumerate(self.Attachments):
                attach = email.message_from_string(a)
                attach_to_event(attach, event)

        cal.add_component(event)
        return ('ics', cal.to_ical().decode('utf-8'))


@dataclass
class OlkFolder:
    # Outlook.sqlite fields
    RecordID: int
    ModDate: datetime = field(repr=False)
    AccountUID: int
    ParentID: int
    FolderClass: int = field(repr=False)
    FolderType: int = field(repr=False)
    SpecialFolderType: int = field(repr=False)
    Name: str
    ContainsPartialDwnldMsgs: int = field(repr=False)
    ExchangeID: str = field(repr=False)
    ExchangeChangeKey: str = field(repr=False)
    OnlineFolderType: int = field(repr=False)
    SubFolderSyncMapReset: int = field(repr=False)
    SyncMapReset: int = field(repr=False)
    IgnoreReminders: int = field(repr=False)
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    ItemCount: int = dataField()
    CalendarOwnerAccount: str = dataField()
    CalendarOwnerName: str = dataField()
    CalendarToken: str = dataField()
    GroupID: int = dataField()
    # OwnedBlock attributes
    #SyncMaps: list = field(default_factory=list, init=False, repr=False)

    def add_data(self, data):
        # Copy remaining useful fields
        # Overwrite ModDate, FolderType, FolderClass
        # SpecialFolderType, ContainsPartialDwnldMsgs
        append(self, data)

        # Store data
        #self.data = data

    def add_blockdata(self, blocks):
        # Store data
        for block in blocks:
            if block['BlockType'] in ('ExFS', 'ExSM'):
                pass
                #self.SyncMaps.append(block)


@dataclass
class OlkTask:
    # Outlook.sqlite fields
    RecordID: int
    ModDate: datetime = field(repr=False)
    FolderID: int
    AccountUID: int
    Completed: bool
    DueDate: datetime
    ExchangeID: str = field(repr=False)
    ExchangeChangeKey: str = field(repr=False)
    StartDate: datetime = field(repr=False)
    UUID: bytes = field(repr=False)
    HasReminder: bool = field(repr=False)
    Name: str
    CategoryID: int = field(repr=False)
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    Body: str = field(init=False, repr=False)
    CompletedDate: datetime = field(init=False, repr=False)
    Reminder: datetime = field(init=False, repr=False)

    def add_data(self, data):
        # Copy remaining useful fields
        append(self, data)

        # Store data
        #self.data = data

        # Manipulate dates
        epoch = datetime(2001, 1, 1, 0, 0, tzinfo=ZoneInfo('UTC'))
        if self.DueDate == epoch:
            self.DueDate = None
        if self.StartDate == epoch:
            self.StartDate = None
        if self.DueDate:
            self.DueDate = self.DueDate.astimezone(ZoneInfo('UTC')).date()
        if self.StartDate:
            self.StartDate = self.StartDate.astimezone(ZoneInfo('UTC')).date()

    def add_blockdata(self, blocks):
        pass # No OwnedBlocks present for Tasks in my archive


@dataclass
class OlkNote:
    # Outlook.sqlite fields
    RecordID: int
    ModDate: datetime = field(repr=False)
    FolderID: int
    AccountUID: int
    ExchangeID: str = field(repr=False)
    ExchangeChangeKey: str = field(repr=False)
    UUID: bytes = field(repr=False)
    Title: str
    CategoryID: int = field(repr=False)
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    CreatedDate: datetime = dataField()
    Body: str = dataField()

    def add_data(self, data):
        # Copy remaining useful fields
        append(self, data)

        # Store data
        #self.data = data

    def add_blockdata(self, blocks):
        pass # No OwnedBlocks present for Notes in my archive

    def to_file(self):
        out = '<HTML>\r<HEAD>\r'
        out += "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'/>\r"
        out += '<TITLE>' + self.Title + '</TITLE>\r'
        out += '</HEAD>\r'
        out += self.Body
        out += '\r</HTML>'
        return ('html', out)


@dataclass
class OlkInternetAddress:
    Type: str = field(repr=False)
    Address: str


@dataclass
class OlkContact:
    # Outlook.sqlite fields
    RecordID: int
    ModDate: datetime = field(repr=False)
    FolderID: int
    AccountUID: int = field(repr=False)
    ContactRecType: int = field(repr=False)
    DisplayName: str
    DisplayNameSort: str = field(repr=False)
    LanguageID: int = field(repr=False)
    DueDate: datetime = field(repr=False)
    ExchangeID: str = field(repr=False)
    ExchangeChangeKey: str = field(repr=False)
    FlagStatus: int = field(repr=False)
    StartDate: datetime = field(repr=False)
    UUID: bytes = field(repr=False)
    HasReminder: bool = field(repr=False)
    CategoryID: int = field(repr=False)
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    Sensitivity: str = dataField()
    DefaultEmailAddress: str = dataField(True)
    DefaultIMAddress: str = dataField()
    EmailAddresses: list = dataField()
    IMAddresses: list = dataField()
    FirstName: str = dataField()
    LastName: str = dataField()
    Notes: str = dataField()
    HomeAddressStreet: str = dataField()
    HomeAddressCity: str = dataField()
    HomeAddressState: str = dataField()
    HomeAddressPostalCode: str = dataField()
    HomeAddressCountry: str = dataField()
    PhoneHome: str = dataField()
    PhoneHomeFax: str = dataField()
    WebPageHome: str = dataField()
    PhoneHome2: str = dataField()
    Company: str = dataField(True)
    WorkTitle: str = dataField(True)
    WorkAddressStreet: str = dataField()
    WorkAddressCity: str = dataField()
    WorkAddressState: str = dataField()
    WorkAddressPostalCode: str = dataField()
    WorkAddressCountry: str = dataField()
    Department: str = dataField()
    OfficeLocation: str = dataField()
    PhoneWork: str = dataField()
    PhoneWorkFax: str = dataField()
    PhonePager: str = dataField()
    WebPageWork: str = dataField()
    PhoneMobile: str = dataField()
    PhoneWork2: str = dataField()
    PhonePrimary: str = dataField()
    Alias: str = dataField()
    PhoneAssistant: str = dataField()
    Nickname: str = dataField()
    Title: str = dataField()
    Suffix: str = dataField()
    Custom1: str = dataField()
    Custom2: str = dataField()
    Custom3: str = dataField()
    Custom4: str = dataField()
    Custom5: str = dataField()
    Custom6: str = dataField()
    Custom7: str = dataField()
    Custom8: str = dataField()
    Date1: datetime = dataField()
    Date2: datetime = dataField()
    Birthday: datetime = dataField()
    Anniversairy: datetime = dataField()
    YomiLastName: str = dataField()
    YomiFirstName: str = dataField()
    YomiCompany: str = dataField()
    Phone1: str = dataField()
    Phone2: str = dataField()
    Phone3: str = dataField()
    Phone4: str = dataField()
    MiddleName: str = dataField()
    Spouse: str = dataField()
    Child: str = dataField()
    AstrologicalSign: str = dataField()
    Age: str = dataField()
    BloodType: str = dataField()
    InterestsHobbies: str = dataField()
    Initials: str = dataField()
    HomeAddressFormat: str = dataField()
    WorkAddressFormat: str = dataField()
    PhoneOther: str = dataField()
    PhoneOtherFax: str = dataField()
    PhoneRadio: str = dataField()
    OtherAddressStreet: str = dataField()
    OtherAddressCity: str = dataField()
    OtherAddressState: str = dataField()
    OtherAddressPostalCode: str = dataField()
    OtherAddressCountry: str = dataField()
    OtherAddressFormat: str = dataField()
    JapaneseFormat: bool = dataField()
    PictureFormat: str = dataField()
    # OwnedBlock attributes
    HasPicture: bool = field(default=False, init=False, repr=False)
    PictureImageData: bytes = dataField()

    def add_data(self, data):
        # Normalize useful fields
        self.Date1 = a_b_d_y(data.pop('Date1', None))
        self.Date2 = a_b_d_y(data.pop('Date2', None))
        self.Birthday = a_b_d_y(data.pop('Birthday', None))
        self.Anniversairy = a_b_d_y(data.pop('Anniversairy', None))
        self.EmailAddresses = [
            OlkInternetAddress(**e) for e in data.pop('EmailAddresses', list())
            ]
        self.IMAddresses = [
            OlkInternetAddress(**e) for e in data.pop('IMAddresses', list())
            ]

        # Copy remaining useful fields
        append(self, data)

        # Store data
        #self.data = data

    def add_blockdata(self, blocks):
        # Store data
        for block in blocks:
            if block['BlockType'] == 'ImgB':
                self.HasPicture = True
                self.PictureImageData = block['FileData']


@dataclass
class OlkCategory:
    # Outlook.sqlite fields
    RecordID: int
    AccountUID: int
    Name: str
    IsLocalCategory: bool = field(repr=False)
    ExchangeGuid: str = field(repr=False)
    BackgroundColor: str = field(repr=False)
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    ModDate: datetime = dataField()
    FolderID: int = dataField()
    CreatedDate: datetime = dataField()

    def add_data(self, data):
        # Copy remaining useful fields
        append(self, data)

        # Store data
        #self.data = data

    def add_blockdata(self, blocks):
        pass # No OwnedBlocks present for Categories in my archive


@dataclass
class OlkSignature:
    # Outlook.sqlite fields
    RecordID: int
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    ModDate: datetime = dataField(True)
    Name: str = dataField(True)
    Body: str = dataField()

    def add_data(self, data):
        # Copy remaining useful fields
        append(self, data)

        # Store data
        #self.data = data

    def add_blockdata(self, blocks):
        pass # No OwnedBlocks present for Signatures in my archive


@dataclass
class OlkSavedSearch:
    # Outlook.sqlite fields
    RecordID: int
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    ModDate: datetime = dataField(True)
    Name: str = dataField(True)
    SearchType: str = dataField(True)

    def add_data(self, data):
        # Copy remaining useful fields
        append(self, data)

        # Store data
        #self.data = data

    def add_blockdata(self, blocks):
        pass # No OwnedBlocks present for Saved Searches in my archive


@dataclass
class OlkMain:
    # Outlook.sqlite fields
    RecordID: int
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    ModDate: datetime = dataField()
    # Main #1 fields
    # Main #2 fields
    ExchangeAccountUID: int = dataField()
    MailAccountUID: int = dataField()
    AddressFormats: dict = dataField()
    LocaleIdentifier: str = dataField()
    CalendarWeekStart: str = dataField()
    CalendarDefaultTimezone: int = dataField()
    CalendarWorkDayStarts: int = dataField()
    CalendarWorkDayEnds: int = dataField()
    CalendarWorkWeekSu: bool = dataField()
    CalendarWorkWeekMo: bool = dataField()
    CalendarWorkWeekTu: bool = dataField()
    CalendarWorkWeekWe: bool = dataField()
    CalendarWorkWeekTh: bool = dataField()
    CalendarWorkWeekFr: bool = dataField()
    CalendarWorkWeekSa: bool = dataField()
    DefaultEventReminderUnit: str = dataField()
    DefaultEventReminderBefore: int = dataField()
    DefaultEventReminderEnabled: bool = dataField()
    WorkOffline: bool = dataField()
    SoundSet: str = dataField()
    PlaySoundNewMessage: bool = dataField()
    PlaySoundNoNewMessages: bool = dataField()
    PlaySoundSentMessage: bool = dataField()
    PlaySoundSyncError: bool = dataField()
    PlaySoundWelcome: bool = dataField()
    PlaySoundReminder: bool = dataField()
    NotifyBounceIconInDock: bool = dataField()
    ReplyWithDefaultEmailAccount: bool = dataField()
    AssignMessagesToContactCategories: bool = dataField()
    NotifyDisplayAlert: bool = dataField()
    NotifyShowPreviewInAlert: bool = dataField()
    # OwnedBlock attributes
    RecentAddresses: dict = dataField()

    def add_data(self, data):
        # Copy remaining useful fields
        append(self, data)

        # Store data
        #self.data = data

    def add_blockdata(self, blocks):
        # Store data
        for block in blocks:
            if block['BlockType'] == 'RcnA':
                self.RecentAddresses = block['RecentAddresses']


@dataclass
class OlkAccountMail:
    # Outlook.sqlite fields
    RecordID: int
    AssociatedAccountOfUID: int = field(repr=False)
    Name: str
    EmailAddress: str
    DeviceGuid: str = field(repr=False)
    ServerType: str = field(repr=False)
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    ModDate: datetime = dataField()
    UserName: str = dataField()
    ExchangeServerURL: str = dataField()
    ExchangeServerPort: int = dataField()
    UseSignatureNewMessage: int = dataField()
    UseSignatureReplyForward: int = dataField()
    SigningAlgorithm: str = dataField()
    SignOutgoingMessages: bool = dataField()
    SignIncludeCertificate: bool = dataField()
    SignSendAsClearText: bool = dataField()
    EncryptionAlgorithm: str = dataField()
    OutlookManageURL: str = dataField()
    DownloadHeadersOnly: int = dataField()
    EncryptOutgoingMessages: bool = dataField()

    def add_data(self, data):
        # Copy remaining useful fields
        append(self, data)

        # Store data
        #self.data = data

    def add_blockdata(self, blocks):
        pass # No OwnedBlocks present for Mail Accounts in my archive


@dataclass
class OlkAccountExchange:
    # Outlook.sqlite fields
    RecordID: int
    AssociatedAccountOfUID: int = field(repr=False)
    LDAPAccountUID: int
    MailAccountUID: int
    Name: str = field(repr=False)
    EmailAddress: str
    # DataFile attributes
    #data: dict = dataField()
    BlockType: str = dataField()
    ModDate: datetime = dataField()
    CreatedDate: datetime = dataField()
    UserName: str = dataField()
    DefaultCategory: str = dataField()
    DirectoryServiceMaxResults: int = dataField()
    DirectoryServicePort: int = dataField()
    DirectoryServiceUseSSL: bool = dataField()
    DirectoryServiceUseExchangeCreds: bool = dataField()
    DownloadHeadersOnly: int = dataField()
    ExchangeServerPort: int = dataField()
    ExchangeServerURL: str = dataField()
    SigningAlgorithm: str = dataField()
    SignOutgoingMessages: bool = dataField()
    SignIncludeCertificate: bool = dataField()
    SignSendAsClearText: bool = dataField()
    EncryptionAlgorithm: str = dataField()
    EncryptOutgoingMessages: bool = dataField()
    SyncSharedMailboxes: bool = dataField()
    ReceiptIPAddress: str = dataField()
    OutlookOABURL: str = dataField()
    OutlookManageURL: str = dataField()
    OutlookAPIURL: str = dataField()
    OutlookSearchURL: str = dataField()
    ExchangeEWSURL: str = dataField()
    ServerType: str = dataField()

    def add_data(self, data):
        # Copy remaining useful fields
        append(self, data)

        # Store data
        #self.data = data

    def add_blockdata(self, blocks):
        pass # No OwnedBlocks present for Exchange Accounts in my archive
