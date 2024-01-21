"""Handle reading of binary olk* data files"""

from collections import defaultdict
from struct import unpack, error

from utils import *

## Enums
OlRepeats = {8202: 'Daily', 8203: 'Weekly', 8204: 'Monthly', 8205: 'Yearly'}
OlRecurrenceEndType = {
    8225: 'ByDate',
    8226: 'AfterCount',
    8227: 'None'
    }
OlRecurrenceType = {
    0: 'Daily', 1: 'Weekly', 2: 'Monthly', 3: 'MonthNth',
    5: 'Yearly', 6: 'YearNth'
    }
OlBusyStatus = {
    0: 'BUSY', 1: 'FREE', 2: 'TENTATIVE', 3: 'OOF'
    }
RESPONSE = {0: 'None', 1: 'Accepted', 2: 'Tentative'}
# Off by 1 from OlMeetingRecipientType, maybe because 'Organizer' is separated?
OlRecipientType = {
    0: 'Required', 1: 'Optional', 2: 'Resource'
    }
OlSearchType = {
    1: 'Mail', 2: 'Contact', 4: 'Task', 5: 'Note'
    }
OlFolderClass = {
    0: 'Mail', 1: 'Contact', 2: 'Event', 4: 'Note', 5: 'Task', 7: 'Group'
    }
OlSensitivity = {
    0: 'PUBLIC',
    1: 'X-PERSONAL',
    2: 'PRIVATE',
    3: 'CONFIDENTIAL'
    }
OlPriority = {
    1: 'High',
    2: 'HighOverride',
    3: 'Normal',
    4: 'LowOverride',
    5: 'Low'
    }
OlAddressPart = {
    2: 'Street',
    3: 'City',
    4: 'State',
    5: 'ZIP',
    6: 'Country'
    }
OlTimeUnit = {1: 'Minutes', 2: 'Hours', 3: 'Days'}
OlDayOfWeek = {1: 'SU', 2: 'MO', 3: 'TU', 4: 'WE', 5: 'TH', 6: 'FR', 7: 'SA'}
OlOrganizerType = {0: 'Other', 128: 'CalendarOwner'}
OlAction = {
    2: 'Reply',
    3: 'Forward',
    11: 'ReplyAll'
    }
OlUserType = {
    1: 'DistributionList',
    2: 'User',
    3: None,
    4: 'AttendeeMe',
    8: 'AttendeeUnknown'
    }
OlFlagStatus = {
    0: None,
    1: 'Flagged',
    2: 'Completed'
    }
OlAttendeeType = {
    0: 'User',
    2: 'Contact Group'
    }

# https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-lcid/a9eac961-e77d-41a6-90a5-ce1a8b0cdb9c
LOCALE = {0: None, 1033: 'en-US'}
# https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.package.variant.varianttype?view=visualstudiosdk-2019
VariantType = {
    '00': 'VT_EMPTY', # An empty variant
    '01': 'VT_NULL', # Null value
    '02': 'VT_I2', # Integer (2 bytes signed)
    '03': 'VT_I4', # Integer (4 bytes signed)
    '04': 'VT_R4', # Float (4 bytes)
    '05': 'VT_R8', # Double (8 bytes)
    '06': 'VT_CY', # Currency (fixed decimal point stored in 64-bits)
    '07': 'VT_DATE', # Date
    '08': 'VT_BSTR', # String
    '09': 'VT_DISPATCH', # Object that implements the IDispatch interface
    '0A': 'VT_ERROR', # An error code
    '0B': 'VT_BOOL', # Boolean
    '0C': 'VT_VARIANT', # Reference to a variant object
    '0D': 'vbDataObject', # Visual Basic Data access object (VT_UNKNOWN - IUnknown interface)
    '0E': 'VT_DECIMAL', # Decimal
    '10': 'VT_I1', # Integer (1 byte signed)
    '11': 'VT_UI1', # Integer (1 byte unsigned)
    '12': 'VT_UI2', # Integer (2 byte unsigned)
    '13': 'VT_UI4', # Integer (4 byte unsigned)
    '14': 'VT_I8', # Integer (8 bytes signed)
    '15': 'VT_UI8', # Integer (8 bytes unsigned)
    '16': 'VT_INT', # Integer (generic signed, typically 4 bytes)
    '17': 'VT_UINT', # Integer (generic unsigned, typically 4 bytes)
    '18': 'VT_VOID', # C-style void type
    '19': 'VT_HRESULT', # An HRESULT or COM return value
    '1A': 'VT_PTR', # Generic pointer
    '1B': 'VT_SAFEARRAY', # Array that is guaranteed to be at least empty (non-null)
    '1C': 'VT_CARRAY', # C-style array (an array of pointers)
    '1D': 'VT_USERDEFINED', # User-defined blob
    '1E': 'VT_LPSTR', # Null-terminated ANSI string
    '1F': 'VT_LPWSTR', # Null-terminated wide character (Unicode) string
    '20': '???',
    '24': 'VT_RECORD', # Variants that contained user-defined types
    '25': 'VT_INT_PTR', # Signed Integer pointer
    '26': 'VT_UINT_PTR', # Unsigned Integer pointer
    '40': 'VT_FILETIME', # A FILETIME value
    '41': 'VT_BLOB', # An arbitrary block of memory
    '42': 'VT_STREAM', # A stream of bytes
    '43': 'VT_STORAGE', # Name of the storage
    '45': 'VT_STORED_OBJECT', # A storage object
    '44': 'VT_STREAMED_OBJECT', # A stream that contains an object
    '46': 'VT_BLOB_OBJECT', # A blob representing an object
    '47': 'VT_CF', # A value specifying a clipboard format
    '48': 'VT_CLSID', # A GUID for a class (a CLSID)
    '4D': 'MacAbsoluteTime', # Apple Mac Absolute timestamp (double)
    '0010': 'VT_VECTOR', # An array with a leading count value
    '0020': 'VT_ARRAY', # An array of variants
    '0040': 'VT_BYREF', # a reference to an object
    '0080': 'VT_RESERVED', # Reserved for future use
    'F0FF': 'VT_ILLEGALMASKED', # A bit mask to isloate valid variant types
    'FFFF': 'VT_ILLEGAL', # An illegal variant type
    }

# Initialize schemas
# Dictionary with up to three attributes:
#  override: format dict entries (see OlkDataClass) to swap in
#  remap: just change a property name
#  skip_null: properties that are always blank or null or zero
#  skip_dupe: properties that are the same as other properties
OlkMain = {'class': 'OlkMain',
    'override': {'03:03': ('BlockID', 'rX', None)},
    'skip_indb': ['BlockID']
    }
OlkFolder = {'class': 'OlkFolder',
    'skip_null': ['bool5F01', 'bool6001'],
    'skip_indb': [
        'FolderID', 'AccountUID', 'ExchangeID', 'ExchangeChangeKey',
        'Name', 'OnlineFolderType', 'SyncMapBlockID', 'FolderSyncBlockID'
        ]
    }
OlkMessage = {'class': 'OlkMessage',
    'override': {
        '02:01': ('HasMessageSource', 'fF', lambda h: h == 1),
        '03:04': ('MessageType', 'rF', ol_type_code),
        '03:07': ('MessageSourceBlockID', 'rX', None),
        '03:2B': ('int2B', 'rF', lambda b: unpack('<q', b)[0]),
        '03:14': ('int14', 'fX', None),
        },
    'remap': {
        '03:1A': 'DownloadState',
        '1E:04': 'Headers',
        '1F:01': 'Subject',
        '1F:1E': 'Body',
        '1F:23': 'RecipientList',
        '1F:6A': 'CardData'
        },
    'skip_null': [],
    'skip_dupe': [
        'From2', 'From3', 'ThreadTopic2', 'References2', 'References3',
        'Reminder2', 'HasAttachmentOrInline', 'Sent2', 'HasvCalendar'
        ],
    'skip_indb': [
        'DownloadState', 'ConversationID', 'FolderID', 'AccountUID', 'Sent',
        'ExchangeID', 'ExchangeChangeKey', 'TimeReceived', 'Priority', 'Read',
        'ThreadTopic', 'MessageID', 'Preview', 'HasAttachment', 'HasReminder',
        'PartiallyDownloaded', 'RecipientList', 'MentionedMe',
        'SuppressAutobackfill', 'MessageSourceBlockID', 'MsrcBlockStruct'
        ]
    }
OlkContact = {'class': 'OlkContact',
    'remap': {
        '1F:01': 'FirstName',
        '1F:02': 'LastName',
        '1F:04': 'Notes',
        '1F:08': 'HomeAddressState',
        '1F:09': 'HomeAddressPostalCode',
        '1F:0A': 'HomeAddressCountry',
        '1F:0B': 'PhoneHome',
        '1F:0C': 'PhoneHomeFax',
        '1F:0E': 'WebPageHome',
        '1F:1E': 'PhoneWorkFax',
        '1F:23': 'PhonePrimary',
        '1F:5A': 'Phone1',
        '1F:5B': 'Phone2',
        '1F:5C': 'Phone3',
        '1F:5D': 'Phone4',
        },
    'skip_null': [],
    'skip_indb': [
        'FolderID', 'AccountUID', 'ExchangeID', 'ExchangeChangeKey',
        'UUID', 'HasReminder', 'PictureBlockID'
        ]
    }
OlkAccountExchange = {'class': 'OlkAccountExchange',
    'skip_dupe': ['EmailAddressUnicode', 'EmailAddress2'],
    'skip_indb': [
        'MailAccountUID', 'DisplayName', 'EmailAddress',
        'LDAPAccountUID'
        ]
    }
OlkNote = {'class': 'OlkNote',
    'skip_indb': [
        'FolderID', 'AccountUID', 'ExchangeID', 'ExchangeChangeKey',
        'UUID', 'ModDate', 'Title',
        ]
    }
OlkTask = {'class': 'OlkTask',
    'remap': {'1F:0B': 'Body'},
    'skip_indb': [
        'FolderID', 'AccountUID', 'ExchangeID', 'ExchangeChangeKey',
        'UUID', 'ModDate', 'Name', 'StartDate', 'DueDate',
        'Completed', 'HasReminder'
        ]
    }
OlkEvent = {'class': 'OlkEvent',
    'remap': {
        '03:1A': 'MasterRecordID',
        '1E:04': 'CalendarUID',
        '1F:01': 'Body',
        '1F:02': 'Subject',
        '1F:04': 'Location',
        '1F:08': 'Conference',
        '1F:09': 'ConferenceJoinLink',
        '1F:0A': 'ConferenceHTTPJoinLink',
        '1F:0B': 'ConferenceCapabilities',
        '1F:0C': 'ConferenceInBand',
        },
    'override': {
        '03:03': ('OrganizerIsCalendarOwner', 'fE', OlOrganizerType),
        '03:0E': ('NextReminderTime', 'fF', dt_winminutes),
        },
    'skip_null': [
        # Maybe not null, but not useful
        'DismissTime', 'DownloadDate', 'MessageSize',
        'Overdue', 'AttachmentExchangeID', 'AttachmentBlockID',
        # Null or zero
        'bool0E', 'bool13', 'bool18'
        ],
    'skip_dupe': [
        'ReplyTo', 'DownloadDate2', 'Address', 'Timezone2'
        ],
    'skip_indb': [
        'MasterRecordID', 'RecurrenceID', 'AttendeeCount',
        'FolderID', 'AccountUID', 'ExchangeID', 'ExchangeChangeKey',
        'UUID', 'ModDate', 'CalendarUID', 'StartDateUTC',
        'EndDateUTC', 'IsRecurring', 'AllowNewTimeProposal'
        ]
    }
OlkCategory = {'class': 'OlkCategory',
    'skip_null': [
        'short3201', 'date3501', 'date3601', 'date3701', 'date3801',
        'date3901'
        ],
    'skip_indb': [
        'AccountUID', 'ExchangeGUID', 'Name', 'IsLocalCategory'
        ]
    }
OlkAccountMail = {'class': 'OlkAccountMail',
    'skip_dupe': ['EmailAddressUnicode'],
    'skip_indb': [
        'ExchangeAccountUID', 'EmailAddress', 'DisplayName'
        ]
    }
OlkSavedSearch = {'class': 'OlkSavedSearch',
    'override': {
        '03:04': ('SearchType', 'fE', OlSearchType),
        '03:06': ('int06', 'rF', lambda b: unpack('<b', b)[0])
        },
    'remap': {'1F:01': 'Name'},
    'skip_null': ['int02', 'int06', 'int09', 'int0A', 'long01']
    }
OlkSignature = {'class': 'OlkSignature'}
OlkRecurrence = {'class': 'OlkRecurrence',
    'override': {
        '02:01': ('Repeats', 'fE', OlRepeats),
        '03:01': ('RecurrenceType', 'fE', OlRecurrenceType),
        '03:03': ('EndType', 'fE', OlRecurrenceEndType),
        '03:07': ('WeekDay', 'fF', lambda i: ','.join(ol_days_of_week(i))),
        '03:09': ('MonthDOW', 'fF', lambda i: ','.join(ol_days_of_week(i))),
        '03:10': ('Until', 'fF', dt_winminutes),
        '0D:01': ('RecurrenceDates', 'fF', parse_date_list),
        '0D:02': ('ExceptionDates', 'fF', parse_date_list)
        },
    'remap': {
        '02:01': 'Freq', '03:02': 'Interval', '03:04': 'Occurrences',
        '03:0A': 'MonthNth'
        },
    'skip_null': ['MessageSize'],
    'skip_dupe': ['AlarmTrigger']
    }
OlkAttendee = {'class': 'OlkAttendee',
    'override': {
        '03:01': ('RecipientType', 'fE', OlRecipientType),
        '03:02': ('AttendeeType', 'fE', OlAttendeeType),
        },
    'remap': {'0B:02': 'bool02', '0B:03': 'bool03'},
    'skip_null': ['bool02', 'bool03', 'bool04']
    }
OlkTimezone = {'class': 'OlkTimezone', 'skip_dupe': ['TZLongName']}
OlkTZProp = {'class': 'OlkTZProp'}
OlkAttachment = {'class': 'OlkAttachment',
    'override': {
        '03:4C01': ('int4C01', 'rF', lambda b: unpack('<q', b)[0])
        },
    'skip_indb': ['AttachmentBlockID'],
    'skip_dupe': ['FileNameUnicode']
    }
OlkContentType = {'class': 'OlkContentType',
    'override': {
        '03:04': ('x-mac-creator', 'rF', ol_type_code),
        '03:05': ('x-mac-type', 'rF', ol_type_code),
        },
    'remap': {
        '02:01': 'ContentTypeId',
        '02:02': 'ContentSubtypeId',
        '03:01': 'StartPos',
        '03:02': 'HeaderEndPos',
        '03:03': 'BodyEndPos',
        '1E:01': 'ContentType',
        '1E:03': 'Charset',
        '1E:04': 'ContentID',
        '1F:01': 'FileName',
        '1F:02': 'FileNameUnicode'
        },
    'skip_dupe': ['FileNameUnicode', 'ContentTypeId', 'ContentSubtypeId']
    }
OlkMultipartType = {'class': 'OlkMultipartType',
    'override': {
        '0D:01': ('Parts', 'fL', OlkContentType),
        },
    'remap': {
        '02:01': 'ContentTypeId',
        '02:02': 'ContentSubtypeId',
        '03:01': 'StartPos',
        '03:02': 'HeaderEndPos',
        '03:03': 'BodyEndPos',
        '1E:01': 'ContentType',
        '1E:02': 'Boundary'
        },
    'skip_dupe': ['ContentTypeId', 'ContentSubtypeId']
    }
OlkAddressFormat = {'class': 'OlkMainCountry',
    'override': {
        '03:01': ('part_1', 'fE', OlAddressPart),
        # Sep = Line break
        '03:02': ('part_2', 'fE', OlAddressPart),
        # Sep = Line break
        '03:05': ('part_5', 'fE', OlAddressPart),
        # Sep = unicode05
        '03:06': ('part_6', 'fE', OlAddressPart),
        # Sep = Space
        '03:07': ('part_7', 'fE', OlAddressPart),
        # Sep = Line break
        '03:09': ('part_9', 'fE', OlAddressPart),
        # Sep = unicode08
        '03:0A': ('part_A', 'fE', OlAddressPart),
        # Sep = Line break
        '03:0D': ('part_D', 'fE', OlAddressPart),
        '03:14': ('int14', 'fX', None),
        },
    'remap': {
        '0B:02': 'bool02', # Always true
        '0B:03': 'bool03', # Always true except for AU
        '1F:01': 'country_code',
        '1F:02': 'sep_street', # Street / House Num separator
        '1F:05': 'sep_5_6', # Separator for items 5 and 6
        '1F:08': 'sep_9_A', # Separator for items 9 and A
        },
    'skip_null': ['bool02', 'bool03']
    }
OlkActionsTaken = {'class': 'OlkActionsTaken'}

# Class ID -> Schema lookup
CLASSTOSCHEMA = {
    1: OlkMain,
    2: OlkFolder,
    3: OlkMessage,
    4: OlkContact,
    5: OlkAccountExchange,
    6: OlkNote,
    7: OlkTask,
    8: OlkEvent,
    9: OlkCategory,
    14: OlkAccountMail,
    19: OlkSavedSearch,
    21: OlkSignature
    }


class OlkDataFile:
    """Class for parsing Olk binary data files"""

    def __init__(self, path):
        self.skip_indb = ['RecordID', 'ItemID']
        self._initialize_format_dicts()
        self.path = path

        # read and parse datafile
        self.parts = self._parse(open(path, 'rb'))

    def data(self):
        return {k: v for k, v in self.parts.items() if k not in self.skip_indb}

    def _initialize_format_dicts(self):
        # Initialize format dict for each section
        #  schema is ([name], [handler_mode], [handler])
        # sub-section format dicts are first, primary format dict is last
        # first part of the key is (usually) a .Net Variant Type enum
        # second part is an index unique within the variant type
        self.OLKDATAFILE = {
            # integers (2 bytes signed)
            '02:01': ('short01',), # Message; Attachment Details; Recurrence Freq
            '02:02': ('short02',), # Attachment Details, ?
            '02:03': ('short03',), # Attachment Details, ?
            '02:04': ('short04',), # Attachment Details, ?
            '02:06': ('short06',), # Message
            '02:65': ('DefaultEmailRaw', 'rX', None), # Contact
            '02:77': ('DefaultIMRaw', 'rX', None), # Contact
            '02:80': ('Sensitivity', 'fE', OlSensitivity), # Message, Contact
            '02:81': ('Priority', 'fE', OlPriority), # Message, Event
            '02:82': ('short82',), # Event, Task
            '02:D4': ('shortD4', 'rX', None), # Contact
            '02:2C01': ('DownloadHeadersOnly',), # Accounts (both)
            '02:2D01': ('SpecialFolderType',), # Folder = {1-10,12,14,99,103,106}
            '02:2F01': ('CalendarWeekStart', 'fE', OlDayOfWeek), # Main
            '02:3001': ('DefaultEventReminderUnit', 'fE', OlTimeUnit), # Main
            '02:3101': ('LocaleIdentifier', 'fE', LOCALE), # Main
            '02:3201': ('short3201', 'rX', None), # Category, always 0
            '02:3301': ('OnlineFolderType',), # Folder = {None, 1}
            '02:3901': ('shortCalendar1',), # Folder
            '02:3A01': ('shortCalendar2',), # Folder

            # integers (4 bytes signed)
            '03:00': ('RecordID',),
            '03:01': ('int01',), # Attendee Type, Recurrence Type
            '03:02': ('int02',), # Search, always 1; Attendee
            '03:03': ('int03',), # Main, Event My Meeting, Msg 528 vs. 36
            '03:04': ('int04',), # Main=0/1, SearchType, MessageType
            '03:05': ('MessageSize',), # Message, Event (always 0 for Event)
            '03:06': ('AlarmTrigger',), # Message, Event (always 0 for Msg)
            '03:07': ('int07',), # Event, always 0, deprecated in v16
            '03:08': ('MonthDay',), # Recurrence
            '03:09': ('int09',), # Search, always 2
            '03:0A': ('int0A',), # Search, always 0
            '03:0C': ('Response', 'fE', RESPONSE),
            '03:0D': ('int0D',), # Main, 107 or null
            '03:0E': ('int0E',), # Main, 105 or null; Event next reminder
            '03:0F': ('StartDate', 'fF', dt_winminutes), # Recurrence
            '03:10': ('int10',), # Message, Recurrence Until
            '03:13': ('StartDateUTC', 'fF', dt_winminutes), # Event
            '03:14': ('EndDateUTC', 'fF', dt_winminutes), # Event, Message (something else)
            '03:15': ('int15',), # Message, 651 or null
            '03:16': ('int16',), # Event, null or in the 1-400 range
            '03:17': ('StartDateOrganizer', 'fF', dt_winminutes), # Event
            '03:18': ('EndDateOrganizer', 'fF', dt_winminutes), # Event
            '03:1A': ('int1A',), # Event: MasterRecordID; Message: DownloadState
            '03:1D': ('BusyStatus', 'fE', OlBusyStatus), # Event
            '03:1E': ('RecurrenceID',), # Event
            '03:20': ('AttendeeCount',), # Event
            '03:23': ('int23',), # Message
            '03:24': ('int24',), # Event, null or 0, 1, 2, 3, deprecated in v16
            '03:27': ('int27',), # Message
            '03:29': ('ConversationID', 'rF', lambda b: unpack('<q', b)[0]),
            '03:2A': ('int2A',), # Message, null or 0, 1
            '03:2B': ('int2B',), # Message
            '03:35': ('int35',), # Contact, always 0
            '03:64': ('EmailCount',), # Contact
            '03:76': ('IMCount',), # Contact
            '03:80': ('intCalendar3',), # Folder
            '03:94': ('int94',), # Contact
            '03:9E': ('int9E',), # Contact, always 0
            '03:E3': ('FlagStatus', 'fE', OlFlagStatus), # Message
            '03:E4': ('EmailTypesRaw',), # Contact
            '03:E5': ('IMTypesRaw',), # Contact
            '03:2C01': ('ServerType', 'rF', ol_type_code), # Account (Mail)
            '03:2E01': ('UseSignatureNewMessage',), # Account (Mail)
            '03:2F01': ('UseSignatureReplyForward',), # Account (Mail)
            '03:3001': ('int3001',), # Account (Exchange), 30
            '03:3201': ('DirectoryServiceMaxResults',), # Account (Exchange)
            '03:3701': ('int3701',), # Account (Exchange), 20
            '03:3801': ('ExchangeServerPort',), # Accounts (both)
            '03:3901': ('int3901',), # Account (Exchange), 25
            '03:3A01': ('DirectoryServicePort',), # Account (Exchange)
            '03:3D01': ('EncryptionAlgorithm', 'rF', ol_type_code), # Accounts (both)
            '03:3E01': ('SigningAlgorithm', 'rF', ol_type_code), # Account (Exchange)
            '03:3F01': ('int3F01',), # Account (Exchange), 2
            '03:4701': ('int4701',), # Account (Exchange), 10
            '03:4801': ('x-mac-type', 'rF', ol_type_code), # Attachment
            '03:4901': ('x-mac-creator', 'rF', ol_type_code), # Attachment
            '03:4A01': ('type4A01', 'rF', ol_type_code), # Attachment
            '03:4B01': ('type4B01', 'rF', ol_type_code), # Attachment
            '03:4C01': ('int4C01',), # Attachment
            '03:4E01': ('FolderType', 'rF', ol_type_code), # Folder
            '03:4F01': ('FolderClass', 'fE', OlFolderClass), # Folder
            '03:5101': ('ItemCount',), # Folder
            '03:5201': ('FolderID',), # Folder
            '03:5401': ('CalendarDefaultTimezone',), # Main, in ms_tzid
            '03:5501': ('CalendarWorkDayStarts',), # Main, minutes
            '03:5601': ('CalendarWorkDayEnds',), # Main, minutes
            '03:5701': ('DefaultEventReminderBefore',), # Main
            '03:5801': ('int5801',), # Main, null or 1
            '03:5901': ('int5901',), # Category, null or 0/2/6
            '03:5A01': ('int5A01',), # Category, null or 0/1
            '03:5B01': ('int5B01',), # Attachment, ?
            '03:5C01': ('int5C01', 'rF', lambda b: unpack('<q', b)[0]), # Attachment, ?
            '03:E803': ('PictureBlockID', 'rX', None), # Contact Picture
            '03:E903': ('PictureFormat', 'rF', ol_type_code), # Contact Picture

            # bstrings
            '08:03': ('bytes03',), # Message
            '08:04': ('bytes04',), # Message
            '08:05': ('SearchData',), # Search

            # booleans
            '0B:02': ('bool02',), # Attendee, Country always true
            '0B:03': ('IsRecurring',), # Event Is Recurring, Attendee
            '0B:04': ('bool04',), # Attendee
            '0B:05': ('Completed',), # Task
            '0B:06': ('bool06',), # Task, always false, deprecated in v16
            '0B:07': ('AllDayEvent',), # Event
            '0B:08': ('HasReminder',), # Task, Message, Contact
            '0B:09': ('bool09',), # Event
            '0B:0A': ('bool0A',), # Event
            '0B:0B': ('IsMyMeeting',), # Event
            '0B:0D': ('bool0D',), # Event
            '0B:0E': ('bool0E',), # Event
            '0B:0F': ('bool0F',), # Event
            '0B:10': ('Overdue',), # Task, Event
            '0B:11': ('AllowNewTimeProposal',), # Event
            '0B:13': ('bool13',), # Event
            '0B:14': ('IsCancelled',), # Event
            '0B:15': ('CanJoinOnline',), # Event
            '0B:16': ('DoNotForward',), # Event
            '0B:18': ('bool18',), # Event
            '0B:1F': ('bool1F',), # Message
            '0B:23': ('HasDownloadedExternalImages',), # Message
            '0B:24': ('bool24',), # Message, false or null
            '0B:25': ('bool25',), # Message, false or null
            '0B:36': ('bool36',), # Message
            '0B:38': ('bool38',), # Message
            '0B:39': ('bool39',), # Message
            '0B:3C': ('bool3C',), # Message
            '0B:3D': ('DidReply',), # Message
            '0B:3E': ('DidForward',), # Message
            '0B:40': ('bool40',), # Message, false or null
            '0B:41': ('HasAttachmentOrInline',), # Message
            '0B:42': ('Sent',), # Message
            '0B:4A': ('Sent2',), # Message
            '0B:4B': ('PartiallyDownloaded',), # Message
            '0B:4D': ('HasvCalendar',), # Message
            '0B:50': ('SuppressAutobackfill',), # Message
            '0B:51': ('MentionedMe',), # Message
            '0B:52': ('bool52',), # Message
            '0B:53': ('HasAttachment',), # Message
            '0B:55': ('bool55',), # Message
            '0B:E1': ('boolE1',), # Contact, always false
            '0B:E2': ('JapaneseFormat',), # Contact
            '0B:2C01': ('bool2C01',), # Account (Mail), true
            '0B:2D01': ('bool2D01',), # Account (Exchange), true
            '0B:3301': ('SignOutgoingMessages',), # Account (both)
            '0B:3401': ('SignIncludeCertificate',), # Account (both)
            '0B:3501': ('SignSendAsClearText',), # Account (both)
            '0B:3601': ('EncryptOutgoingMessages',), # Accounts (both)
            '0B:3B01': ('bool3B01',), # Account (Exchange), true
            '0B:3C01': ('bool3C01',), # Account (Exchange), false
            '0B:3D01': ('bool3D01',), # Account (Exchange), true
            '0B:3F01': ('DirectoryServiceUseSSL',), # Account (Exchange)
            '0B:4001': ('DirectoryServiceUseExchangeCreds',), # Account (Exchange)
            '0B:5301': ('bool5301',), # Account (Exchange), false
            '0B:5401': ('bool5401',), # Account (Exchange), false
            '0B:5601': ('bool5601',), # Account (Exchange), true
            '0B:5A01': ('bool5A01',), # Account (Exchange), true
            '0B:5B01': ('bool5B01',), # Attachment, ?
            '0B:5C01': ('bool5C01',), # Attachment, ?
            '0B:5E01': ('bool5E01',), # Attachment, ?
            '0B:5F01': ('bool5F01',), # Folder, false or null
            '0B:6001': ('bool6001',), # Folder, false or null            
            '0B:6401': ('bool6401',), # Folder
            '0B:6501': ('ContainsPartialDwnldMsgs',), # Folder
            '0B:6601': ('WorkOffline',), # Main
            '0B:6701': ('bool6701',), # Main, 1/2
            '0B:6801': ('DefaultEventReminderEnabled',), # Main
            '0B:6901': ('PlaySoundNewMessage',), # Main
            '0B:6A01': ('PlaySoundNoNewMessages',), # Main
            '0B:6B01': ('PlaySoundSentMessage',), # Main
            '0B:6C01': ('PlaySoundSyncError',), # Main
            '0B:6D01': ('PlaySoundWelcome',), # Main
            '0B:6E01': ('PlaySoundReminder',), # Main
            '0B:6F01': ('CalendarWorkWeekSu',), # Main
            '0B:7001': ('CalendarWorkWeekMo',), # Main
            '0B:7101': ('CalendarWorkWeekTu',), # Main
            '0B:7201': ('CalendarWorkWeekWe',), # Main
            '0B:7301': ('CalendarWorkWeekTh',), # Main
            '0B:7401': ('CalendarWorkWeekFr',), # Main
            '0B:7501': ('CalendarWorkWeekSa',), # Main
            '0B:7601': ('NotifyBounceIconInDock',), # Main
            '0B:7801': ('ReplyWithDefaultEmailAccount',), # Main
            '0B:7901': ('AssignMessagesToContactCategories',), # Main
            '0B:7A01': ('NotifyDisplayAlert',), # Main
            '0B:7B01': ('NotifyShowPreviewInAlert',), # Main
            '0B:7C01': ('bool7C01',), # Category, 5/15
            '0B:7D01': ('Read',), # Message
            '0B:7E01': ('IsLocalCategory',), # Category
            '0B:8001': ('bool8001',), # Accounts (both), true
            '0B:8101': ('bool8101',), # Account (Mail)
            '0B:8201': ('bool8201',), # Account (Mail)
            '0B:8601': ('bool8601',), # Account (Mail)
            '0B:9501': ('bool9501',), # Folder
            '0B:9801': ('bool9801',), # Account (Mail)
            '0B:9A01': ('SyncSharedMailboxes',), # Account (Exchange)
            '0B:9B01': ('bool9B01',), # Folder
            '0B:9C01': ('bool9C01',), # Account (Exchange)
            '0B:9D01': ('bool9D01',), # Folder
            '0B:9E01': ('bool9E01',), # Account (Exchange)
            '0B:9F01': ('bool9F01',), # Folder
            '0B:A101': ('boolA101',), # Folder
            '0B:A201': ('boolA201',), # Folder
            '0B:A301': ('boolA301',), # Folder
            '0B:A401': ('boolA401',), # Account (Mail)
            '0B:A501': ('boolA501',), # Account (Mail)

            # data access objects - collections, lists, etc.
            #  Recurrence, MessageSourceHeader
            '0D:01': ('obj01',),
            #  Events
            '0D:02': ('RRule', 'fC', OlkRecurrence),
            '0D:07': ('ReplyTo', 'fF', self._event_reply_to_parse),
            '0D:09': ('Timezone', 'fC', OlkTimezone),
            '0D:0B': ('Attendees', 'fL', OlkAttendee),
            '0D:0D': ('Organizer', 'fF', self._message_user_parse),
            '0D:0E': ('AttachmentExchangeID', 'fX', None),
            '0D:0F': ('Timezone2', 'fC', OlkTimezone), # duplicate
            '0D:82': ('AttachmentBlockID', 'fX', None),
            #  Messages
            '0D:03': ('From', 'fF', self._message_user_list_parse),
            '0D:04': ('From2', 'fF', self._message_user_list_parse), # same as from
            '0D:05': ('MsrcBlockStruct', 'fC', OlkMultipartType),
            '0D:06': ('From3', 'fF', self._message_user_list_parse), # same as from
            '0D:1E': ('To', 'fF', self._message_user_list_parse),
            '0D:1F': ('CC', 'fF', self._message_user_list_parse),
            '0D:20': ('BCC', 'fF', self._message_user_list_parse),
            '0D:21': ('AttachmentMetadata', 'fL', OlkAttachment),
            '0D:2D': ('MeetingAttendees', 'fF', self._message_user_list_parse),
            '0D:C1': ('ActionsTaken', 'fF', self._actions_taken_parse),
            '0D:80': tuple(),
            #  Contacts
            '0D:62': ('obj62',), # always 0? probably a list, but not sure of what
            #  Attachments
            '0D:3301': ('AttcBlockStruct', 'fC', OlkContentType),
            '0D:3E01': ('obj3E01',),
            #  Categories
            '0D:3401': ('BackgroundColor', 'fF', ol_color),
            #  Main
            '0D:3801': tuple(),
            '0D:3901': ('AddressFormats', 'fL', OlkAddressFormat),
            '0D:4501': ('NewOutlookObject',), # something to do with New Outlook
            #  Time Zone
            '0D:3F01': ('Standard', 'fL', OlkTZProp),
            '0D:4001': ('Daylight', 'fL', OlkTZProp),
            #  Account (both)
            '0D:2C01': ('Certificates',),
            #  Account (Exchange)
            '0D:4201': ('bplist1',),
            '0D:4301': ('bplist2',),

            # long integers (8 bytes signed)
            '14:01': ('long01',), # Search, always 0
            '14:61': ('long61', 'rX'), # Contact
            '14:2C01': ('AttachmentBlockID', 'rX', None), # Attachment Block Id
            '14:2D01': ('SyncMapBlockID', 'rX', None), # Folder SyncMap Block Id
            '14:2E01': ('FolderSyncBlockID', 'rX', None), # Folder SyncMap Block Id
            '14:3001': ('AccountUID',), # Event, Category
            '14:3201': ('ExchangeAccountUID',), # Account (Mail)
            '14:3301': ('MailAccountUID',), # Main
            '14:3401': ('LDAPAccountUID',), # Account (Exchange)
            '14:3601': ('ExchangeAccountUID',), # Main
            '14:3701': ('long3701',), # Main, 55834574849
            '14:3801': ('MailAccountUID',), # Main
            '14:3901': ('GroupID',),

            # user-defined blobs
            #  Contact
            '1D:66': ('EmailAddress_1',),
            '1D:67': ('EmailAddress_2',),
            '1D:68': ('EmailAddress_3',),
            '1D:69': ('EmailAddress_4',),
            '1D:6A': ('EmailAddress_5',),
            '1D:6B': ('EmailAddress_6',),
            '1D:6C': ('EmailAddress_7',),
            '1D:6D': ('EmailAddress_8',),
            '1D:6E': ('EmailAddress_9',),
            '1D:6F': ('EmailAddress_10',),
            '1D:70': ('EmailAddress_11',),
            '1D:71': ('EmailAddress_12',),
            '1D:72': ('EmailAddress_13',),
            '1D:78': ('IMAddress_1',),
            '1D:79': ('IMAddress_2',),
            '1D:7A': ('IMAddress_3',),
            '1D:7B': ('IMAddress_4',),
            '1D:7C': ('IMAddress_5',),
            '1D:7D': ('IMAddress_6',),
            '1D:7E': ('IMAddress_7',),
            '1D:7F': ('IMAddress_8',),
            '1D:80': ('IMAddress_9',),
            '1D:81': ('IMAddress_10',),
            '1D:82': ('IMAddress_11',),
            '1D:83': ('IMAddress_12',),
            '1D:84': ('IMAddress_13',),

            # ANSI strings
            '1E:01': ('Address',), # Event, Attendee
            '1E:02': ('MessageID',),
            '1E:03': ('string03',), # Message.Simple?
            '1E:04': ('string04',), # Event CalendarUID, Message Header
            '1E:07': ('string07',), # Event, ?
            '1E:0A': ('MessageClass',), # Event
            '1E:1E': ('References2',), # Message, only present once, similar to references
            '1E:1F': ('References3',), # Message, only present once, similar to references
            '1E:22': ('InReplyTo',), # Message
            '1E:23': ('vCalendar',), # Message
            '1E:24': ('References',), # Message
            '1E:25': ('string25',), # Message
            '1E:2B': ('string2B',), # Message
            '1E:2C': ('string2C',), # Message
            '1E:40': ('MessageClass',), # Message
            '1E:41': ('string41',), # Message
            '1E:67': ('ExchangeID',),
            '1E:68': ('ExchangeChangeKey',),
            '1E:2C01': ('EmailAddress',), # Accounts (both)
            '1E:2D01': ('ExchangeServerURL',), # Accounts (both)
            '1E:3101': ('string3101',), # Account (Exchange), AAMK...AAA=
            '1E:3401': ('string3401',), # Account (Exchange), empty
            '1E:3501': ('directory_service_search_base',), # Account (Exchange)
            '1E:3801': ('string3801',), # Account (Exchange), empty
            '1E:3901': ('string3901',), # Account (Exchange), empty
            '1E:3A01': ('EmailAddress2',), # Account (Exchange)
            '1E:3B01': ('OutlookOABURL',), # Account (Exchange)
            '1E:3C01': ('ReceiptIPAddress',), # Account (Exchange)
            '1E:3E01': ('FileType',), # Attachment
            '1E:3F01': ('ContentType',), # Attachment
            '1E:4001': ('FileName',), # Attachment
            '1E:4201': ('ExchangeGUID',), # Category
            '1E:4301': ('OutlookManageURL',), # Account (both)
            '1E:4401': ('OutlookClutterURL',),
            '1E:4D01': ('OutlookAPIURL',), # Account (Exchange)
            '1E:4E01': ('CalendarOwnerAccount',), # Folder
            '1E:4F01': ('OutlookSearchURL',), # Account (Exchange)
            '1E:5001': ('CalendarToken',), # Folder
            '1E:5101': ('string5101',), # Attachment
            '1E:5201': ('ExchangeEWSURL',), # Account (Exchange)

            # Unicode strings (message contents, xml, etc.)
            '1F:01': ('Name',), # Contact first name, Event body, Message subject, Search/Task/Attendee name
            '1F:02': ('unicode02',), # Contact last name, Event subject
            '1F:04': ('unicode04',), # Contact body, Event location
            '1F:05': ('CalendarOwnerName',), # Event
            '1F:06': ('HomeAddressStreet',), # Contact
            '1F:07': ('HomeAddressCity',), # Contact
            '1F:08': ('unicode08',), # Event conference; Country, - or ,; Contact Home Address 
            '1F:09': ('unicode09',), # Event conference; Contact Home Address
            '1F:0A': ('unicode0A',), # Event conference; Contact Home Address
            '1F:0B': ('unicode0B',), # Event conference; Contact Home Phone; Task body
            '1F:0C': ('unicode0C',), # Event conference; Contact Home Fax
            '1F:0D': ('ConferenceSettings',), # Event conference
            '1F:0E': ('ConferenceSettings2',), # Event conference
            '1F:0F': ('PhoneHome2',), # Contact
            '1F:10': ('ConferenceUUID',), # Event conference
            '1F:14': ('Company',), # Contact
            '1F:15': ('WorkTitle',), # Contact
            '1F:16': ('WorkAddressStreet',), # Contact
            '1F:17': ('WorkAddressCity',), # Contact
            '1F:18': ('WorkAddressState',), # Contact
            '1F:19': ('WorkAddressPostalCode',), # Contact
            '1F:1A': ('WorkAddressCountry',), # Contact
            '1F:1B': ('Department',), # Contact
            '1F:1C': ('OfficeLocation',), # Contact
            '1F:1D': ('PhoneWork',), # Contact
            '1F:1E': ('unicode1E',), # Message body; Contact Work Fax
            '1F:1F': ('PhonePager',), # Contact
            '1F:20': ('WebPageWork',), # Contact
            '1F:21': ('PhoneMobile',), # Contact
            '1F:22': ('PhoneWork2',), # Contact
            '1F:23': ('unicode23',), # Message Recipients; Contact Primary Phone
            '1F:24': ('Alias',), # Contact
            '1F:25': ('PhoneAssistant',), # Contact
            '1F:27': ('Preview',), # Message
            '1F:2A': ('ThreadTopic',), # Message
            '1F:2F': ('ThreadTopic2',), # Message
            '1F:3E': ('Nickname',), # Contact
            '1F:3F': ('Title',), # Contact
            '1F:40': ('Suffix',), # Contact
            '1F:41': ('Custom1',), # Contact
            '1F:42': ('Custom2',), # Contact
            '1F:43': ('Custom3',), # Contact
            '1F:44': ('Custom4',), # Contact
            '1F:45': ('Custom5',), # Contact
            '1F:46': ('Custom6',), # Contact
            '1F:47': ('Custom7',), # Contact
            '1F:48': ('Custom8',), # Contact
            '1F:49': ('Date1',), # Contact, DOW, Mon DD, YYYY
            '1F:4A': ('Date2',), # Contact, DOW, Mon DD, YYYY
            '1F:4B': ('Birthday',), # Contact, DOW, Mon DD, YYYY
            '1F:4C': ('Anniversairy',), # Contact, DOW, Mon DD, YYYY
            '1F:57': ('YomiLastName',), # Contact
            '1F:58': ('YomiFirstName',), # Contact
            '1F:59': ('YomiCompanyName',), # Contact
            '1F:5A': ('XML:Tasks',), # Event/Message; Contact Extra Phones
            '1F:5B': ('XML:Meetings',), # Event/Message; Contact Extra Phones
            '1F:5C': ('XML:Addresses',), # Event/Message; Contact Extra Phones
            '1F:5D': ('XML:Emails',), # Event/Message; Contact Extra Phones
            '1F:5E': ('XML:Phones',), # Event, Message
            '1F:5F': ('XML:Urls',), # Event, Message
            '1F:60': ('XML:Contacts',), # Event, Message
            '1F:61': ('ThreadTopic',), # Event, Message
            '1F:62': ('HTMLBody',), # Message
            '1F:6A': ('MiddleName',), # Contact, for Message this is Card Data
            '1F:C8': ('Spouse',), # Contact
            '1F:C9': ('Child',), # Contact
            '1F:D5': ('AstrologicalSign',), # Contact
            '1F:D6': ('Age',), # Contact
            '1F:E5': ('BloodType',), # Contact
            '1F:E6': ('InterestsHobbies',), # Contact
            '1F:E7': ('Initials',), # Contact
            '1F:FA': ('HomeAddressFormat',), # Contact, Country Code
            '1F:FB': ('WorkAddressFormat',), # Contact, Country Code
            '1F:FD': ('PhoneOther',), # Contact
            '1F:FE': ('PhoneOtherFax',), # Contact
            '1F:FF': ('PhoneRadio',), # Contact
            '1F:0001': ('OtherAddressStreet',), # Contact
            '1F:0101': ('OtherAddressCity',), # Contact
            '1F:0201': ('OtherAddressState',), # Contact
            '1F:0301': ('OtherAddressPostalCode',), # Contact
            '1F:0401': ('OtherAddressCountry',), # Contact
            '1F:0601': ('OtherAddressFormat',), # Contact, Seems bugged in Outlook 16 
            '1F:2C01': ('DisplayName',), # Accounts (both)
            '1F:2E01': ('UserName',), # Accounts (Mail)
            '1F:2F01': ('EmailAddressUnicode',), # Accounts (both), same as string email address
            '1F:3101': ('unicode3101',), # Account (Exchange), null or empty
            '1F:3401': ('FileNameUnicode',), # Attachment
            '1F:3501': ('Name',), # Category
            '1F:3601': ('Name',), # Folder
            '1F:3701': ('Title',), # Note
            '1F:3801': ('Body',), # Note
            '1F:3901': ('Name',), # Signature
            '1F:3A01': ('Body',), # Signature
            '1F:3B01': ('SoundSet',), # Main
            '1F:3C01': ('DefaultCategory',), # Accounts (Exchange)
            '1F:4401': ('unicode4401',), # Attachment
            '1F:4C01': ('CalendarOwnerName',), # Folder

            # additional long codes
            #   14 == max(15.values()) except for a few Messages
            #   Mostly 1, but can be up to 15
            #   16 is not always present, always zero except for a few Messages
            #   Can be 1, 4, or 5 when non-zero
            '20:14': ('foot14',),
            '20:15': ('foot15', 'rF', lambda b: self._read_sizes(b, 'q')),
            '20:16': ('foot16',),

            # GUIDs for a class (CLSID)
            '48:00': ('UUID',), # Category, Contact, Event, Note, Task

            # Apple Mac Absolute timestamps (seconds since Jan 1 2001)
            '4D:01': ('TimeSent',),
            '4D:02': ('TimeReceived',),
            '4D:04': ('ModDate',),
            '4D:09': ('StartDate',), # Task
            '4D:0A': ('CompletedDate',), # Task
            '4D:0B': ('DueDate',), # Task
            '4D:0C': ('Reminder',), # Task
            '4D:0D': ('Reminder2',), # Task, same as reminder 
            '4D:10': ('DownloadDate2',), # Event, sent by me, same as 11
            '4D:11': ('DownloadDate',),
            '4D:12': ('CreationTime',),
            '4D:15': ('date15',), # Message, off from Received by ~seconds
            '4D:16': ('DismissTime',), # Event
            '4D:17': ('ReplyTime',), # Event
            '4D:18': ('OwnerCriticalChange',), # Event
            '4D:19': ('date19',), # Event
            '4D:1A': ('date1A',), # Message, similar to date15?
            '4D:1B': ('ScheduledSendDate',), # Message
            '4D:2C01': ('date2C01',), # Account (Exchange), 2016-08-27
            '4D:2D01': ('date2D01',), # Account (Exchange)
            '4D:2E01': ('date2E01',), # Folder, ?
            '4D:2F01': ('date2F01',), # Folder, ?
            '4D:3001': ('date3001',), # Folder
            '4D:3101': ('CreatedDate',), # Note
            '4D:3201': ('CreatedDate',), # Account (Exchange)
            '4D:3301': ('CreatedDate',), # Category
            '4D:3401': ('date3401',), # Category, null except one
            '4D:3501': ('date3501',), # Category, always null
            '4D:3601': ('date3601',), # Category, always null
            '4D:3701': ('date3701',), # Category, always null
            '4D:3801': ('date3801',), # Category, always null
            '4D:3901': ('date3901',), # Category, always null
            '4D:3A01': ('date3A01',), # Account (Exchange), 2016-08-27
            '4D:3D01': ('date3D01',), # Account (Mail)

            # TZPROP attributes
            '4643:7A74': ('TZID', 'rF', lambda b: b.decode()),
            '5A54:4449': ('MSTZID', 'rF', lambda b: unpack('<i', b)[0]),
            '614E:656D': ('TZLongName', 'rF', lambda b: b.decode()),
            # 7453 -> STANDARD time property
            '7453:6C52': ('RRule', 'rF', lambda b: b.decode()),
            '7453:6F54': ('OffsetTo', 'rF', lambda b: b.decode()),
            '7453:7246': ('OffsetFrom', 'rF', lambda b: b.decode()),
            '7453:7453': ('StartDate', 'rF', lambda b: dt_winminutes(unpack('<i', b)[0])),
            # 4C44 -> DAYLIGHT time property
            '4C44:6C52': ('RRule', 'rF', lambda b: b.decode()),
            '4C44:6F54': ('OffsetTo', 'rF', lambda b: b.decode()),
            '4C44:7246': ('OffsetFrom', 'rF', lambda b: b.decode()),
            '4C44:7453': ('StartDate', 'rF', lambda b: dt_winminutes(unpack('<i', b)[0])),
            }

    def _parse(self, buff):
        # check magic bytes
        assert buff.read(4) == b'\xd0\x0d\x00\x00'
        _ = buff.read(4)

        out = dict()
        # check if this is an entity or a block
        entity_block = unpack('<i', buff.read(4))[0]
        # if it's a block, use the block parsing logic instead
        if entity_block == 1:
            return self._parse_entity(buff, out)
        elif entity_block == 2:
            return self._parse_block(buff, out)
        else:
            print('Invalid file, entity/block value =', entity_block)
            return out

    def _parse_entity(self, buff, out):
        # read the initial header
        out['RecordID'] = unpack('<i', buff.read(4))[0]
        class_id = unpack('<i', buff.read(4))[0]
        out['head:20'] = buff.read(12)
        out['BlockType'] = ol_type_code(buff.read(4))
        out['ItemID'] = buff.read(4)

        # get the schema to use based on the class ID
        schema = CLASSTOSCHEMA[class_id]
        self.skip_indb.extend(schema.get('skip_indb', list()))

        # read the main collection
        out.update(self._parse_collection(buff.read(), schema))

        # allow for extra processing
        if schema['class'] == 'OlkEvent':
            out = self._additional_parsing_event(out)
        elif schema['class'] == 'OlkContact':
            out = self._additional_parsing_contact(out)
        elif schema['class'] == 'OlkMain':
            out = self._additional_parsing_main(out)
        out = self._additional_parsing_collect_xml(out)

        return out

    def _parse_block(self, buff, out):
        # read the initial header
        out['BlockID'] = buff.read(20)
        out['BlockType'] = ol_type_code(buff.read(4))
        out['ItemID'] = buff.read(4)

        if out['BlockType'] == 'ImgB':
            # Binary file data
            out['FileData'] = buff.read()
        elif out['BlockType'] in ('Attc', 'MSrc', 'ClAt'):
            # Text file data
            out['FileContents'] = buff.read().decode()
        elif out['BlockType'] == 'RcnA':
            out.update(self._block_rcna_parse(buff))
        # TBD ----
        elif out['BlockType'] == 'ExSM':
            out = self._block_sync_map(buff, out)
        elif out['BlockType'] == 'ExFS':
            out = self._block_folder_sync(buff, out)
        else:
            print('Unknown block type', out['BlockType'])
            out['BlockData'] = buff.read()
        return out

    def _parse_collection(self, chunk, schema):
        # this is a common pattern across several sections of Olk data files
        # first, there are three integers
        #  1. number of items
        #  2. size of the header (including these 12 bytes)
        #  3. size of the body
        (num_items, head_size, body_size) = unpack('<3i', chunk[:12])

        # then, the rest of the header is as array defining the size of each
        # item in the body
        sizes = self._read_sizes(chunk[12:head_size])

        # we can use the sizes to split the data in the body section into
        # items corresponding to a key
        items = self._split_with_array(sizes, chunk[head_size:])

        # finally, the OLKDATAFILE format dictionary tells us what the human-
        # readable name and type of each entry is, based on the key
        items = self._format_items(self.OLKDATAFILE, items, schema)

        return items

    def _read_sizes(self, chunk, fmt='i'):
        # array schema is BBBBAAAA ########
        # AAAA corresponds to a .Net Varient Type Code
        #   Some regions have longer, custom codes
        # BBBB is an ID that's unique within each varient type
        # flip them around for the format dict mapping
        array = defaultdict(dict)
        valsize = 8 if fmt == 'q' else 4
        for i in range(0, len(chunk), 4 + valsize):
            item = chunk[i:i + 4 + valsize]
            a = hex_str_arr(item[2:4] if item[2] > 0 else item[3:4])
            b = hex_str_arr(item[0:2] if item[1] > 0 else item[0:1])
            array[(a, b)] = unpack('<' + fmt, item[4:])[0]
        return array

    def _split_with_array(self, arr, chunk):
        # split body into sections using the header array
        split = dict()
        for k, size in arr.items():
            if size == 0:
                split[k] = b''
            else:
                split[k] = chunk[:size]
                chunk = chunk[size:]
        if chunk:
            print(len(chunk), 'bytes remaining')
        return split

    def _format_items(self, formats, items, schema):
        # format items
        # first, use the varient type code from the size array to read bytes
        # then, look up the mapped field name and an optional handler from
        #  the formats dict
        out = dict()
        for ((vartype, idx), chunk) in items.items():
            # grab the format information
            # format is ([name], [handler_mode], [handler])
            # handler_mode determines how to handle the value
            #   first character is r or f - raw or formated
            #   second character is C, L, E, F, or X - collection, list, enum,
            #     function, or None
            key = vartype + ':' + idx
            if key not in formats:
                print('Unmapped key: ', key, schema['class'])
            fmt = formats.get(key, tuple())
            out_name = fmt[0] if len(fmt) >= 1 else key
            handler_mode = fmt[1] if len(fmt) >= 2 else None
            handler = fmt[2] if len(fmt) == 3 else None

            # get override handlers
            if key in schema.get('override', dict()):
                out_name, handler_mode, handler = schema['override'][key]

            # see VarientType dict up top, these are the ones we see in
            # the OlkData files
            if handler_mode and handler_mode[0] == 'r':
                pass
            elif vartype == '02':   # short (signed 2 byte int)
                try:
                    chunk = unpack('<h', chunk)[0]
                except:
                    print('error on', key, schema['class'], chunk)
            elif vartype == '03': # int (signed 4 byte int)
                try:
                    chunk = unpack('<i', chunk)[0]
                except:
                    print('error on', key, schema['class'], chunk)
            elif vartype == '08': # bstring (byte string)
                chunk = chunk
            elif vartype == '0B': # bool
                chunk = unpack('<?', chunk)[0]
            elif vartype == '0D': # data access object (pass bytes to handler)
                pass
            elif vartype == '14': # long (signed 8 byte int)
                chunk = unpack('<q', chunk)[0]
            elif vartype == '1D': # simple strings (emails, chat names, etc.)
                chunk = chunk.decode()
            elif vartype == '1E': # string (ANSI)
                chunk = chunk.decode()
            elif vartype == '1F': # Unicode string ("wide")
                chunk = chunk.decode('utf-16')
            elif vartype == '20': # ?
                chunk = unpack('<q', chunk)[0]
            elif vartype == '48': # GUID
                pass
            elif vartype == '4D': # Apple Mac Absolute Date (stored as double)
                chunk = dt_macabsolute(unpack('<d', chunk)[0])
            # TIMEZONE and TZPROP types
            elif vartype in ('4643', '7453', '4C44'):
                pass
            else:
                print('New Variant type:', vartype, schema['class'])

            # invoke handler using handler_mode
            if handler is not None:
                if handler_mode[1] == 'L':
                    chunk = self._parse_list(chunk, handler)
                elif handler_mode[1] == 'C':
                    chunk = self._parse_collection(chunk, handler)
                elif handler_mode[1] == 'E':
                    try:
                        chunk = handler[chunk]
                    except KeyError:
                        print(handler, chunk)
                elif handler_mode[1] == 'F':
                    chunk = handler(chunk)
                else:
                    raise ValueError("Invalid handler mode")

            # check to see if we remap this
            if 'remap' in schema:
                out_name = schema['remap'].get(key, out_name)

            # skip attributes that aren't useful
            skip_items = ['foot14', 'foot15', 'foot16']
            skip_items.extend(schema.get('skip_null', list()))
            skip_items.extend(schema.get('skip_dupe', list()))
            if out_name in skip_items:
                continue

            # store item in output dictionary
            out[out_name] = chunk

        return out

    def _parse_list(self, chunk, item_format_dict):
        # This is another common patter, a list of collections of the same type
        # First is an integer indicating how long the list is
        (length,) = unpack('<i', chunk[:4])
        chunk = chunk[4:]

        # then, that many shorts with the length of each item
        item_sizes = unpack('<' + str(length) + 'h', chunk[:length * 2])
        chunk = chunk[length * 2:]

        # split items and parse
        items = list()
        for size in item_sizes:
            # each item is a collection
            items.append(self._parse_collection(chunk[:size], item_format_dict))
            chunk = chunk[size:]

        return items

    def _event_reply_to_parse(self, chunk):
        # structure: one null byte, then two shorts
        #  first short is 1, second is the length of the list
        (_, num_entries) = unpack('<xhh', chunk[:5])
        chunk = chunk[5:]
        
        # each entry starts with a 4 byte int of the string length followed by
        #  one byte with the same value, then the string
        # entries are terminated by 4 null bytes
        email_list = list()
        while chunk:
            (size,) = unpack('<i', chunk[:4])
            email_list.append(chunk[5:size + 5].decode('ascii'))
            chunk = chunk[size + 9:]
        return email_list

    def _message_user_list_parse(self, chunk):
        # first four bytes are the length, then an \x02
        (n, _) = unpack('<ib', chunk[:5])
        chunk = chunk[5:]
        out = list()
        for i in range(n):
            size = unpack('<h', chunk[:2])[0]
            chunk = chunk[2:]
            out.append(self._message_user_parse(chunk[:size]))
            chunk = chunk[size:]

        return out

    def _message_user_parse(self, chunk):
        # First are some flags, then 22 zero bytes
        # H1: always 3 (Type = SMTP?)
        # B2: OlUserType
        # B3: always 3
        # B4: 0, 1, 2, 7, 8
        #   2: Distribution list
        #   7: External Email
        #   8: Private Group?
        # B5: 0, 1
        flags = unpack('<h4b', chunk[:6])
        chunk = chunk[28:]

        # Then an email string
        email_size = unpack('<i', chunk[:4])[0]
        chunk = chunk[4:]
        email = chunk[:email_size].decode()
        chunk = chunk[email_size:]

        # Then a name string
        name_size = unpack('<i', chunk[:4])[0]
        chunk = chunk[4:]
        name = chunk[:name_size].decode('utf-16')
        chunk = chunk[name_size:]

        return {'Address': email, 'Name': name, 'Type': OlUserType[flags[1]]}

    def _actions_taken_parse(self, chunk):
        # this is like a collection, but can have arbitrary numbers of values
        (num_items, head_size, body_size) = unpack('<3i', chunk[:12])
        sizes = self._read_sizes(chunk[12:head_size])
        items = self._split_with_array(sizes, chunk[head_size:])

        # first grab the number of actions
        action_count = unpack('<h', items.pop(('00', '01')))[0]

        # then, loop through actions and add them to a list
        actions = list()
        for i in range(action_count):
            a = items.pop(('00', hex(100 + i*10)[2:].upper()))
            b = items.pop(('00', hex(101 + i*10)[2:].upper()))
            c = items.pop(('00', hex(102 + i*10)[2:].upper()), b'\xff\xff\xff\xff')
            actions.append({
                'Type': OlAction[unpack('<h', a)[0]],
                'Date': dt_macabsolute(unpack('<d', b)[0]),
                'RecordID': unpack('<i', c)[0]
                })
        return actions

    def _additional_parsing_event(self, out):
        # normalize the recurrence rule parameters
        if 'RRule' in out:
            # for some reason, the Daily intervals are in minutes
            if out['RRule']['RecurrenceType'] == 'Daily':
                out['RRule']['Interval'] = out['RRule']['Interval'] / 1440
            # Weekly uses a list of week days
            elif out['RRule']['RecurrenceType'] == 'Weekly':
                out['RRule']['Day'] = out['RRule'].pop('WeekDay')
            # Monthly uses a month day, already mapped
            # MonthNth uses a weekday and an offset (nth XXday of the month)
            elif out['RRule']['RecurrenceType'] == 'MonthNth':
                out['RRule']['Day'] = out['RRule'].pop('MonthDOW')
                setpos = out['RRule'].pop('MonthNth')
                # 5 means "last XXday of month", swap it for -1
                out['RRule']['SetPos'] = -1 if setpos == 5 else setpos
            # Yearly and YearNth do not seem to be in use

            # truncate recurrence rule dates
            out['RRule']['StartDate'] = out['RRule']['StartDate'].date()
            out['RRule']['Until'] = out['RRule']['Until'].date()

            # depending on EndType, keep either until, occurrences, or neither
            if out['RRule']['EndType'] in ('NoEndDate', 'EndAfterCount'):
                _ = out['RRule'].pop('Until', None)
            if out['RRule']['EndType'] in ('NoEndDate', 'EndOnDate'):
                _ = out['RRule'].pop('Occurrences', None)

            # Fill in missing Exceptions
            if 'ExceptionDates' not in out['RRule']:
                out['RRule']['ExceptionDates'] = list()

        # add timezone to organizer start & end date
        tz = out['Timezone']['TZID']
        out['StartDateOrganizer'] = out['StartDateOrganizer'].replace(tzinfo=ZoneInfo(tz))
        out['EndDateOrganizer'] = out['EndDateOrganizer'].replace(tzinfo=ZoneInfo(tz))
        out['ReplyTime'] = out['ReplyTime'].replace(tzinfo=ZoneInfo(tz))

        return out

    def _additional_parsing_contact(self, out):
        # get email address types
        email_flags = out.pop('EmailTypesRaw', 0)
        email_count = out.pop('EmailCount', 0)
        types = self._type_list_parse(email_flags, email_count)

        # get email addresses
        emails = self._extract_address_list(out, 'EmailAddress')
        emails = [{'Type': t, 'Address': e} for (t, (i, e)) in zip(types, emails)]
        out['EmailAddresses'] = emails

        # note default
        email_default = out.pop('DefaultEmailRaw', b'\x66\x00\x00\x1d')
        if emails:
            out['DefaultEmailAddress'] = emails[email_default[0] - 102]['Address']

        # get im address types
        im_flags = out.pop('IMTypesRaw', 0)
        im_count = out.pop('IMCount', 0)
        types = self._type_list_parse(im_flags, im_count)

        # get im addresses
        ims = self._extract_address_list(out, 'IMAddress')
        ims = [{'Type': t, 'Address': e} for (t, (i, e)) in zip(types, ims)]
        out['IMAddresses'] = ims

        # note default
        im_default = out.pop('DefaultIMRaw', b'\x78\x00\x00\x1d')
        if ims:
            out['DefaultIMAddress'] = ims[im_default[0] - 120]['Address']
        
        return out

    def _additional_parsing_main(self, out):
        # Create address format strings
        if 'AddressFormats' in out:
            fmt_map = dict()
            for fmt in out.pop('AddressFormats'):
                code = fmt.pop('country_code')
                fmt_string = ''

                # Lines 1 and 2
                if (p1 := fmt.pop('part_1', None)):
                    fmt_string += '{' + p1 + '}'
                if (p2 := fmt.pop('part_2', None)):
                    fmt_string += '\n{' + p2 + '}'

                # Line 3
                p5 = fmt.pop('part_5', None)
                p6 = fmt.pop('part_6', None)
                p7 = fmt.pop('part_7', None)
                s5 = fmt.pop('sep_5_6', '') + ' '
                if p5 or p6 or p7:
                    fmt_string += '\n'
                if p5:
                    fmt_string += '{' + p5 + '}'
                    if p6 or p7:
                        fmt_string += s5
                if p6:
                    fmt_string += '{' + p6 + '}'
                    if p7:
                        fmt_string += ' '
                if p7:
                    fmt_string += '{' + p7 + '}'

                # Line 4
                p9 = fmt.pop('part_9', None)
                pA = fmt.pop('part_A', None)
                s9 = fmt.pop('sep_9_A', '') + ' '
                if p9 or pA:
                    fmt_string += '\n'
                if p9:
                    fmt_string += '{' + p9 + '}'
                    if pA:
                        fmt_string += s9
                if pA:
                    fmt_string += '{' + pA + '}'

                # Line 5
                if (pD := fmt.pop('part_D', None)):
                    fmt_string += '\n{' + pD + '}'

                fmt_map[code] = {
                    'format_string': fmt_string,
                    'int14': fmt.get('int14'),
                    'sep_street': fmt.get('sep_street', ' ')
                    }

            out['AddressFormats'] = fmt_map

        return out

    def _additional_parsing_collect_xml(self, out):
        # Collect XML data data (for Messages and Events)
        xml = dict()
        keys = [k for k in out.keys() if k.startswith('XML:')]
        for k in keys:
            xml[k.split(':')[1]] = out.pop(k)
        if xml:
            out['XML'] = xml
        return out

    def _type_list_parse(self, flag, count):
        types = list()
        for _ in range(count):
            t = 'Work'
            if flag & 1:
                t = 'Home'
            elif flag & 2:
                t = 'Other'
            types.append(t)
            flag = flag >> 2
        return types

    def _extract_address_list(self, out, prefix):
        addresses = list()
        for k in list(out.keys()):
            if k.startswith(prefix):
                addr = out.pop(k)
                idx = int(k.split('_')[-1])
                addresses.append((idx, addr))
        return sorted(addresses)

    def _block_rcna_parse(self, buff):
        # Similar to parsing a list...
        # Get the number of chunks and their sizes
        n_chunks = unpack('<i', buff.read(4))[0]
        chunk_sizes = unpack('<' + str(n_chunks) + 'h', buff.read(2 * n_chunks))

        # Break the buffer into chunks
        chunks = list()
        for s in chunk_sizes:
            chunks.append(buff.read(s))

        # The first chunk is the number of address fields
        listcount = unpack('<i', chunks.pop(0))[0]

        # Each address field is a chunk of text, followed by a chunk of sizes
        fields = list()
        for _ in range(listcount):
            fields.append(list())
            values = chunks.pop(0)
            sizesraw = chunks.pop(0)
            sizes = unpack('<' + str(int(len(sizesraw)/4)) + 'i', sizesraw)
            x0 = sizes[0]
            for x1 in sizes[1:]:
                fields[-1].append(values[x0:x1])
                x0 = x1
        # Zip the fields together and map the outputs
        addresses = [{
            'Address': a.decode(),
            'FirstName': f.decode('utf-16'),
            'LastName': l.decode('utf-16')
            } for a, f, l in zip(*fields)]

        out = {
            'RecentAddresses': addresses,
            'tail': chunks
            }

        return out

    def _block_sync_map(self, buff, out):
        # parse body
        _ = buff.read(8) # 00000019 00000000
        out['flag1'] = buff.read(8)
        size1 = unpack('>i', buff.read(4))[0]
        out['part1'] = buff.read(size1)
        _ = buff.read(4) # 00000000
        try:
            count = unpack('>i', buff.read(4))[0]
        except:
            return out
        notnull = unpack('>i', buff.read(4))[0] == 1

        items = list()
        if notnull:
            for _ in range(count):
                item = dict()
                item['h'] = buff.read(20)
                size = unpack('>i', buff.read(4))[0]
                item['ExchangeID'] = buff.read(size).decode()
                size = unpack('>i', buff.read(4))[0]
                item['ExchangeChangeKey'] = buff.read(size).decode()
                count2 = unpack('>i', buff.read(4))[0]
                for _ in range(count2):
                    k_size = unpack('>I', buff.read(4))[0]
                    k = buff.read(k_size).decode()
                    v_size = unpack('>I', buff.read(4))[0]
                    v = buff.read(v_size).decode()
                    item[k] = v
                _ = buff.read(4) # 01000000
                items.append(item)
        out['items'] = items

        rem = buff.read().strip(b'\x00')
        if rem:
            out['rem'] = rem

        return out

    def _block_folder_sync(self, buff, out):
        out['data'] = buff.read()
        return out
