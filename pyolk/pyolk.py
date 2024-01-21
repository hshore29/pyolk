import os
import json
import sqlite3
from os.path import expanduser
from zoneinfo import ZoneInfo
from datetime import date, datetime

from datafiles import OlkDataFile
from mailobjects import *
from utils import *

class PyOLKReader:
    PATH = '/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data'
    tables = list()

    def __init__(self, path=None, mytz=None):
        # Save current directory
        cwd = os.getcwd()

        # Get path to Outlook cache
        mypath = expanduser('~') + self.PATH
        os.chdir(path or mypath)

        # Set default timezone
        self.localtime = ZoneInfo(mytz or 'US/Eastern')

        # Connect to Outlook sqlite db
        db = sqlite3.connect('Outlook.sqlite')
        db.row_factory = sqlite3.Row
        self.cur = db.cursor()

        # Get list of tables that are present in sqlite db
        self.cur.execute("SELECT name FROM sqlite_schema WHERE type ='table';")
        for r in self.cur.fetchall():
            self.tables.append(r['name'])

        # Load the archive
        self.load_archive()

        # Return to original directory
        os.chdir(cwd)

    def get_items(self):
        # Return a list of all archived items, regardless of type
        return list(self.Messages.values()) + \
               list(self.Events.values()) + \
               list(self.Folders.values()) + \
               list(self.Tasks.values()) + \
               list(self.Notes.values()) + \
               list(self.Contacts.values()) + \
               list(self.Categories.values()) + \
               list(self.Signatures.values()) + \
               list(self.SavedSearches.values()) + \
               list(self.Mains.values()) + \
               list(self.AccountsMail.values()) + \
               list(self.AccountsExchange.values())

    def load_archive(self):
        # My archive missing: AccountsLdap, Rules
        t, q = self._mail_query()
        self.Messages = self._get_items(t, q, OlkMessage)
        
        t, q = self._calendar_event_query()
        self.Events = self._get_items(t, q, OlkEvent)

        t, q = self._folder_query()
        self.Folders = self._get_items(t, q, OlkFolder)

        t, q = self._note_query()
        self.Notes = self._get_items(t, q, OlkNote)

        t, q = self._task_query()
        self.Tasks = self._get_items(t, q, OlkTask)

        t, q = self._contact_query()
        self.Contacts = self._get_items(t, q, OlkContact)

        t, q = self._category_query()
        self.Categories = self._get_items(t, q, OlkCategory)

        t, q = self._signature_query()
        self.Signatures = self._get_items(t, q, OlkSignature)

        t, q = self._search_query()
        self.SavedSearches = self._get_items(t, q, OlkSavedSearch)

        t, q = self._main_query()
        self.Mains = self._get_items(t, q, OlkMain)

        t, q = self._acctmail_query()
        self.AccountsMail = self._get_items(t, q, OlkAccountMail)

        t, q = self._acctexch_query()
        self.AccountsExchange = self._get_items(t, q, OlkAccountExchange)

    def _get_items(self, table, select_query, ItemClass):
        # Load all the archived items in a particular table,
        # using the provided ItemClass
        self.cur.execute(select_query)
        items = dict()
        for row in self.cur.fetchall():
            data = self._process_record(dict(row))
            path_to_item = data.pop('PathToDataFile').replace('%20', ' ')
            item = ItemClass(**data)
            item.add_data(OlkDataFile(path_to_item).data())
            if table + '_OwnedBlocks' in self.tables:
                blocks = list()
                self.cur.execute(self._block_query(table), (item.RecordID,))
                for x in self.cur.fetchall():
                    path_to_block = x['PathToDataFile'].replace('%20', ' ')
                    blocks.append(OlkDataFile(path_to_block).data())
                item.add_blockdata(blocks)
            items[item.RecordID] = item
        return items

    ### Get columns from Outlook.sqlite database
    def _process_record(self, r):
        bool_cols = [
            'IsRecurring', 'Completed', 'HasReminder', 'HasAttachment',
            'Hidden', 'IsOutgoingMessage', 'MarkedForDelete', 'MentionedMe',
            'PartiallyDownloaded', 'ReadFlag', 'Sent', 'SuppressAutoBackfill',
            'IsLocalCategory'
            ]
        for k in bool_cols:
            if k in r:
                r[k] = r[k] == 1

        dt_win_cols = ['StartDateUTC', 'EndDateUTC']
        for k in dt_win_cols:
            if k in r:
                r[k] = dt_winminutes(r[k])

        dt_ts_cols = [
            'ModDate', 'DueDate', 'StartDate', 'TimeReceived', 'TimeSent'
            ]
        for k in dt_ts_cols:
            if k in r:
                r[k] = datetime.fromtimestamp(r[k]).replace(tzinfo=self.localtime)

        return r

    def _mail_query(self):
        return ('Mail', """
            SELECT m.PathToDataFile, m.Record_RecordID AS RecordID,
                 m.Record_FolderID AS FolderID,
                 Record_AccountUID AS AccountUID,
                 Record_ModDate AS ModDate,
                 Message_type AS MessageType,
                 Message_HasAttachment AS HasAttachment,
                 Message_Hidden AS Hidden,
                 Message_ImapUID AS IMAPUID,
                 Message_IsOutgoingMessage AS IsOutgoingMessage,
                 Message_MarkedForDelete AS MarkedForDelete,
                 Message_MentionedMe AS MentionedMe,
                 Message_MessageID AS MessageID,
                 Message_NormalizedSubject AS NormalizedSubject,
                 Message_PartiallyDownloaded AS PartiallyDownloaded,
                 Message_DownloadState AS DownloadState,
                 Message_ReadFlag AS ReadFlag,
                 Message_RecipientList AS RecipientList,
                 Message_DisplayTo AS DisplayTo,
                 Message_Preview AS Preview,
                 Message_SenderList AS SenderList,
                 Message_Sent AS Sent,
                 Message_Size AS Size,
                 Message_Status AS Status,
                 Message_SuppressAutoBackfill AS SuppressAutoBackfill,
                 Conversation_ConversationID AS ConversationID,
                 Message_ThreadTopic AS ThreadTopic,
                 Message_TimeReceived AS TimeReceived,
                 Message_TimeSent AS TimeSent,
                 Record_DueDate AS DueDate,
                 Record_ExchangeOrEasId AS ExchangeID,
                 Record_ExchangeChangeKey AS ExchangeChangeKey,
                 Record_FlagStatus AS FlagStatus,
                 Record_Priority AS Priority,
                 Record_HasReminder AS HasReminder,
                 Message_InferenceClassification AS InferenceClassification,
                 c.Category_RecordID AS CategoryID
            FROM Mail m
              LEFT JOIN Mail_Categories c
                ON m.Record_RecordID = c.Record_RecordID""")

    def _calendar_event_query(self):
        return ('CalendarEvents', """
            SELECT e.PathToDataFile, e.Record_RecordID AS RecordID,
                 e.Record_FolderID AS FolderID,
                 Record_AccountUID AS AccountUID,
                 Record_ModDate AS ModDate,
                 Calendar_StartDateUTC AS StartDateUTC,
                 Calendar_EndDateUTC AS EndDateUTC,
                 Calendar_IsRecurring AS IsRecurring,
                 Calendar_RecurrenceID AS RecurrenceID,
                 Calendar_AttendeeCount AS AttendeeCount,
                 Calendar_AllowNewTimeProposal AS AllowNewTimeProposal,
                 Record_UUID AS UUID,
                 Calendar_HasReminder AS HasReminder,
                 Calendar_IsRecurring AS IsRecurring,
                 Calendar_MasterRecordID AS MasterRecordID,
                 Record_ExchangeOrEasId AS ExchangeID,
                 Record_ExchangeChangeKey AS ExchangeChangeKey,
                 c.Category_RecordID AS CategoryID
            FROM CalendarEvents e
              LEFT JOIN CalendarEvents_Categories c
                ON e.Record_RecordID = c.Record_RecordID""")

    def _folder_query(self):
        return ('Folders', """
            SELECT PathToDataFile,
                 Record_RecordID AS RecordID,
                 Record_ModDate AS ModDate,
                 Record_AccountUID AS AccountUID,
                 Folder_ParentID AS ParentID,
                 Folder_FolderClass AS FolderClass,
                 Folder_FolderType AS FolderType,
                 Folder_SpecialFolderType AS SpecialFolderType,
                 Folder_Name AS Name,
                 Folder_ContainsPartialDwnldMsgs AS ContainsPartialDwnldMsgs,
                 Record_ExchangeOrEasId AS ExchangeID,
                 Record_ExchangeChangeKey AS ExchangeChangeKey,
                 Folder_OnlineFolderType AS OnlineFolderType,
                 Folder_SubFolderSyncMapReset AS SubFolderSyncMapReset,
                 Folder_SyncMapReset AS SyncMapReset,
                 Folder_IgnoreReminders AS IgnoreReminders
            FROM Folders""")

    def _task_query(self):
        return ('Tasks', """
            SELECT t.PathToDataFile, t.Record_RecordID AS RecordID,
                 Record_ModDate AS ModDate,
                 t.Record_FolderID AS FolderID,
                 Record_AccountUID AS AccountUID,
                 Task_Completed AS Completed,
                 Record_DueDate AS DueDate,
                 Record_ExchangeOrEasId AS ExchangeID,
                 Record_ExchangeChangeKey AS ExchangeChangeKey,
                 Record_StartDate AS StartDate,
                 Record_HasReminder AS HasReminder,
                 Record_UUID AS UUID,
                 Task_Name AS Name,
                 c.Category_RecordID AS CategoryID
            FROM Tasks t
              LEFT JOIN Tasks_Categories c
                ON t.Record_RecordID = c.Record_RecordID""")

    def _note_query(self):
        return ('Notes', """
            SELECT n.PathToDataFile, n.Record_RecordID AS RecordID,
                 Record_ModDate AS ModDate,
                 n.Record_FolderID AS FolderID,
                 Record_AccountUID AS AccountUID,
                 Record_ExchangeOrEasId AS ExchangeID,
                 Record_ExchangeChangeKey AS ExchangeChangeKey,
                 Record_UUID AS UUID,
                 Note_Title AS Title,
                 c.Category_RecordID AS CategoryID
            FROM Notes n
              LEFT JOIN Notes_Categories c
                ON n.Record_RecordID = c.Record_RecordID""")

    def _contact_query(self):
        return ('Contacts', """
            SELECT c.PathToDataFile, c.Record_RecordID AS RecordID,
                 Record_ModDate AS ModDate,
                 c.Record_FolderID AS FolderID,
                 Record_AccountUID AS AccountUID,
                 Contact_ContactRecType AS ContactRecType,
                 Contact_DisplayName AS DisplayName,
                 Contact_DisplayNameSort AS DisplayNameSort,
                 Contact_LanguageID AS LanguageID,
                 Record_DueDate AS DueDate,
                 Record_ExchangeOrEasId AS ExchangeID,
                 Record_ExchangeChangeKey AS ExchangeChangeKey,
                 Record_FlagStatus AS FlagStatus,
                 Record_StartDate AS StartDate,
                 Record_UUID AS UUID,
                 Record_HasReminder AS HasReminder,
                 cat.Category_RecordID AS CategoryID
            FROM Contacts c
              LEFT JOIN Contacts_Categories cat
                ON c.Record_RecordID = cat.Record_RecordID""")

    def _category_query(self):
        return ('Categories', """
            SELECT PathToDataFile, Record_RecordID AS RecordID,
                Record_AccountUID AS AccountUID,
                Category_Name AS Name,
                Category_Exchange_IsLocalCategory AS IsLocalCategory,
                Cateogry_ExchangeGuid AS ExchangeGuid,
                Category_BackgroundColor AS BackgroundColor
            FROM Categories""")

    def _signature_query(self):
        return ('Signatures',
            "SELECT PathToDataFile, Record_RecordID AS RecordID FROM Signatures"
            )

    def _search_query(self):
        return ('SavedSpotlightSearch',
            "SELECT PathToDataFile, Record_RecordID AS RecordID FROM SavedSpotlightSearch"
            )

    def _main_query(self):
        return ('Main',
            "SELECT PathToDataFile, Record_RecordID AS RecordID FROM Main"
            )

    def _acctmail_query(self):
        return ('AccountsMail', """
            SELECT PathToDataFile, Record_RecordID AS RecordID,
                Account_AssociatedAccountOfUID AS AssociatedAccountOfUID,
                Account_Name AS Name,
                Account_EmailAddress AS EmailAddress,
                Account_DeviceGuid AS DeviceGuid,
                Account_ServerType AS ServerType
            FROM AccountsMail""")

    def _acctexch_query(self):
        return ('AccountsExchange', """
            SELECT PathToDataFile, Record_RecordID AS RecordID,
                Account_AssociatedAccountOfUID AS AssociatedAccountOfUID,
                Account_LdapAccountUID AS LDAPAccountUID,
                Account_MailAccountUID AS MailAccountUID,
                Account_Name AS Name,
                Account_EmailAddress AS EmailAddress
            FROM AccountsExchange""")

    def _block_query(self, table):
        return f"""
            SELECT PathToDataFile
            FROM Blocks b
              JOIN {table}_OwnedBlocks ob ON b.BlockTag = ob.BlockTag
                AND b.BlockID = ob.BlockID
            WHERE ob.Record_RecordID = ?"""

    
    ### EXPORT ###
    def export(self, path='Recovered Outlook Data'):
        # Make folders
        if path:
            if not os.path.exists(path):
                os.mkdir(path)
            os.chdir(path)
        paths = self._build_folders()
        os.makedirs('Categories', exist_ok=True)
        os.makedirs('SavedSearches', exist_ok=True)
        os.makedirs('Signatures', exist_ok=True)

        # Write files
        _ = [export(x, '') for x in self.Mains.values()]
        _ = [export(x, '') for x in self.AccountsExchange.values()]
        _ = [export(x, '') for x in self.AccountsMail.values()]
        _ = [export(x, 'Categories') for x in self.Categories.values()]
        _ = [export(x, 'SavedSearches') for x in self.SavedSearches.values()]
        _ = [export(x, 'Signatures') for x in self.Signatures.values()]
        _ = [export(x, paths[x.RecordID]) for x in self.Folders.values()]
        _ = [export(x, paths[x.FolderID]) for x in self.Notes.values()]
        _ = [export(x, paths[x.FolderID]) for x in self.Events.values()]
        _ = [export(x, paths[x.FolderID]) for x in self.Messages.values()]

    def _build_folders(self):
        # Get paths from folder structure
        parents = {f.RecordID: f.ParentID for f in self.Folders.values()}
        roots = set(f for f in parents.values() if f not in set(parents.keys()))
        paths = dict()
        for node in parents.keys():
            path = [node]
            while path[-1] not in roots:
                path.append(parents[path[-1]])
            path = reversed(path)
            names = list()
            for i in path:
                if i in roots:
                    continue
                else:
                    fldr = self.Folders[i]
                    names.append((fldr.Name or str(fldr.RecordID)).replace('/', ''))
            names = '/'.join(names)
            paths[node] = names

        # Make folders
        for path in paths.values():
            os.makedirs(path, exist_ok=True)

        return paths
