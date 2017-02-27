namespace Rally
{
    #region: System Libraries
    using System;
    using System.IO;
    using RestApi;
    using RestApi.Json;
    using RestApi.Response;
    using System.Collections.Generic;
    using System.Net;
    using ActiveUp.Net.Mail;
    using System.Drawing;
    #endregion

    class Sync
    {
        private RallyRestApi _rallyApi;
        private Imap4Client _imap4Client;

        public string RallyUserName { get; set; }
        public string RallyPassword { get; set; }
        public string OutlookUserName { get; set; }
        public string OutlookPassword { get; set; }

        public Sync(string rallyUserName, string rallyPassword, string outlookUserName, string outlookPassword)
        {
            _rallyApi = new RallyRestApi();
            _imap4Client = new Imap4Client();
            this.OutlookUserName = outlookUserName;
            this.OutlookPassword = outlookPassword;
            this.RallyUserName = rallyUserName;
            this.RallyPassword = rallyPassword;
        }

        private Mailbox _inbox;
        private int[] _unreadMsg;
        private List<Message> _unreadMsgCollection = new List<Message>();
        private FlagCollection _markAsUnreadFlag;
        private Dictionary<string, string> _attachmentsDictionary = new Dictionary<string, string>();
        private DynamicJsonObject _toCreate = new DynamicJsonObject();
        private DynamicJsonObject _attachmentContent = new DynamicJsonObject();
        private DynamicJsonObject _attachmentContainer = new DynamicJsonObject();
        private CreateResult _createUserStory;
        private CreateResult _attachmentContentCreateResult;
        private CreateResult _attachmentContainerCreateResult;
        private string _base64String;
        private string _attachmentFileName;
        private string[] _attachmentPaths;
        private string _userStoryReference;

        /// <summary>
        /// Authenticate with Outlook with valid credentials.
        /// </summary>
        private void LoginToOutlook()
        {
            _imap4Client.ConnectSsl(OutlookConstant.OutlookHost, OutlookConstant.OutlookPort);
            _imap4Client.Login(this.OutlookUserName, this.OutlookPassword);
        }

        /// <summary>
        /// Authenticate with Rally, with valid credentials.
        /// </summary>
        private void LoginToRally()
        {
            if (this._rallyApi.AuthenticationState != RallyRestApi.AuthenticationResult.Authenticated)
            {
                _rallyApi.Authenticate(this.RallyUserName, this.RallyPassword, RallyConstant.ServerId, null, RallyConstant.AllowSso);
            }
        }

        /// <summary>
        /// Fetches all the unread email objects and populates a List with the collection
        /// </summary>
        /// <param name="unread"></param>
        /// <param name="inbox"></param>
        private void FetchUnreadMessages(int[] unread, Mailbox inbox)
        {
            for (int i = 0; i < unread.Length; i++)
            {
                Message msg = inbox.Fetch.MessageObject(unread[i]);
                _unreadMsgCollection.Add(msg);
            }
        }

        /// <summary>
        /// Converts each attachment from an unread email object to base 64
        /// This allows for the unqiue string to be shipped over a network
        /// </summary>
        /// <param name="attachment"></param>
        /// <returns>base64EncodedString</returns>
        private static string FileToBase64(string attachment)
        {
            Byte[] attachmentBytes = File.ReadAllBytes(attachment);
            string base64EncodedString = Convert.ToBase64String(attachmentBytes);
            return base64EncodedString;
        }

        /// <summary>
        /// Populates the Dictionary with a unique base64 string along with the respective file name
        /// Dictionary prevents duplicate attachments from being uploaded to Rally if two attachments have the same base64String
        /// </summary>
        private void PopulateAttachmentsDictionary()
        {
            _attachmentPaths = Directory.GetFiles(SyncConstant.AttachmentsDirectory);

            foreach (var file in _attachmentPaths)
            {
                _base64String = FileToBase64(file);
                _attachmentFileName = Path.GetFileName(file);
                var fileName = string.Empty;

                if (!(_attachmentsDictionary.TryGetValue(_base64String, out fileName)))
                {
                    _attachmentsDictionary.Add(_base64String, _attachmentFileName);
                }

                Console.WriteLine("Uploading: " + file);
                File.Delete(file);
            }
        }

        /// <summary>
        /// Iterates over the populated Dictionary object and pushes each attachment to the respective user story
        /// </summary>
        /// <param name="attachmentsDictionary"></param>
        /// <param name="attachmentContent"></param>
        /// <param name="attachmentContainer"></param>
        /// <param name="createUserStory"></param>
        private void PushAttachments(Dictionary<string, string> attachmentsDictionary, DynamicJsonObject attachmentContent, DynamicJsonObject attachmentContainer, CreateResult createUserStory)
        {
            foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
            {
                try
                {
                    //create attachment content
                    attachmentContent[RallyConstant.Content] = attachmentPair.Key;
                    _attachmentContentCreateResult = _rallyApi.Create(RallyConstant.AttachmentContent, attachmentContent);
                    _userStoryReference = _attachmentContentCreateResult.Reference;

                    //create attachment contianer
                    attachmentContainer[RallyConstant.Artifact] = createUserStory.Reference;
                    attachmentContainer[RallyConstant.Content] = _userStoryReference;
                    attachmentContainer[RallyConstant.Name] = attachmentPair.Value;
                    attachmentContainer[RallyConstant.Description] = RallyConstant.EmailAttachment;
                    attachmentContainer[RallyConstant.ContentType] = SyncConstant.FileType;

                    //Create & associate the attachment to the respecitve user story
                    _attachmentContainerCreateResult = _rallyApi.Create(RallyConstant.Attachment, attachmentContainer);
                }
                catch (WebException e)
                {
                    Console.WriteLine("Attachment Failed: " + e.Message);
                }
            }
        }

        /// <summary>
        /// Email objects will be marked as unread and moved to a different folder for the next iteration 
        /// </summary>
        /// <param name="unread"></param>
        /// <param name="markAsUnreadFlag"></param>
        /// <param name="inbox"></param>
        private static void MarkAsUnread(int[] unread, FlagCollection markAsUnreadFlag, Mailbox inbox)
        {
            foreach (var item in unread)
            {
                markAsUnreadFlag.Add(OutlookConstant.OutlookSeenMessages);
                inbox.RemoveFlags(item, markAsUnreadFlag);
                inbox.MoveMessage(item, OutlookConstant.OutlookProcessedFolder);
            }
        }

        /// <summary>
        /// Returns a length of unread messages
        /// </summary>
        /// <param name="unread"></param>
        /// <returns>int</returns>
        private int UnreadMessageLength(int[] unread)
        {
            return unread.Length;
        }

        /// <summary>
        /// Parses unread email objects, and creates user stories with attachments from the data provided in an email object
        /// </summary>
        /// <param name="workspace"></param>
        /// <param name="project"></param>
        public void SyncUserStories(string workspace, string project)
        {
            _unreadMsgCollection.Capacity = 25;
            _toCreate[RallyConstant.WorkSpace] = workspace;
            _toCreate[RallyConstant.Project] = project;

            try
            {
                LoginToOutlook();
                LoginToRally();

                _inbox = _imap4Client.SelectMailbox(OutlookConstant.OutlookInboxFolder);
                _unreadMsg = _inbox.Search(OutlookConstant.OutlookUnseenMessages);
                _markAsUnreadFlag = new FlagCollection();

                if (UnreadMessageLength(_unreadMsg) > 0)
                {
                    Console.WriteLine("Syncing: " + _unreadMsg.Length + " Unread Messages");
                    FetchUnreadMessages(_unreadMsg, _inbox);

                    for (int i = 0; i < _unreadMsgCollection.Count; i++)
                    {
                        //If the email object does not have a subject
                        if (string.IsNullOrWhiteSpace(_unreadMsgCollection[i].Subject))
                        {
                            _unreadMsgCollection[i].Subject = OutlookConstant.NoSubject;
                        }

                        _toCreate[RallyConstant.Name] = (_unreadMsgCollection[i].Subject);
                        _toCreate[RallyConstant.Description] = (_unreadMsgCollection[i].BodyText.Text);
                        _toCreate[RallyConstant.PortfolioItem] = RallyQueryConstant.FeatureShareProject;
                        _createUserStory = _rallyApi.Create(RallyConstant.HierarchicalRequirement, _toCreate);

                        if (_unreadMsgCollection[i].Attachments.Count > 0) //can return a boolean
                        {
                            _unreadMsgCollection[i].Attachments.StoreToFolder(SyncConstant.AttachmentsDirectory);
                        }

                        PopulateAttachmentsDictionary();
                        PushAttachments(_attachmentsDictionary, _attachmentContent, _attachmentContainer, _createUserStory);
                        _attachmentsDictionary.Clear();
                    }

                    MarkAsUnread(_unreadMsg, _markAsUnreadFlag, _inbox);
                    Console.WriteLine("Created " + _unreadMsg.Length + " User Stories");
                }
                else
                {
                    Console.WriteLine("Inbox does not contain unread messages");
                }
            }
            catch (Imap4Exception imap)
            {
                Console.WriteLine(imap.Message);
            }
            catch (IOException io)
            {
                Console.WriteLine(io.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                _imap4Client.Disconnect();
            }
        }

    }
}
