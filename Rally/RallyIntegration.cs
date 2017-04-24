namespace Rally
{
    #region: Libraries

    using System;
    using System.IO;
    using RestApi;
    using RestApi.Json;
    using RestApi.Response;
    using System.Collections.Generic;
    using System.Net;
    using ActiveUp.Net.Mail;
    using Slack.Webhooks;

    #endregion

    class RallyIntegration
    {
        private RallyRestApi _rallyApi;
        private Imap4Client _imap4Client;
        private SlackClient _slackClient;
        private Mailbox _inbox;
        private int[] _unreadMsg;
        private FlagCollection _markAsUnreadFlag;
        private List<Message> _unreadMsgCollection = new List<Message>();
        private Dictionary<string, string> _attachmentsDictionary = new Dictionary<string, string>();
        private DynamicJsonObject _toCreate = new DynamicJsonObject();
        private DynamicJsonObject _attachmentContent = new DynamicJsonObject();
        private DynamicJsonObject _attachmentContainer = new DynamicJsonObject();
        private CreateResult _createUserStory;
        private CreateResult _attachmentContentCreateResult;
        private CreateResult _attachmentContainerCreateResult;
        private string _base64String;
        private string _attachmentFileName;
        private string _userStoryReference;
        private string _inlineFileName;
        private string[] _attachmentPaths;
        private string[] _inlineAttachmentPaths;
        private byte[] _inlineFileBinaryContent;
        private string _objectId;
        private string _userStoryUrl;
        private string _slackAttachmentString;

        public string RallyUserName { get; set; }
        public string RallyPassword { get; set; }
        public string GmailUserName { get; set; }
        public string GmailPassword { get; set; }

        /// <summary>
        /// Default Constrcutor will authenticate with Rally, Outlook, Slack
        /// </summary>
        /// <param name="rallyUserName"></param>
        /// <param name="rallyPassword"></param>
        /// <param name="gmailUserName"></param>
        /// <param name="gmailPassword"></param>
        public RallyIntegration(string rallyUserName, string rallyPassword, string gmailUserName, string gmailPassword)
        {
            _rallyApi = new RallyRestApi();
            _imap4Client = new Imap4Client();
            _slackClient = new SlackClient(RallyConstant.SlackApiToken, 100);
            this.GmailUserName = gmailUserName;
            this.GmailPassword = gmailPassword;
            this.RallyUserName = rallyUserName;
            this.RallyPassword = rallyPassword;
        }

        /// <summary>
        /// Authenticate with Google Mail
        /// </summary>
        private void LoginToGmail()
        {
            _imap4Client.ConnectSsl(EmailConstant.GoogleHost, EmailConstant.ImapPort);
            _imap4Client.Login(EmailConstant.GoogleUsername, EmailConstant.GenericPassword);
        }

        /// <summary>
        /// Authenticate with Rally, with valid credentials.
        /// </summary>
        private void LoginToRally()
        {
            if (this._rallyApi.AuthenticationState != RallyRestApi.AuthenticationResult.Authenticated)
            {
                _rallyApi.Authenticate(this.RallyUserName, this.RallyPassword, RallyConstant.ServerId, null,
                    RallyConstant.AllowSso);
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
            _attachmentPaths = Directory.GetFiles(StorageConstant.AttachmentsDirectory);

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
        private void UploadAttachmentsToRally(Dictionary<string, string> attachmentsDictionary,
            DynamicJsonObject attachmentContent, DynamicJsonObject attachmentContainer, CreateResult createUserStory)
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
                    attachmentContainer[RallyConstant.ContentType] = StorageConstant.FileType;

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
                markAsUnreadFlag.Add(EmailConstant.SeenMessages);
                inbox.RemoveFlags(item, markAsUnreadFlag);
            }
        }

        /// <summary>
        /// Email objects need to be moved to the "PROCESSED" folder
        /// </summary>
        /// <param name="unread"></param>
        /// <param name="markAsUnreadFlag"></param>
        /// <param name="inbox"></param>
        private static void MoveMessage(int[] unread, FlagCollection markAsUnreadFlag, Mailbox inbox)
        {
            foreach (var item in unread)
            {
                inbox.MoveMessage(item, EmailConstant.OutloookProcessedFolder);
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
        /// Inline attachments are downloaded and written to a directory on home "inlineAttachments"
        /// </summary>
        /// <param name="embeddedImg"></param>
        private void DownloadInlineAttachments(MimePart embeddedImg)
        {
            _inlineFileName = embeddedImg.ContentName;
            _inlineFileBinaryContent = embeddedImg.BinaryContent;
            File.WriteAllBytes(Path.Combine(StorageConstant.InlineImageDirectory, _inlineFileName),
                _inlineFileBinaryContent);
        }

        ///<summary>
        /// Parses each inline image attached in the email and populates the dictionary
        /// </summary>
        private void PopulateInlineAttachments()
        {
            _inlineAttachmentPaths = Directory.GetFiles(StorageConstant.InlineImageDirectory);

            foreach (var file in _inlineAttachmentPaths)
            {
                //convert to base64 String
                string base64String = FileToBase64(file);
                string attachmentFileName = Path.GetFileName(file);
                var emptyFileString = string.Empty;

                Console.WriteLine("Adding to Dictionary: " + attachmentFileName);

                if (!(_attachmentsDictionary.TryGetValue(base64String, out _inlineFileName)))
                {
                    _attachmentsDictionary.Add(base64String, attachmentFileName);
                }

                File.Delete(file);
            }
        }

        /// <summary>
        /// Processes Each Inline Image attached within an email
        /// </summary>
        /// <param name="i"></param>
        private void ProcessInlineAttachments(int i)
        {
            foreach (MimePart embeddedImg in _unreadMsgCollection[i].EmbeddedObjects)
            {
                DownloadInlineAttachments(embeddedImg);
                PopulateInlineAttachments();
                UploadAttachmentsToRally(_attachmentsDictionary, _attachmentContent, _attachmentContainer, _createUserStory);
                _attachmentsDictionary.Clear();
            }
        }

        /// <summary>
        /// Pushes a notification into Slack for each user story created
        /// </summary>
        /// <param name="i"></param>
        private void PushSlackNotification(int i)
        {
            _objectId = Ref.GetOidFromRef(_createUserStory.Reference);
            _userStoryUrl = String.Concat(RallyConstant.UserStoryUrlFormat, _objectId);
            _slackAttachmentString = String.Format("User Story: <{0} | {1} >", _userStoryUrl,
                _unreadMsgCollection[i].Subject);

            SlackMessage message = new SlackMessage
            {
                Channel = RallyConstant.SlackChannel,
                Text = RallyConstant.SlackNotificationText,
                Username = RallyConstant.SlackUser
            };

            var slackAttachment = new SlackAttachment
            {
                Fallback = _slackAttachmentString,
                Text = _slackAttachmentString,
                Color = RallyConstant.HexColor
            };

            message.Attachments = new List<SlackAttachment> {slackAttachment};
            _slackClient.Post(message);
        }

        /// <summary>
        /// Creates user stories with RallyFeature, Attachments, Description by parsing unread email objects.
        /// </summary>
        /// <param name="workspace"></param>
        /// <param name="project"></param>
        public void SyncUserStories(string workspace, string project)
        {
            _unreadMsgCollection.Capacity = 50;
            _toCreate[RallyConstant.WorkSpace] = workspace;
            _toCreate[RallyConstant.Project] = project;

            try
            {
                LoginToGmail();
                LoginToRally();

                _inbox = _imap4Client.SelectMailbox(EmailConstant.OutlookInboxFolder);
                _unreadMsg = _inbox.Search(EmailConstant.UnseenMessages);
                _markAsUnreadFlag = new FlagCollection();

                if (UnreadMessageLength(_unreadMsg) > 0)
                {
                    Console.WriteLine("Syncing: " + _unreadMsg.Length + " Unread Messages");
                    FetchUnreadMessages(_unreadMsg, _inbox);

                    for (int i = 0; i < _unreadMsgCollection.Count; i++)
                    {
                        if (string.IsNullOrWhiteSpace(_unreadMsgCollection[i].Subject))
                        {
                            _unreadMsgCollection[i].Subject = EmailConstant.NoSubject;
                        }

                        _toCreate[RallyConstant.Name] = (_unreadMsgCollection[i].Subject);
                        _toCreate[RallyConstant.Description] = (_unreadMsgCollection[i].BodyText.Text);
                        _toCreate[RallyConstant.PortfolioItem] = RallyQueryConstant.FeatureShareProject;
                        _createUserStory = _rallyApi.Create(RallyConstant.HierarchicalRequirement, _toCreate);

                        if (_unreadMsgCollection[i].Attachments.Count > 0)
                        {
                            _unreadMsgCollection[i].Attachments.StoreToFolder(StorageConstant.AttachmentsDirectory);
                        }

                        if (_unreadMsgCollection[i].EmbeddedObjects.Count > 0)
                        {
                            ProcessInlineAttachments(i);
                        }

                        PopulateAttachmentsDictionary();
                        UploadAttachmentsToRally(_attachmentsDictionary, _attachmentContent, _attachmentContainer, _createUserStory);
                        _attachmentsDictionary.Clear();
                        PushSlackNotification(i);
                    }

                    Console.WriteLine("Created " + _unreadMsg.Length + " User Stories");
                }
                else
                {
                    Console.WriteLine("No Unread Messages Found");
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
            catch (WebException we)
            {
                Console.WriteLine(we.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                _imap4Client.Disconnect();
                _rallyApi.Logout();
            }
        }
    }
}
