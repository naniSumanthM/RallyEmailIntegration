using System.Linq;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using Rally.RestApi.Exceptions;
using ServiceStack;
using Slack.Webhooks;

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

    class RallyIntegrationUsingMimeKit
    {
        private RallyRestApi _rallyRestApi;
        private ImapClient _imapClient;
        private SlackClient _slackClient;
        private IMailFolder _inboxFolder;
        private IMailFolder _processedFolder;
        private IList<UniqueId> _emailMessageIdsList;
        private MimeMessage _message;
        private DynamicJsonObject _toCreate = new DynamicJsonObject();
        private DynamicJsonObject _attachmentContent = new DynamicJsonObject();
        private DynamicJsonObject _attachmentContainer = new DynamicJsonObject();
        private CreateResult _createUserStory;
        private CreateResult _attachmentContentCreateResult;
        private Dictionary<string, string> _attachmentsDictionary;
        private string _userStoryReference;
        private string _emailSubject;
        private string _emailBody;
        private string[] _allAttachments;
        private string _base64String;
        private string _attachmentFileName;
        private string _attachmentFileNameForDictionary;
        private string _objectId;
        private string _userStoryUrl;
        private string _slackAttachmentString;
        private int _duplicateFileCount = 0;
        private int _unreadMessages = 0;
        public string RallyUserName { get; set; }
        public string RallyPassword { get; set; }
        public string GmailUserName { get; set; }
        public string GmailPassword { get; set; }

        /// <summary>
        /// Constructor initializes instances for Rally, Gmail, and Slack.
        /// Sets values to authenticate with Rally and Gmail
        /// </summary>
        /// <param name="rallyUserName"></param>
        /// <param name="rallyPassword"></param>
        /// <param name="gmailUserName"></param>
        /// <param name="gmailPassword"></param>
        public RallyIntegrationUsingMimeKit(string rallyUserName, string rallyPassword, string gmailUserName, string gmailPassword)
        {
            _rallyRestApi = new RallyRestApi();
            _imapClient = new ImapClient();
            _slackClient = new SlackClient(RallyConstant.SlackApiToken, 100);
            this.GmailUserName = gmailUserName;
            this.GmailPassword = gmailPassword;
            this.RallyUserName = rallyUserName;
            this.RallyPassword = rallyPassword;
        }

        /// <summary>
        /// Authenticate with Rally server if not authenticated
        /// </summary>
        private void LoginToRally()
        {
            if (this._rallyRestApi.AuthenticationState != RallyRestApi.AuthenticationResult.Authenticated)
            {
                _rallyRestApi.Authenticate(this.RallyUserName, this.RallyPassword, RallyConstant.ServerId, null,
                    RallyConstant.AllowSso);
            }
        }

        /// <summary>
        /// Authenticate with Gmail if not authenticated
        /// </summary>
        /// <param name="client"></param>
        private void LoginToGmail(ImapClient client)
        {
            if (!this._imapClient.IsAuthenticated)
            {
                client.ServerCertificateValidationCallback = (s, c, ch, e) => true;
                client.Connect(EmailConstant.GoogleHost, EmailConstant.ImapPort, SecureSocketOptions.SslOnConnect);
                client.AuthenticationMechanisms.Remove(EmailConstant.GoogleOAuth);
                client.Authenticate(EmailConstant.GoogleUsername, EmailConstant.GenericPassword);
            }
        }

        /// <summary>
        /// When passed an ImapClient object, gets a refrence to the "Inbox" folder along with the count of unread messages
        /// </summary>
        /// <param name="client"></param>
        private void SetUpMailbox(ImapClient client)
        {
            client.Inbox.Open(FolderAccess.ReadWrite);
            _inboxFolder = client.GetFolder(EmailConstant.GmailInbox);
            _emailMessageIdsList = client.Inbox.Search(SearchQuery.NotSeen);
            _unreadMessages = _emailMessageIdsList.Count;
        }

        /// <summary>
        /// When passed an id of email message, method will create the user story with subject as user story title, and body as the description
        /// </summary>
        /// <param name="messageId"></param>
        private void CreateUserStoryWithEmail(UniqueId messageId)
        {
            _message = _inboxFolder.GetMessage(messageId);
            _emailSubject = _message.Subject;
            _emailBody = _message.TextBody;

            if (_emailSubject.IsEmpty())
            {
                _emailSubject = EmailConstant.NoSubject;
            }

            _toCreate[RallyConstant.Name] = (_emailSubject);
            _toCreate[RallyConstant.Description] = (_emailBody);
            _toCreate[RallyConstant.PortfolioItem] = RallyQueryConstant.FeatureShareProject;
            _createUserStory = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, _toCreate);

            Console.WriteLine("Created User Story: " + _emailSubject);
        }

        /// <summary>
        /// When passed the file path of an attachment, method will convert the file to a base 64 string
        /// </summary>
        /// <param name="attachment"></param>
        /// <returns></returns>
        private string FileToBase64(string attachment)
        {
            Byte[] attachmentBytes = File.ReadAllBytes(attachment);
            string convertToBase64 = Convert.ToBase64String(attachmentBytes);
            return convertToBase64;
        }

        /// <summary>
        /// Download all attachments (regular and embedded) to a local directory
        /// 
        /// </summary>
        /// <param name="message"></param>
        private void DownloadAttachments(MimeMessage message)
        {
            if (message.BodyParts.Count() > 0)
            {
                foreach (MimeEntity attachment in message.BodyParts)
                {
                    string attachmentFile = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                    string attachmentFilePath = String.Concat(StorageConstant.MimeKitAttachmentsDirectoryWork,
                        Path.GetFileName(attachmentFile));

                    if (!string.IsNullOrWhiteSpace(attachmentFile))
                    {
                        if (File.Exists(attachmentFilePath))
                        {
                            string extension = Path.GetExtension(attachmentFilePath);
                            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(attachmentFilePath);
                            attachmentFile = string.Format(fileNameWithoutExtension + "-{0}" + "{1}",
                                ++_duplicateFileCount, extension);
                            attachmentFilePath = Path.Combine(StorageConstant.MimeKitAttachmentsDirectoryWork,
                                attachmentFile);
                        }

                        using (FileStream attachmentStream = File.Create(attachmentFilePath))
                        {
                            MimeKit.MimePart part = (MimeKit.MimePart)attachment;
                            part.ContentObject.DecodeTo(attachmentStream);
                        }

                        Console.WriteLine("Downloaded: " + attachmentFile);
                    }
                }
                _duplicateFileCount = 0;
            }
        }

        /// <summary>
        /// Converts each file into base 64, and add the key-value pair of the 64BitString, fileName to the Dictionary.
        /// Need to change attachments directory accroding to enviornment
        /// </summary>
        private void ProcessAttachments()
        {
            _attachmentsDictionary = new Dictionary<string, string>();
            _allAttachments = Directory.GetFiles(StorageConstant.MimeKitAttachmentsDirectoryWork);

            foreach (string file in _allAttachments)
            {
                _base64String = FileToBase64(file);
                _attachmentFileName = Path.GetFileName(file);
                _attachmentFileNameForDictionary = string.Empty;

                if (!_attachmentsDictionary.TryGetValue(_base64String, out _attachmentFileNameForDictionary))
                {
                    _attachmentsDictionary.Add(_base64String, _attachmentFileName);
                    Console.WriteLine("Accepted: " + file);
                }
                else
                {
                    Console.WriteLine("Omitting Duplicate: " + file);
                }

                File.Delete(file);
            }
        }

        /// <summary>
        /// With a populated Dictionary, method will iterate over the collection and upload each attachment to the respective user story
        /// </summary>
        private void UploadAttachmentsToRallyUserStory()
        {
            foreach (KeyValuePair<string, string> attachmentPair in _attachmentsDictionary)
            {
                try
                {
                    _attachmentContent[RallyConstant.Content] = attachmentPair.Key;
                    _attachmentContentCreateResult = _rallyRestApi.Create(RallyConstant.AttachmentContent, _attachmentContent);
                    _userStoryReference = _attachmentContentCreateResult.Reference;
                    _attachmentContainer[RallyConstant.Artifact] = _createUserStory.Reference;
                    _attachmentContainer[RallyConstant.Content] = _userStoryReference;
                    _attachmentContainer[RallyConstant.Name] = attachmentPair.Value;
                    _attachmentContainer[RallyConstant.Description] = RallyConstant.EmailAttachment;
                    _attachmentContainer[RallyConstant.ContentType] = StorageConstant.FileType;
                    _rallyRestApi.Create(RallyConstant.Attachment, _attachmentContainer);
                }
                catch (RallyUnavailableException)
                {
                    throw new WebException();
                }
            }
            _attachmentsDictionary.Clear();
        }

        /// <summary>
        /// Post userstory notification to Slack upon successful creation of user story
        /// </summary>
        private void PostSlackUserStoryNotification()
        {
            _objectId = Ref.GetOidFromRef(_createUserStory.Reference);
            _userStoryUrl = String.Concat(RallyConstant.UserStoryUrlFormat, _objectId);
            _slackAttachmentString = String.Format("User Story: <{0} | {1} >", _userStoryUrl, _message.Subject);

            SlackMessage message = new SlackMessage
            {
                Channel = RallyConstant.SlackChannel,
                Text = RallyConstant.SlackNotificationText,
                Username = RallyConstant.SlackUser
            };

            SlackAttachment slackAttachment = new SlackAttachment
            {
                Fallback = _slackAttachmentString,
                Text = _slackAttachmentString,
                Color = RallyConstant.HexColor
            };

            message.Attachments = new List<SlackAttachment> { slackAttachment };
            _slackClient.Post(message);
        }

        /// <summary>
        /// Moves the collection of unread messages that have been uploaded to Rally to "Processed" folder.
        /// This helps identify processed messages in the email server
        /// </summary>
        /// <param name="messageId"></param>
        private void MoveMessagesToProcessedFolder(UniqueId messageId)
        {
            _processedFolder = _imapClient.GetFolder(EmailConstant.GmailProcessedFolder);
            _imapClient.Inbox.MoveTo(messageId, _processedFolder);
        }

        /// <summary>
        /// Parses emails with attachments to create user stories in Rally if there are unread email messages in the "Inbox" folder
        /// </summary>
        /// <param name="rallyWorkspace"></param>
        /// <param name="rallyScrumTeam"></param>
        public void SyncUsingMimeKit(string rallyWorkspace, string rallyScrumTeam)
        {
            try
            {
                LoginToRally();

                _toCreate[RallyConstant.WorkSpace] = rallyWorkspace;
                _toCreate[RallyConstant.Project] = rallyScrumTeam;

                using (_imapClient)
                {
                    LoginToGmail(_imapClient);
                    SetUpMailbox(_imapClient);

                    if (_unreadMessages > 0)
                    {
                        Console.WriteLine("Syncing-" + _emailMessageIdsList.Count + " Messages");

                        foreach (UniqueId messageId in _emailMessageIdsList)
                        {
                            CreateUserStoryWithEmail(messageId);
                            DownloadAttachments(_message);
                            ProcessAttachments();
                            UploadAttachmentsToRallyUserStory();
                            PostSlackUserStoryNotification();
                            MoveMessagesToProcessedFolder(messageId);
                        }

                        Console.WriteLine("Synced-" + _emailMessageIdsList.Count + " Messages");
                    }
                    else
                    {
                        Console.WriteLine("No Unread Messages Found");
                    }

                    _imapClient.Disconnect(true);
                }
            }
            catch (IOException io)
            {
                Console.WriteLine(io.Message);
            }
            catch (ImapProtocolException imapProtocolException)
            {
                Console.WriteLine(imapProtocolException.Message);    
            }
            catch (RallyUnavailableException rallyUnavailableException)
            {
                Console.WriteLine(rallyUnavailableException.Message);
            }
            catch(WebException webException)
            {
                Console.WriteLine(webException.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                _rallyRestApi.Logout();
            }
        }
    }
}
