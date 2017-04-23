using System.Linq;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
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
        private IList<UniqueId> _emailMessageIds;
        private MimeMessage _message;
        private DynamicJsonObject _toCreate = new DynamicJsonObject();
        private DynamicJsonObject _attachmentContent = new DynamicJsonObject();
        private DynamicJsonObject _attachmentContainer = new DynamicJsonObject();
        private CreateResult _createUserStory;
        private CreateResult _attachmentContentCreateResult;
        private CreateResult _attachmentContainerCreateResult;
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
        public string RallyUserName { get; set; }
        public string RallyPassword { get; set; }
        public string GmailUserName { get; set; }
        public string GmailPassword { get; set; }

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

        private void LoginToRally()
        {
            if (this._rallyRestApi.AuthenticationState != RallyRestApi.AuthenticationResult.Authenticated)
            {
                _rallyRestApi.Authenticate(this.RallyUserName, this.RallyPassword, RallyConstant.ServerId, null,
                    RallyConstant.AllowSso);
            }
        }

        private void LoginToGmail(ImapClient client)
        {
            client.ServerCertificateValidationCallback = (s, c, ch, e) => true;
            client.Connect(EmailConstant.GoogleHost, EmailConstant.ImapPort, SecureSocketOptions.SslOnConnect);
            client.AuthenticationMechanisms.Remove(EmailConstant.GoogleOAuth);
            client.Authenticate(EmailConstant.GoogleUsername, EmailConstant.GenericPassword);
        }

        private void SetUpMailbox(ImapClient client)
        {
            client.Inbox.Open(FolderAccess.ReadWrite);
            _inboxFolder = client.GetFolder(EmailConstant.GmailInbox);
            _emailMessageIds = client.Inbox.Search(SearchQuery.All);
        }

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

        private static string FileToBase64(string attachment)
        {
            Byte[] attachmentBytes = File.ReadAllBytes(attachment);
            string base64EncodedString = Convert.ToBase64String(attachmentBytes);
            return base64EncodedString;
        }

        private void DownloadAttachments(MimeMessage message)
        {
            if (message == null) throw new ArgumentNullException(nameof(message));

            if (message.BodyParts.Count() > 0)
            {
                foreach (MimeEntity attachment in message.BodyParts)
                {
                    string attachmentFile = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                    string attachmentFilePath = String.Concat(StorageConstant.MimeKitAttachmentsDirectory,
                        Path.GetFileName(attachmentFile));

                    if (!string.IsNullOrWhiteSpace(attachmentFile))
                    {
                        if (File.Exists(attachmentFilePath))
                        {
                            string extension = Path.GetExtension(attachmentFilePath);
                            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(attachmentFilePath);
                            attachmentFile = string.Format(fileNameWithoutExtension + "-{0}" + "{1}",
                                ++_duplicateFileCount, extension);
                            attachmentFilePath = Path.Combine(StorageConstant.MimeKitAttachmentsDirectory,
                                attachmentFile);
                        }

                        using (var attachmentStream = File.Create(attachmentFilePath))
                        {
                            MimeKit.MimePart part = (MimeKit.MimePart)attachment;
                            part.ContentObject.DecodeTo(attachmentStream);
                        }

                        Console.WriteLine("Downloaded: " + attachmentFile);
                    }
                }
                _duplicateFileCount = 0;
            }
            else
            {
                Console.WriteLine("Omitting Duplicate: " + message.Subject);
            }
        }

        private void ProcessAttachments()
        {
            _attachmentsDictionary = new Dictionary<string, string>();
            _allAttachments = Directory.GetFiles(StorageConstant.MimeKitAttachmentsDirectory);

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
                    _attachmentContainerCreateResult = _rallyRestApi.Create(RallyConstant.Attachment, _attachmentContainer);
                }
                catch (WebException)
                {
                    throw new WebException();
                }
            }
            _attachmentsDictionary.Clear();
        }

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

            var slackAttachment = new SlackAttachment
            {
                Fallback = _slackAttachmentString,
                Text = _slackAttachmentString,
                Color = RallyConstant.HexColor
            };

            message.Attachments = new List<SlackAttachment> { slackAttachment };
            _slackClient.Post(message);
        }

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

                    foreach (UniqueId messageId in _emailMessageIds)
                    {
                        CreateUserStoryWithEmail(messageId);
                        DownloadAttachments(_message);
                        ProcessAttachments();
                        UploadAttachmentsToRallyUserStory();
                        PostSlackUserStoryNotification();
                    }
                }
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
                _rallyRestApi.Logout();
            }
        }
    }
}
