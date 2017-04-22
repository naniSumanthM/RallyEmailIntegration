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
        private DynamicJsonObject toCreate = new DynamicJsonObject();
        private DynamicJsonObject attachmentContent = new DynamicJsonObject();
        private DynamicJsonObject attachmentContainer = new DynamicJsonObject();
        private CreateResult createUserStory;
        private CreateResult attachmentContentCreateResult;
        private CreateResult attachmentContainerCreateResult;
        private string userStoryReference;
        private string emailSubject;
        private string emailBody;
        private Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();
        private string[] allAttachments;
        private string base64String;
        private string attachmentFileName;
        private string attachmentFileNameForDictionary;
        private int duplicateFileCount = 0;
        private IMailFolder inboxFolder;
        private IList<UniqueId> emailMessageIds;
        private MimeMessage message;

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

        private static void LoginToGmail(ImapClient client)
        {
            client.ServerCertificateValidationCallback = (s, c, ch, e) => true;
            client.Connect(EmailConstant.GoogleHost, EmailConstant.ImapPort, SecureSocketOptions.SslOnConnect);
            client.AuthenticationMechanisms.Remove(EmailConstant.GoogleOAuth);
            client.Authenticate(EmailConstant.GoogleUsername, EmailConstant.GenericPassword);
        }

        private void SetUpMailbox(ImapClient client)
        {
            client.Inbox.Open(FolderAccess.ReadWrite);
            inboxFolder = client.GetFolder(EmailConstant.GmailInbox);
            emailMessageIds = client.Inbox.Search(SearchQuery.All);
        }

        private void CreateUserStoryWithEmail()
        {
            if (emailSubject.IsEmpty())
            {
                emailSubject = EmailConstant.NoSubject;
            }
            toCreate[RallyConstant.Name] = (emailSubject);
            toCreate[RallyConstant.Description] = (emailBody);
            createUserStory = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);
            Console.WriteLine("Created User Story: " + emailSubject);
        }

        private static string FileToBase64(string attachment)
        {
            Byte[] attachmentBytes = File.ReadAllBytes(attachment);
            string base64EncodedString = Convert.ToBase64String(attachmentBytes);
            return base64EncodedString;
        }

        private void DownloadAttachments(MimeMessage message)
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
                        attachmentFile = string.Format(fileNameWithoutExtension + "-{0}" + "{1}", ++duplicateFileCount, extension);
                        attachmentFilePath = Path.Combine(StorageConstant.MimeKitAttachmentsDirectory, attachmentFile);
                    }

                    using (var attachmentStream = File.Create(attachmentFilePath))
                    {
                        MimeKit.MimePart part = (MimeKit.MimePart)attachment;
                        part.ContentObject.DecodeTo(attachmentStream);
                    }

                    Console.WriteLine("Downloaded: " + attachmentFile);
                }
            }
        }

        private void ProcessAttachments()
        {
            allAttachments = Directory.GetFiles(StorageConstant.MimeKitAttachmentsDirectory);
            foreach (string file in allAttachments)
            {
                base64String = FileToBase64(file);
                attachmentFileName = Path.GetFileName(file);
                attachmentFileNameForDictionary = string.Empty;

                if (!(attachmentsDictionary.TryGetValue(base64String, out attachmentFileNameForDictionary)))
                {
                    attachmentsDictionary.Add(base64String, attachmentFileName);
                    Console.WriteLine("Accepted: " + file);
                }
                else
                {
                    Console.WriteLine("Omitting Duplicate: " + file);
                }

                //instead of deleting the files, maybe just invoke a method to erase all the content in the directory
                File.Delete(file);
            }
        }

        private void UploadAttachmentsToRallyUserStory()
        {
            foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
            {
                try
                {
                    attachmentContent[RallyConstant.Content] = attachmentPair.Key;
                    attachmentContentCreateResult = _rallyRestApi.Create(RallyConstant.AttachmentContent, attachmentContent);
                    userStoryReference = attachmentContentCreateResult.Reference;
                    attachmentContainer[RallyConstant.Artifact] = createUserStory.Reference;
                    attachmentContainer[RallyConstant.Content] = userStoryReference;
                    attachmentContainer[RallyConstant.Name] = attachmentPair.Value;
                    attachmentContainer[RallyConstant.Description] = RallyConstant.EmailAttachment;
                    attachmentContainer[RallyConstant.ContentType] = StorageConstant.FileType;
                    attachmentContainerCreateResult = _rallyRestApi.Create(RallyConstant.Attachment, attachmentContainer);
                }
                catch (WebException)
                {
                    throw new WebException();
                }
            }
        }

        public void SyncUsingMimeKit(string rallyWorkspace, string rallyScrumTeam)
        {
            LoginToRally();

            toCreate[RallyConstant.WorkSpace] = rallyWorkspace;
            toCreate[RallyConstant.Project] = rallyScrumTeam;

            using (ImapClient client = new ImapClient())
            {
                LoginToGmail(client);
                SetUpMailbox(client);

                foreach (UniqueId messageId in emailMessageIds)
                {
                    message = inboxFolder.GetMessage(messageId);
                    emailSubject = message.Subject;
                    emailBody = message.TextBody;

                    CreateUserStoryWithEmail();
                    DownloadAttachments(message);
                    ProcessAttachments();
                    UploadAttachmentsToRallyUserStory();
                    attachmentsDictionary.Clear();
                    duplicateFileCount = 0;
                } 
            }
        }
   
    }
}
