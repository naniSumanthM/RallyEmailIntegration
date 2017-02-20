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
        public RallyRestApi _rallyApi;
        public Imap4Client _imap4CLient;

        public string RallyUserName { get; set; }
        public string RallyPassword { get; set; }
        public string OutlookUserName { get; set; }
        public string OutlookPassword { get; set; }

        public Sync(string rallyUserName, string rallyPassword, string outlookUserName, string outlookPassword)
        {
            _rallyApi = new RallyRestApi();
            _imap4CLient = new Imap4Client();
            this.OutlookUserName = outlookUserName;
            this.OutlookPassword = outlookPassword;
            this.RallyUserName = rallyUserName;
            this.RallyPassword = rallyPassword;
        }

        public void LoginToOutlook()
        {
            _imap4CLient.ConnectSsl(OutlookConstant.OutlookHost, OutlookConstant.OutlookPort);
            _imap4CLient.Login(this.OutlookUserName, this.OutlookPassword);
        }

        public void LoginToRally()
        {
            if (this._rallyApi.AuthenticationState != RallyRestApi.AuthenticationResult.Authenticated)
            {
                _rallyApi.Authenticate(this.RallyUserName, this.RallyPassword, RallyConstant.ServerId, null, RallyConstant.AllowSso);
            }
        }

        public static string FileToBase64(string attachment)
        {
            Byte[] attachmentBytes = File.ReadAllBytes(attachment);
            string base64EncodedString = Convert.ToBase64String(attachmentBytes);
            return base64EncodedString;
        }

        //Email variables
        List<Message> unreadMsgCollection = new List<Message>();
        Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();

        //Rally variables
        DynamicJsonObject toCreate = new DynamicJsonObject();
        DynamicJsonObject attachmentContent = new DynamicJsonObject();
        DynamicJsonObject attachmentContainer = new DynamicJsonObject();
        CreateResult createUserStory;
        CreateResult attachmentContentCreateResult;
        CreateResult attachmentContainerCreateResult;

        //base 64 conversion variables
        string base64String;
        string attachmentFileName;
        string[] attachmentPaths;
        string userStoryReference;

        public void SyncUserStories(String workspace, string project)
        {
            toCreate[RallyConstant.WorkSpace] = workspace;
            toCreate[RallyConstant.Project] = project;

            try
            {
                LoginToOutlook();
                LoginToRally();

                Mailbox inbox = _imap4CLient.SelectMailbox(OutlookConstant.OutlookInboxFolder);
                int[] unread = inbox.Search(OutlookConstant.OutlookUnseenMessages);
                FlagCollection markAsUnreadFlag = new FlagCollection();

                if (unread.Length > 0)
                {
                    Console.WriteLine("Syncing: " + unread.Length + " Unread Messages");
                    FetchUnreadMessages(unread, inbox);

                    //Iterate through the collection 1) Create the user story 2) Check for attachments 3) Convert attachments to base 64 4)Delete attachments once pushed 
                    for (int i = 0; i < unreadMsgCollection.Count; i++)
                    {
                        //stage the user story
                        if (unreadMsgCollection[i].Subject.Equals(""))
                        {
                            unreadMsgCollection[i].Subject = OutlookConstant.NoSubject;
                        }
                        toCreate[RallyConstant.Name] = (unreadMsgCollection[i].Subject);
                        toCreate[RallyConstant.Description] = (unreadMsgCollection[i].BodyText.Text);
                        createUserStory = _rallyApi.Create(RallyConstant.HierarchicalRequirement, toCreate);

                        //check to see if message has attachments & then store them
                        if (unreadMsgCollection[i].Attachments.Count > 0)
                        {
                            //Do all the attachments from the email object get stored here?? _No it has to complete one iteration
                            unreadMsgCollection[i].Attachments.StoreToFolder(SyncConstant.AttachmentsDirectory);
                        }

                        //reference the path where the attachments live for the [ith] message
                        attachmentPaths = Directory.GetFiles(SyncConstant.AttachmentsDirectory);

                        //Convert each attachment to base64, populate the map, and move the file
                        PopulateAttachmentsDictionary();
                        PushAttachments(attachmentsDictionary, attachmentContent, attachmentContainer, createUserStory);
                        attachmentsDictionary.Clear();
                    }

                    //Move mail to processed folder and mark each mail object as unread
                    MarkAsUnread(unread, markAsUnreadFlag, inbox);

                    Console.WriteLine("Created " + unread.Length + " User Stories");
                }
                else
                {
                    Console.WriteLine("No Unread Messages Found...");
                }
            }
            catch (Imap4Exception i)
            {
                Console.WriteLine(i.Message);
            }
            catch (IOException i)
            {
                Console.WriteLine(i.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                _imap4CLient.Disconnect();
            }

        }

        private void FetchUnreadMessages(int[] unread, Mailbox inbox)
        {
            for (int i = 0; i < unread.Length; i++)
            {
                Message msg = inbox.Fetch.MessageObject(unread[i]);
                unreadMsgCollection.Add(msg);
            }
        }

        private void PopulateAttachmentsDictionary()
        {
            foreach (var file in attachmentPaths)
            {
                //Converting attachments to base 64
                base64String = FileToBase64(file);
                attachmentFileName = Path.GetFileName(file);
                var fileName = string.Empty;

                //populate the dictionary - eliminate adding duplicate files
                if (!(attachmentsDictionary.TryGetValue(base64String, out fileName)))
                {
                    attachmentsDictionary.Add(base64String, attachmentFileName);
                }

                Console.WriteLine("Uploading: " + file);
                File.Delete(file);
            }
        }

        private void PushAttachments(Dictionary<string, string> attachmentsDictionary, DynamicJsonObject attachmentContent,
                                     DynamicJsonObject attachmentContainer, CreateResult createUserStory)
        {
            foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
            {
                try
                {
                    //create attachment content
                    attachmentContent[RallyConstant.Content] = attachmentPair.Key;
                    attachmentContentCreateResult = _rallyApi.Create(RallyConstant.AttachmentContent, attachmentContent);
                    userStoryReference = attachmentContentCreateResult.Reference;

                    //create attachment contianer
                    attachmentContainer[RallyConstant.Artifact] = createUserStory.Reference;
                    attachmentContainer[RallyConstant.Content] = userStoryReference;
                    attachmentContainer[RallyConstant.Name] = attachmentPair.Value;
                    attachmentContainer[RallyConstant.Description] = RallyConstant.EmailAttachment;
                    attachmentContainer[RallyConstant.ContentType] = "file/";

                    //Create & associate the attachment
                    attachmentContainerCreateResult = _rallyApi.Create(RallyConstant.Attachment, attachmentContainer);
                }
                catch (WebException e)
                {
                    Console.WriteLine("Attachment: " + e.Message);
                }
            }
        }

        private static void MarkAsUnread(int[] unread, FlagCollection markAsUnreadFlag, Mailbox inbox)
        {
            foreach (var item in unread)
            {
                markAsUnreadFlag.Add(OutlookConstant.OutlookSeenMessages);
                inbox.RemoveFlags(item, markAsUnreadFlag);
                inbox.MoveMessage(item, OutlookConstant.OutlookProcessedFolder);
            }
        }
    }
}
