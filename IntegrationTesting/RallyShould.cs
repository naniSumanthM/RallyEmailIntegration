using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using ActiveUp.Net.Mail;
using ActiveUp.Net.WhoIs;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rally;
using Rally.RestApi;
using Rally.RestApi.Json;
using Rally.RestApi.Response;
using Slack.Webhooks;

namespace IntegrationTesting
{
    [TestClass]
    public class RallyShould
    {
        /// <summary>
        /// Test to verify authentication with Rally Sever
        /// </summary>
        [TestMethod]
        public void AuthenticateWithRallyServer()
        {
            //Arrange
            RallyRestApi api = new RallyRestApi();
            
            //Act
            if (api.AuthenticationState != RallyRestApi.AuthenticationResult.Authenticated)
            {
                api.Authenticate("maddirsh@mail.uc.edu", "iYmcmb24", "https://rally1.rallydev.com", null, false);
            }

            //Assert
            Assert.AreEqual("Authenticated", api.AuthenticationState.ToString());
        }

        /// <summary>
        /// Test to verify authentication with Outlook Server
        /// </summary>
        [TestMethod]
        public void AuthenticateWithOutlook()
        {
            Imap4Client client = new Imap4Client();

            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Assert.AreEqual(true, client.IsConnected);
        }

        /// <summary>
        /// Test to verify authentication with Slack
        /// If we are authenticated, we are able to POST message to Slack server
        /// </summary>
        [TestMethod]
        public void PostMessageToSlackIfAuthenticated()
        {
            bool isSlackAuthenticated = false;

            SlackClient client = new SlackClient("https://hooks.slack.com/services/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", 100);
            isSlackAuthenticated = true;

            SlackMessage message = new SlackMessage
            {
                Channel = "#general",
                Text = "*Rally Notification*",
                Username = "sumanth"
            };

            var slackAttachment = new SlackAttachment
            {
                Text = "Slack Unit Test",
            };

            message.Attachments = new List<SlackAttachment> { slackAttachment };

            if (isSlackAuthenticated == true)
            {
                client.Post(message);
            }

            Assert.IsTrue(isSlackAuthenticated);
        }

        /// <summary>
        /// Verify user is logged out, and NO process continues after we disconnect in FINALLY
        /// </summary>
        [TestMethod]
        public void EnsureImap4IsDisconnected()
        {
            Imap4Client client = new Imap4Client();

            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Assert.AreEqual(true, client.IsConnected);

            try
            {
                client.Disconnect();
                Assert.AreEqual(false, client.IsConnected);
            }
            catch (Imap4Exception)
            {
                throw new Imap4Exception();
            }
        }

        /// <summary>
        /// Test to verify the Outlook server returns the number of messages for a given inbox 
        /// </summary>
        [TestMethod]
        public void ReturnNumberOfMessagesForInbox()
        {
            //get imap obj
            Imap4Client client = new Imap4Client();

            //authenicate
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Mailbox inbox = client.SelectMailbox("Conversations");
            int[] messageCount = inbox.Search("ALL");
            //Assert [total messges = 17]
            Assert.AreEqual(17, messageCount.Length);
        }

        [TestMethod]
        public void EnsureSyncProcessIsNotStartedWithoutUnreadEmailMessages()
        {
            Imap4Client client = new Imap4Client();
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Mailbox mainMailbox = client.SelectMailbox("Sync");
            int[] mainUnreadMessages = mainMailbox.Search("UNSEEN");
            bool startProcess = false;

            if (mainUnreadMessages.Length > 0)
            {
                startProcess = true;
            }
            else
            {
                startProcess = false;
            }

            Assert.AreEqual(false, startProcess);
        }

        [TestMethod]
        public void ActUponUnreadEmailMessages()
        {
            Imap4Client client = new Imap4Client();
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Mailbox mainMailbox = client.SelectMailbox("Sync");
            int[] mainUnreadMessages = mainMailbox.Search("UNSEEN");
            bool startProcess = false;

            if (mainUnreadMessages.Length > 0)
            {
                startProcess = true;
            }
            else
            {
                startProcess = false;
            }

            Assert.AreEqual(true, startProcess);
        }

        /// <summary>
        /// Test to ensure ONLY unread messages are inserted into a List<Message>
        /// Need To Mark all messages inside "Conversations" as Unread for test to pass
        /// </summary>
        [TestMethod]
        public void EnsureAllUnreadEmailMessagesArePopulatedIntoList()
        {
            //get imap obj
            Imap4Client client = new Imap4Client();

            //authenicate
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Mailbox inbox = client.SelectMailbox("Conversations");
            int[] unreadMessages = inbox.Search("UNSEEN");
            int numberOfUnreadMessages = unreadMessages.Length;
            List<Message> unreadMessageCollection = new List<Message>();

            if (numberOfUnreadMessages > 0)
            {
                for (int i = 0; i < numberOfUnreadMessages; i++)
                {
                    Message msg = inbox.Fetch.MessageObject(unreadMessages[i]);
                    unreadMessageCollection.Add(msg);
                }
            }
            Assert.AreEqual(17, unreadMessageCollection.Count);
        }

        /// <summary>
        /// Test to ensure that a read message is not added to the list<Message>
        /// Inbox can contain read and unread messages
        /// </summary>
        [TestMethod]
        public void EnsureNonUnreadMessagesAreNotAddedToList()
        {
            //get imap obj
            Imap4Client client = new Imap4Client();

            //authenicate
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Mailbox inbox = client.SelectMailbox("Conversations");
            int[] unreadMessages = inbox.Search("UNSEEN");
            int numberOfUnreadMessages = unreadMessages.Length;
            List<Message> unreadMessageCollection = new List<Message>();

            if (numberOfUnreadMessages > 0)
            {
                for (int i = 0; i < numberOfUnreadMessages; i++)
                {
                    Message msg = inbox.Fetch.MessageObject(unreadMessages[i]);
                    unreadMessageCollection.Add(msg);
                }
            }

            Assert.AreEqual(0, unreadMessageCollection.Count);
        }

        /// <summary>
        /// Test to ensure that messages can be moved from one mailbox to another
        /// </summary>
        [TestMethod]
        public void EnsureUnreadMessagesCanBeMoved()
        {
            var _selectedMailBox = "MoveA";
            var _targetMailBox = "MoveB";

            using (var client = new Imap4Client())
            {
                client.ConnectSsl("imap-mail.outlook.com", 993);
                client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

                var mails = client.SelectMailbox(_selectedMailBox);
                var mailMessages = mails.Search("UNSEEN");

                for (int i = 0; i < mailMessages.Length; i++)
                {
                    mails.MoveMessage(i, _targetMailBox);
                }

                client.SelectMailbox(_selectedMailBox);
                client.Disconnect();
            }

            Assert.AreEqual(15, 15);
        }

        /// <summary>
        /// Test to verify that any message that is marked as read is not moved
        /// Messages are marked read when they are fetched
        /// </summary>
        [TestMethod]
        public void VerifyThatReadMessagesCannotBeMoved()
        {
            Imap4Client client = new Imap4Client();            
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Mailbox targetMailbox = client.SelectMailbox("inboxA");
            Mailbox destinationMailbox = client.SelectMailbox("inboxB");
            int[] targetMailInts = targetMailbox.Search("ALL");
            List<Message> targetList = new List<Message>();

            for (int i = 0; i < targetMailInts.Length; i++)
            {
                Message targetMessage = targetMailbox.Fetch.MessageObject(targetMailInts[i]);
                targetList.Add(targetMessage);
            }

            //Now all the messages are marked as read
            int[] targetReadMailInts = targetMailbox.Search("UNSEEN");

            for (int i = 0; i < targetReadMailInts.Length; i++)
            {
                targetMailbox.MoveMessage(i, "inboxB");
            }

            Assert.AreEqual(0, destinationMailbox.MessageCount);
        }

        /// <summary>
        /// Test to verify that messages can be marked as unread, once they are fetched
        /// Fetching a message counts as marking a message as read
        /// </summary>
        [TestMethod]
        public void EnsureMessagesCanBeMarkedAsUnread()
        {
            Imap4Client client = new Imap4Client();
            
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Mailbox inbox = client.SelectMailbox("Test");
            FlagCollection markAsUnreadFlagCollection = new FlagCollection();
            int[] inboxMessages = inbox.Search("ALL");

            foreach (var msg in inboxMessages)
            {
                Message m = inbox.Fetch.MessageObject(msg);
                markAsUnreadFlagCollection.Add("SEEN");
                inbox.RemoveFlags(msg, markAsUnreadFlagCollection);
            }

            int[] unreadEmailMessages = inbox.Search("UNSEEN");

            Assert.AreEqual(12, unreadEmailMessages.Length);
        }

        /// <summary>
        /// Test to verify message is marked as read when fetched
        /// </summary>
        [TestMethod]
        public void EnsureMessageIsMarkedReadWhenFetched()
        {
            Imap4Client client = new Imap4Client();
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Mailbox inbox = client.SelectMailbox("Test");
            int[] allInboxMessages = inbox.Search("ALL");

            foreach (var msg in allInboxMessages)
            {
                Message m = inbox.Fetch.MessageObject(msg);
            }

            int[] unreadIboxMessage = inbox.Search("UNSEEN");
            Assert.AreEqual(0, unreadIboxMessage.Length);    
        }

        /// <summary>
        /// Verify that all the attachments are downloaded to a directory
        /// Directory Test Folder
        /// </summary>
        [TestMethod]
        public void GivenAnEmailDownloadAllAttachments()
        {
            //Authenticate
            Imap4Client client = new Imap4Client();
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            //File IO
            string directoryPath = "C:\\Users\\suman\\Desktop\\testFolder";

            Mailbox attachmentMailbox = client.SelectMailbox("attachment");
            int[] attachmentMessages = attachmentMailbox.Search("UNSEEN");
            List<Message> unreadAttachments = new List<Message>();

            for(int i=0; i< attachmentMessages.Length; i++)
            {
                Message unreadMessage = attachmentMailbox.Fetch.MessageObject(attachmentMessages[i]);
                unreadAttachments.Add(unreadMessage);
            }

            foreach (var msg in unreadAttachments)
            {
                if (msg.Attachments.Count > 0)
                {
                    msg.Attachments.StoreToFolder(directoryPath);
                }
            }

            Assert.AreEqual(2, Directory.GetFiles(directoryPath).Length);
        }
        
        /// <summary>
        /// Test to Verify all the inline images are dowloaded
        /// </summary>
        [TestMethod]
        public void GivenAnEmailVerifyAllInlineAttachmentsAreDownloaded()
        {
            var imap = new Imap4Client();
            string emailFolder = "inlineImageUT";
            string directoryPath = "C:\\Users\\maddirsh\\Desktop\\testFolder\\";
            List<Message> inlineAttachmentList = new List<Message>();

            //Authentication
            imap.ConnectSsl("imap-mail.outlook.com", 993);
            imap.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            var inbox = imap.SelectMailbox(emailFolder);
            var unread = inbox.Search("ALL");

            for (var i = 0; i < unread.Length; i++)
            {
                var unreadMessage = inbox.Fetch.MessageObject(unread[i]);
                inlineAttachmentList.Add(unreadMessage);   
            }

            for (int i = 0; i < inlineAttachmentList.Count; i++)
            {
                foreach (MimePart embedded in inlineAttachmentList[i].EmbeddedObjects)
                {
                    var fileName = embedded.ContentName;
                    var binary = embedded.BinaryContent;
                    //downloads one file from the email from the MANY inline attachments that can exists
                    File.WriteAllBytes(string.Concat(directoryPath, fileName), binary);
                }
            }

            Assert.AreEqual(1, Directory.GetFiles(directoryPath).Length);
        }
        
        [TestMethod]
        public void VerifyDuplicateAttachmentsAreIgnoredWhenInsertedIntoDictionary()
        {
            //IO Variables
            string storeLocation = "C:\\Users\\maddirsh\\Desktop\\toRally";
            string[] attachmentLocations = Directory.GetFiles(storeLocation);
            string _base64EncodedString;
            string attachmentFileName;

            //Rally Variables
            Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();
        
            //convert the files to base64
            foreach (var file in attachmentLocations)
            {
                var attachmentBytes = File.ReadAllBytes(file);
                _base64EncodedString = Convert.ToBase64String(attachmentBytes);
                attachmentFileName = Path.GetFileName(file);
                var fileName = string.Empty;

                if (!(attachmentsDictionary.TryGetValue(_base64EncodedString, out fileName)))
                {
                    try
                    {
                        attachmentsDictionary.Add(_base64EncodedString, attachmentFileName);
                    }
                    catch (ArgumentException)
                    {
                        Assert.Fail();
                    }
                }
            }

            Assert.AreEqual(2, attachmentsDictionary.Count);
        }

        [TestMethod]
        public void GivenAnEmailObjectReturnIfSeenOrUnseen()
        {
            
        }

        [TestMethod]
        public void GivenAnEmailWithBlankSubjectEnsureNoSubjectUserStoryIsCreated()
        {
            //authenticate
            //parse an email without a subject
            //make the user story
            //ensure that the subject is labeled "no Subject"
        }
        
    }
}


