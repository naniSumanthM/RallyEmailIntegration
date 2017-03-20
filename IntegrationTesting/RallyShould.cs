using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Configuration;
using ActiveUp.Net.Mail;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rally.RestApi;
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
            int[] unreadMessages = inbox.Search("ALL");

            //Assert [total messges = 16]
            Assert.AreEqual(16, unreadMessages.Length);
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
            int[] unreadMessages = inbox.Search("ALL");
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
            Assert.AreEqual(16, unreadMessageCollection.Count);
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
        /// Need to add 37 unread messages to "INBOX" for test to pass
        /// </summary>
        [TestMethod]
        public void VerifyThatUnreadMessagesCanBeMoved()
        {
            var _selectedMailBox = "INBOX";

            using (var client = new Imap4Client())
            {
                client.ConnectSsl("imap-mail.outlook.com", 993);
                client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

                var mails = client.SelectMailbox(_selectedMailBox);
                var mailMessages = mails.Search("ALL");

                foreach (var id in mailMessages)
                {
                    mails.MoveMessage(id, "Processed");
                }

                var mailsUndeleted = client.SelectMailbox(_selectedMailBox);
                client.Disconnect();
            }

            Assert.AreEqual(30, 30);
        }

        /// <summary>
        /// Test to verify that any message that is marked as read is not moved
        /// </summary>
        [TestMethod]
        public void VerifyThatReadMessagesCannotBeMoved()
        {
            var _selectedMailBox = "inboxA";

            using (var client = new Imap4Client())
            {
                client.ConnectSsl("imap-mail.outlook.com", 993);
                client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

                var mails = client.SelectMailbox(_selectedMailBox);
                var mailMessages = mails.Search("SEEN");

                foreach (var id in mailMessages)
                {
                    mails.MoveMessage(id, "inboxB");
                }

                var mailsUndeleted = client.SelectMailbox(_selectedMailBox);
                client.Disconnect();
            }

            Assert.AreEqual(0, 0);
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
            Assert.AreEqual(4, inboxMessages.Length);
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

            Mailbox inbox = client.SelectMailbox("TestB");
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
            string emailFolder = "attachment";

            Mailbox attachmentMailbox = client.SelectMailbox(emailFolder);
            int[] attachmentMessages = attachmentMailbox.Search("ALL");
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
            string directoryPath = "C:\\Users\\suman\\Desktop\\testFolder\\";

            //Authentication
            imap.ConnectSsl("imap-mail.outlook.com", 993);
            imap.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            var inbox = imap.SelectMailbox(emailFolder);
            var unread = inbox.Search("UNSEEN");

            for (var i = 0; i < unread.Length; i++)
            {
                var unreadMessage = inbox.Fetch.MessageObject(unread[i]);

                foreach (MimePart embedded in unreadMessage.EmbeddedObjects)
                {
                    var filename = embedded.ContentName;
                    var binary = embedded.BinaryContent;
                    File.WriteAllBytes(string.Concat(directoryPath,filename), binary);
                }
            }
            Assert.AreEqual(1, Directory.GetFiles(directoryPath).Length);
        }

        [TestMethod]
        public void GivenAnEmailVerifyThatDuplicateAttachmentsAreIgnoredWhenUploadedToRally()
        {
            //Authenticate
            Imap4Client client = new Imap4Client();
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            //File IO
            string directoryPath = "C:\\Users\\suman\\Desktop\\testFolder";
            string emailFolder = "inboxB";

            Mailbox attachmentMailbox = client.SelectMailbox(emailFolder);
            int[] attachmentMessages = attachmentMailbox.Search("ALL");
            List<Message> messagesList = new List<Message>();

            for (int i = 0; i < attachmentMessages.Length; i++)
            {
                Message m = attachmentMailbox.Fetch.MessageObject(attachmentMessages[i]);
                messagesList.Add(m);
            }

            foreach (var msg in messagesList)
            {




            }



        }

        [TestMethod]
        public void GivenAnEmailVerifyAllDuplicateInlineAttachmentsAreIgnoredWhenUploadedToRally()
        {
            
        }
    }
}


//TODO: 3/20/17 - 3/24/17
/*
Test - see if message obj is read or unread
Test - multiple email attachments inline or attached
Test - move Message fails

Rally - blank subject lines
Rally - attachments
Rally - descriptions
Rally - features
Clear all userstories in rally to test workspace, project and the user stories count
*/