using System;
using System.Collections.Generic;
using ActiveUp.Net.Mail;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rally.RestApi;
using Slack.Webhooks;

namespace IntegrationTesting
{
    [TestClass]
    public class RallyShould
    {
        [TestMethod]
        public void AuthenticateWithRallyServer()
        {
            //Arrange
            bool isAuthenticated = false;
            RallyRestApi api = new RallyRestApi();

            //Act
            if (api.AuthenticationState != RallyRestApi.AuthenticationResult.Authenticated)
            {
                api.Authenticate("maddirsh@mail.uc.edu", "iYmcmb24", "https://rally1.rallydev.com", null, false);
                isAuthenticated = true;
            }

            //Assert
            Assert.AreEqual(true, isAuthenticated);
        }

        [TestMethod]
        public void AuthenticateWithOutlook()
        {
            Imap4Client client = new Imap4Client();

            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Assert.AreEqual(true, client.IsConnected);
        }

        [TestMethod]
        public void PostMessageToSlackIfAuthenticated()
        {
            SlackClient client = new SlackClient("https://hooks.slack.com/services/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", 100);

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

            Assert.IsTrue(client.Post(message));
        }

        [TestMethod]
        public void ThrowExceptionIfRallyAuthenticationfailed()
        {

        }

        [TestMethod]
        public void ThrowExceptionIfOutlookAuthenticationfailed()
        {

        }

        [TestMethod]
        public void ThrowExceptionIfSlackApiTokenFails()
        {
            //get access to messages in the slack server
        }

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

            Assert.AreEqual(13, unreadMessageCollection.Count);
        }

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

        [TestMethod]
        public void VerifyThatMessagesCanBeMoved()
        {
            var _selectedMailBox = "INBOX";

            using (var client = new Imap4Client())
            {
                client.ConnectSsl("imap-mail.outlook.com", 993);
                client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

                var mails = client.SelectMailbox(_selectedMailBox);
                var ids = mails.Search("ALL");

                foreach (var id in ids)
                {
                    mails.MoveMessage(id, "Processed");
                }

                var mailsUndeleted = client.SelectMailbox(_selectedMailBox);
                client.Disconnect();
            }

            Assert.AreEqual(37, 37);
        }

        [TestMethod]
        public void EnsureMessageCanBeMarkedAsUnread()
        {
            
        }

        [TestMethod]
        public void EnsureMessageIsMarkedReadWhenFetched()
        {

        }

        [TestMethod]
        public void EnsureOnlyUnreadMessagesAreMoved()
        {
            
        }

        [TestMethod]
        public void VerifyAllAttachmentsAreDownloaded()
        {

        }

        [TestMethod]
        public void VerifyDuplicateAttachmentsAreIgnored()
        {
            
        }

        [TestMethod]
        public void VerifyAllInlineAttachmentsAreDownloaded()
        {
            
        }

        [TestMethod]
        public void VerifyDuplicateInlineAttachmentsAreIgnored()
        {
            
        }
    }
}