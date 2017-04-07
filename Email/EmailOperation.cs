﻿using System;
using ActiveUp.Net.Mail;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Remoting.Messaging;
using System.Threading;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using Header = ActiveUp.Net.Mail.Header;
using MimePart = ActiveUp.Net.Mail.MimePart;

namespace Email
{
    class EmailOperation
    {
        #region getAllEmail()

        /// <summary>
        /// Returns the subject along with the body of an email.
        /// Does not take into account read and unread email messages.
        /// </summary>
        public void getAllEmail()
        {
            Imap4Client imap = new Imap4Client();

            try
            {
                //Authenticate with the Outlook Server
                imap.ConnectSsl(Constant.OutlookImapHost, Constant.ImapPort);
                imap.Login(Constant.OutlookUserName, Constant.GenericPassword);

                //Select a mailbox folder
                Mailbox inbox = imap.SelectMailbox("inbox");
                Console.WriteLine("Message Count: " + inbox.MessageCount);

                //Iterate through the mailbox and fetch the mail objects
                for (int i = 1; i <= inbox.MessageCount; i++)
                {
                    Header header = (inbox.Fetch.HeaderObject(i));
                    Message msg = (inbox.Fetch.MessageObject(i));
                    //need not fetch the header if we r fetching the whole object
                    Console.WriteLine(header.Subject);
                    Console.WriteLine(msg.BodyText.Text);
                }
            }
            catch (Imap4Exception)
            {
                throw new Imap4Exception();
            }
            catch (Exception)
            {
                throw new Exception();
            }
            finally
            {
                imap.Disconnect();
            }
        }

        #endregion

        #region fetchUnreadEmails()

        /// <summary>
        /// Fetches all the unread emails and prints their subject along with their body
        /// </summary>
        public void FetchUnreadEmails()
        {
            Imap4Client client = new Imap4Client();
            List<Message> unreadList = new List<Message>();

            try
            {
                //Authenticate 
                client.ConnectSsl(Constant.OutlookImapHost, Constant.ImapPort);
                client.Login(Constant.OutlookUserName, Constant.GenericPassword);

                //Stage the enviornment
                Mailbox inbox = client.SelectMailbox("INBOX");
                int[] unread = inbox.Search("UNSEEN");
                //returns an int of the number of unread email objects, given an inbox
                Console.WriteLine("Unread Messages: " + unread.Length);

                if (unread.Length > 0)
                {
                    //iterate and store each unread email object into a list
                    for (int i = 0; i < unread.Length; i++)
                    {
                        Message newMessage = inbox.Fetch.MessageObject(unread[i]);
                        unreadList.Add(newMessage);
                    }
                    foreach (Message item in unreadList)
                    {
                        Console.WriteLine(item.Subject);
                        Console.WriteLine(item.BodyText.Text);
                    }
                }
                else
                {
                    Console.WriteLine("No Unread Messages found");
                }
            }
            catch (Imap4Exception)
            {
                throw new Imap4Exception();
            }
            catch (Exception)
            {
                throw new Exception();
            }
            finally
            {
                client.Disconnect();
                unreadList.Clear();
            }
        }

        #endregion

        #region fetchUnreadSubjectLines()

        /// <summary>
        /// Method that (FETCHES unread EMAIL OBJECT) and stores it into a list. 
        /// The list is iterated to parse only the subject line, but it can be used to get the body as well.
        /// iSSUE: fetch entire email object and mark as read or fetch only the Header obj and mark as read manually??
        /// </summary>
        public void fetchUnreadSubjectLines()
        {
            //Authenticate with the Outlook Server
            Imap4Client imap = new Imap4Client();
            List<Message> unreadList = new List<Message>();

            try
            {
                imap.ConnectSsl(Constant.OutlookImapHost, Constant.ImapPort);
                imap.Login(Constant.OutlookUserName, Constant.GenericPassword);

                //setup Enviornment
                Mailbox inbox = imap.SelectMailbox("INBOX");
                int[] unread = inbox.Search("UNSEEN");
                Console.WriteLine("Unread Messages: " + unread.Length);

                if (unread.Length > 0)
                {
                    //Loop to load the unread subject lines into a list<Message>
                    //Messages will be marked as read
                    for (int i = 0; i < unread.Length; i++) //difference between index at 1 and <= unread.Length
                    {
                        Message msg = inbox.Fetch.MessageObject(unread[i]);
                        unreadList.Add(msg);
                    }
                    //iterate through the list and print the subject lines
                    foreach (var item in unreadList)
                    {
                        Console.WriteLine(item.Subject);
                    }
                }
                else
                {
                    Console.WriteLine("No unread messages found");
                }
            }
            catch (Imap4Exception)
            {
                throw new Imap4Exception();
            }
            catch (Exception)
            {
                throw new Exception();
            }
            finally
            {
                imap.Disconnect();
                unreadList.Clear();
            }
        }

        #endregion

        #region CreateMailbox()

        /// <summary>
        /// Create a new folder within the outlook email server
        /// </summary>
        public void createMailBox()
        {
            Imap4Client client = new Imap4Client();
            try
            {
                //Connect and Authenticate
                client.ConnectSsl(Constant.OutlookImapHost, 993);
                client.Login(Constant.OutlookUserName, Constant.GenericPassword);

                //create mailbox
                client.CreateMailbox("Mailbox-A");
                Console.WriteLine("Created Mailbox");
            }
            catch (Imap4Exception)
            {
                throw new Imap4Exception();
            }
            catch (Exception)
            {
                throw new Exception();
            }
            finally
            {
                client.Disconnect();
            }
        }

        #endregion

        #region moveMessages()

        /// <summary>
        /// Move email objects from one folder to another. 
        /// Unread email objects will be left unread, UNLESS fetched
        /// </summary>
        public void moveMessages()
        {
            Imap4Client client = new Imap4Client();

            try
            {
                //Authenticate 
                client.ConnectSsl("imap.gmail.com", 993);
                client.Login("rallyintegration@gmail.com", "iYmcmb24");

                Mailbox inbox = client.SelectMailbox("Inbox");
                Console.WriteLine(inbox.MessageCount);
                int[] messageCount = inbox.Search("ALL");

                for (int i = 0; i < messageCount.Length; i++)
                {
                    inbox.MoveMessage(i, "Starred");
                }

                Console.WriteLine("Moved Messages to: " + Constant.ProcessedFolder);
            }
            catch (Imap4Exception i)
            {
                Console.WriteLine("Imap4Exception Response" + i.Response + Environment.NewLine + "Imap4 Message" +
                                  i.Message);
                Console.WriteLine("Imap 4 target " + i.TargetSite);
                Console.WriteLine("Imap 4 source" + i.Source);
            }
            catch (WebException w)
            {
                Console.WriteLine("Web exception response " + w.Response + Environment.NewLine + "Web exception message" +
                                  w.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        #endregion

        #region moveUnreadEmail()

        /// <summary>
        /// Method will move all the unread emails to a folder called "Processed"
        /// </summary>
        public void moveUnreadEmail()
        {
            Imap4Client imap = new Imap4Client();

            try
            {
                //Authenticate
                imap.ConnectSsl(Constant.OutlookImapHost, 993);
                imap.Login("", Constant.GenericPassword);

                Mailbox inbox = imap.SelectMailbox(Constant.InboxFolder);
                int[] unread = inbox.Search(Constant.UnseenMessages);
                Console.WriteLine("Unread Messages: " + unread.Length);

                if (unread.Length > 0)
                {
                    for (int i = 0; i < unread.Length; i++)
                    {
                        inbox.MoveMessage(i, "Tickets");
                    }
                }
                else
                {
                    Console.WriteLine("No Unread Email");
                }
            }
            catch (Imap4Exception imap4)
            {
                Console.WriteLine(imap4.Message);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                imap.Disconnect();
            }
        }

        #endregion

        #region markMessageObjAsUnread()

        /// <summary>
        /// Method marks all READ mail as UNREAD MAIL.
        /// Method fetched the messageObject which automatically marks a message as read
        /// </summary>
        public void markMsgObjAsUnread()
        {
            Imap4Client client = new Imap4Client();
            FlagCollection markAsUnreadFlag = new FlagCollection();

            try
            {
                //Authenticate
                client.ConnectSsl(Constant.OutlookImapHost, Constant.ImapPort);
                client.Login(Constant.OutlookUserName, Constant.GenericPassword);

                //stage the enviornment
                Mailbox inbox = client.SelectMailbox("inbox");
                int[] allMessages = inbox.Search("ALL");

                Console.WriteLine("Message-Count: " + allMessages.Length);

                //Itearate and mark each mail object as unread
                foreach (var id in allMessages)
                {
                    Message msg = inbox.Fetch.MessageObject(id);
                    markAsUnreadFlag.Add("SEEN"); //adding all the read email objects to the flag collection
                    inbox.RemoveFlags(id, markAsUnreadFlag);
                    //then removing the flags, making each mail object as unread
                }
            }
            catch (Imap4Exception ie)
            {
                Console.WriteLine(string.Format("Imap4 Exception: {0}", ie.Message));
            }
            catch (Exception e)
            {
                Console.WriteLine(string.Format("Unexpected Exception: {0}"), e.Message);
            }
            finally
            {
                client.Disconnect();
            }
        }

        #endregion

        #region dowloadAttachments()

        /// <summary>
        /// Download all the attchments from an undread email message and store them into a folder
        /// </summary>
        public void dowloadAttachments()
        {
            Imap4Client imap = new Imap4Client();
            List<Message> unreadAttachments = new List<Message>();

            try
            {
                //Authenticate
                imap.ConnectSsl(Constant.GoogleImapHost, Constant.ImapPort);
                imap.Login(Constant.GoogleUserName, Constant.GenericPassword);

                Mailbox inbox = imap.SelectMailbox("Inbox");
                int[] unread = inbox.Search("UNSEEN");
                Console.WriteLine("Unread Messgaes: " + unread.Length);

                Console.WriteLine("Start");
                if (unread.Length > 0)
                {
                    //fetch each unread message
                    for (int i = 0; i < unread.Length; i++)
                    {
                        Message unreadMessage = inbox.Fetch.MessageObject(unread[i]);
                        unreadAttachments.Add(unreadMessage);
                    }

                    foreach (var attachemntMsg in unreadAttachments)
                    {
                        if (attachemntMsg.Attachments.Count > 0)
                        {
                            attachemntMsg.Attachments.StoreToFolder(Constant.RegularAttachmentsDirectory);
                        }
                        else
                        {
                            Console.WriteLine("No attachments for: " + attachemntMsg.Subject);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No Unread Messages");
                }
            }
            catch (IOException i)
            {
                Console.WriteLine(i.Message);
            }
            catch (Imap4Exception ie)
            {
                Console.WriteLine(ie.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                imap.Disconnect();
                unreadAttachments.Clear();
            }
            Console.WriteLine("End");
        }

        #endregion

        #region embeddedImages()

        /// <summary>
        /// Method to pull images that could have been copied & pasted, instead of attaching
        /// </summary>
        public void downlodInlineAttachments()
        {
            var imap = new Imap4Client();

            //Authenticate
            imap.ConnectSsl(Constant.GoogleImapHost, Constant.ImapPort);
            imap.Login(Constant.GoogleUserName, Constant.GenericPassword);

            var inbox = imap.SelectMailbox("Inbox");
            var unread = inbox.Search("unseen");

            Console.WriteLine("Unread Messgaes: " + unread.Length);

            if (unread.Length > 0)
                for (int i = 0; i < unread.Length; i++)
                {
                    var unreadMessage = inbox.Fetch.MessageObject(unread[i]);

                    foreach (MimePart embedded in unreadMessage.EmbeddedObjects)
                    {
                        var filename = embedded.ContentName;
                        var binary = embedded.BinaryContent;
                        File.WriteAllBytes(String.Concat(Constant.InlineAttachmentsDirectory, filename), binary);
                        Console.WriteLine("Downloaded: " + filename);
                    }
                }
            else
            {
                Console.WriteLine("Unread Messages Not Found");
            }
        }

        #endregion

        #region MoveMessagesUsingMimeKit()

        /// <summary>
        /// Authenticate with Mime Kit
        /// </summary>
        public void MimeKitTest()
        {
            using (var client = new ImapClient())
            {
                client.Connect("imap-mail.outlook.com", 993, SecureSocketOptions.SslOnConnect);
                client.Authenticate("rallyintegration@outlook.com", "iYmcmb24");
                Console.WriteLine(client.IsAuthenticated);

                if (client.IsConnected == true)
                {
                    FolderAccess inboxAccess = client.Inbox.Open(FolderAccess.ReadWrite);
                    IMailFolder destination = client.GetFolder("Tickets");
                    IList<UniqueId> uids = client.Inbox.Search(SearchQuery.All);

                    if (destination != null)
                    {
                        client.Inbox.MoveTo(uids, destination);
                    }
                    else
                    {
                        Console.WriteLine("Something is wrong");
                    }
                }
                else
                {
                    Console.WriteLine("ERROR>>>>>");
                }

                client.Disconnect(true);
            }
        }
        #endregion

        #region : get Subject & Body with MimeKit
        /// <summary>
        /// Looks like Microsoft keeps blocking me from the email server
        /// </summary>
        public void getMessages()
        {
            using (var client = new ImapClient())
            {
                client.Connect("imap.gmail.com", 993, SecureSocketOptions.SslOnConnect);
                client.Authenticate("rallyintegration@gmail.com", "iYmcmb24");
                Console.WriteLine(client.IsAuthenticated);

                if (client.IsConnected == true)
                {
                    FolderAccess inboxAccess = client.Inbox.Open(FolderAccess.ReadWrite);
                    IMailFolder destination = client.GetFolder("Inbox");
                    IList<UniqueId> uids = client.Inbox.Search(SearchQuery.All);

                    if (destination != null)
                    {
                        foreach (var x in uids)
                        {
                            var message = destination.GetMessage(x);
                            string subject = message.Subject;
                            string body = message.TextBody;
                            Console.WriteLine(body);
                        }
                    }
                }
                else
                {
                    throw new NullReferenceException();
                }
            }
        }
        #endregion

        #region : get attachemnts using mimeKit

        public void getAtttachments()
        {
            using (var client = new ImapClient())
            {
                //authenticate
                client.Connect("imap.gmail.com", 993, SecureSocketOptions.SslOnConnect);
                client.Authenticate("rallyintegration@gmail.com", "iYmcmb24");
                Console.WriteLine(client.IsAuthenticated);

                if (client.IsConnected == true)
                {
                    FolderAccess inboxAccess = client.Inbox.Open(FolderAccess.ReadWrite);
                    IMailFolder destination = client.GetFolder("Inbox");
                    //gets the id,s of all the messages within a gievn inbox
                    IList<UniqueId> uids = client.Inbox.Search(SearchQuery.All);
                    string saveAttachmentPath = "C:\\Users\\mimeKitAttachments\\";


                    if (destination != null)
                    {
                        foreach (var x in uids)
                        {
                            //What does get message do, is this the best way to get the message?
                            //I have read about examples like fetch to strip specific parts of a message

                            var message = destination.GetMessage(x);

                            foreach (MimeEntity attachment in message.BodyParts)
                            {
                                //TODO: store the attachments with the original filename within the email to the local disk
                            }
                        }
                    }
                }
                else
                {
                    throw new NullReferenceException();
                }
            }
        }



        #endregion

    }
}
