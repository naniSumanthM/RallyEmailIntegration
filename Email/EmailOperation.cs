using System;
using ActiveUp.Net.Mail;
using System.Collections.Generic;
using System.IO;

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
                imap.ConnectSsl(Constant.OutlookImapHost, Constant.OutlookImapPort);
                imap.Login(Constant.OutlookUserName, Constant.OutlookPassword);

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
                client.ConnectSsl(Constant.OutlookImapHost, Constant.OutlookImapPort);
                client.Login(Constant.OutlookUserName, Constant.OutlookPassword);

                //Stage the enviornment
                Mailbox inbox = client.SelectMailbox("INBOX");
                int[] unread = inbox.Search("UNSEEN"); //returns an int of the number of unread email objects, given an inbox
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
                imap.ConnectSsl(Constant.OutlookImapHost, Constant.OutlookImapPort);
                imap.Login(Constant.OutlookUserName, Constant.OutlookPassword);

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
                client.Login(Constant.OutlookUserName, Constant.OutlookPassword);

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
            List<Message> unreadList = new List<Message>();

            try
            {
                //Authenticate 
                client.ConnectSsl(Constant.OutlookImapHost, Constant.OutlookImapPort);
                client.Login(Constant.OutlookUserName, Constant.OutlookPassword);

                //client.CreateMailbox("Created");
                Mailbox inbox = client.SelectMailbox(Constant.InboxFolder);
                Console.WriteLine(inbox.MessageCount);

                //Array of ALL email objects in selected mailbox
                int[] ids = inbox.Search("ALL");

                //iterate and move each message to a different folder
                foreach (var id in ids)
                {
                    inbox.MoveMessage(id, Constant.ProcessedFolder);
                }

                var mailsUndeleted = client.SelectMailbox(Constant.InboxFolder);
                Console.WriteLine("Moved Messages to: " + Constant.ProcessedFolder);
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

        #region moveUnreadEmail()
        /// <summary>
        /// Method will move all the unread emails to a folder called "Processed"
        /// </summary>

        public void moveUnreadEmail()
        {
            Imap4Client imap = new Imap4Client();
            List<Message> unreadList = new List<Message>();

            try
            {
                //Connect and Authenticate
                imap.ConnectSsl(Constant.OutlookImapHost, Constant.OutlookImapPort);
                imap.Login(Constant.OutlookUserName, Constant.OutlookPassword);

                //setup Enviornment
                Mailbox inbox = imap.SelectMailbox(Constant.InboxFolder);
                int[] unread = inbox.Search("UNSEEN");
                Console.WriteLine("Unread Messages: " + unread.Length);

                //Crawl through the inbox and parse unread subject lines, then move those email objects to a folder
                if (unread.Length > 0)
                {
                    //Add the unread emails to a collection
                    for (int i = 0; i < unread.Length; i++)
                    {
                        Message msg = inbox.Fetch.MessageObject(unread[i]);
                        //explicitly mark as read for each email obejct?? - but the action of fetching does that for us
                        unreadList.Add(msg);
                    }

                    //print out the unread subejct line
                    foreach (var item in unreadList)
                    {
                        Console.WriteLine(item.Subject);
                        //Console.WriteLine(item.BodyText.Text);
                    }

                    //Move messages to the processed folder
                    foreach (var item in unread)
                    {
                        inbox.MoveMessage(item, Constant.ProcessedFolder);
                    }
                    //line could cause an error
                    //Mailbox movedFrom = imap.SelectMailbox(Credential.inboxFolder);
                }
                else
                {
                    Console.WriteLine("No Unread Email");
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
                client.ConnectSsl(Constant.OutlookImapHost, Constant.OutlookImapPort);
                client.Login(Constant.OutlookUserName, Constant.OutlookPassword);

                //stage the enviornment
                Mailbox inbox = client.SelectMailbox("inbox");
                int[] allMessages = inbox.Search("ALL");

                Console.WriteLine("Message-Count: " + allMessages.Length);

                //Itearate and mark each mail object as unread
                foreach (var id in allMessages)
                {
                    Message msg = inbox.Fetch.MessageObject(id);
                    markAsUnreadFlag.Add("SEEN"); //adding all the read email objects to the flag collection
                    inbox.RemoveFlags(id, markAsUnreadFlag); //then removing the flags, making each mail object as unread
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
                imap.ConnectSsl(Constant.OutlookImapHost, Constant.OutlookImapPort);
                imap.Login(Constant.OutlookUserName, Constant.OutlookPassword);

                Mailbox inbox = imap.SelectMailbox("inbox");
                int[] unread = inbox.Search("unseen");
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

                    //download the attachments and store it in a folder
                    //cannot attempt to download an attachment for messages that do not contain any attachments.
                    //need a universal path or some server to store or not just store but copy and paste the attachments
                    foreach (var attachemntMsg in unreadAttachments)
                    {
                        if (attachemntMsg.Attachments.Count > 0)
                        {
                            attachemntMsg.Attachments.StoreToFolder("C:\\Users\\maddirsh\\Desktop\\IntegrationService\\Email\\attachments");
                            //move the mssageObj into processed folder
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

            //Authentication
            imap.ConnectSsl(Constant.OutlookImapHost, Constant.OutlookImapPort);
            imap.Login(Constant.OutlookUserName, Constant.OutlookPassword);

            var inbox = imap.SelectMailbox("inbox");
            var unread = inbox.Search("unseen");
            Console.WriteLine("Unread Messgaes: " + unread.Length);

            if (unread.Length > 0)
                for (var i = 0; i < unread.Length; i++)
                {
                    var unreadMessage = inbox.Fetch.MessageObject(unread[i]);
                    foreach (MimePart embedded in unreadMessage.EmbeddedObjects)
                    {
                        var filename = embedded.ContentName;
                        var binary = embedded.BinaryContent;
                        File.WriteAllBytes(Constant.AttachmentPath + filename, binary);
                        Console.WriteLine("Downloaded: " + filename);
                    }
                }
            else
                Console.WriteLine("Unread Messages Not Found");
        }
        #endregion

    }
}


/*                                                      Design Queries
                                            Fetching HeaderObject vs MessageObject

Header object when fecthed is quicker, but needs to be marked as read explicitly and then moved to the processed folder
Message object will be needed in some point, |Even though the library marks it as read| we need to state explicitly to mark it as read 
*/

