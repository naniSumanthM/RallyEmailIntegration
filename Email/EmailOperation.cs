using System;
using ActiveUp.Net.Mail;
using System.Collections.Generic;

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
                imap.ConnectSsl(Credential.outlookImapHost, Credential.outlookImapPort);
                imap.Login(Credential.outlookUserName, Credential.outlookPassword);

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
                imap.ConnectSsl(Credential.outlookImapHost, Credential.outlookImapPort);
                imap.Login(Credential.outlookUserName, Credential.outlookPassword);

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
                client.ConnectSsl(Credential.outlookImapHost, Credential.outlookImapPort);
                client.Login(Credential.outlookUserName, Credential.outlookPassword);

                //Stage the enviornment
                Mailbox inbox = client.SelectMailbox("INBOX");
                int[] unread = inbox.Search("UNSEEN");
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
                client.ConnectSsl(Credential.outlookImapHost, 993);
                client.Login(Credential.outlookUserName, Credential.outlookPassword);

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
                client.ConnectSsl(Credential.outlookImapHost, Credential.outlookImapPort);
                client.Login(Credential.outlookUserName, Credential.outlookPassword);

                //client.CreateMailbox("Created");
                Mailbox inbox = client.SelectMailbox(Credential.inboxFolder);
                Console.WriteLine(inbox.MessageCount);

                //Array of ALL email objects in selected mailbox
                int[] ids = inbox.Search("ALL");

                //iterate and move each message to a different folder
                foreach (var id in ids)
                {
                    inbox.MoveMessage(id, Credential.processedFolder);
                }

                var mailsUndeleted = client.SelectMailbox(Credential.inboxFolder);
                Console.WriteLine("Moved Messages to: " + Credential.processedFolder);
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
                imap.ConnectSsl(Credential.outlookImapHost, Credential.outlookImapPort);
                imap.Login(Credential.outlookUserName, Credential.outlookPassword);

                //setup Enviornment
                Mailbox inbox = imap.SelectMailbox(Credential.inboxFolder);
                int[] unread = inbox.Search("UNSEEN");
                Console.WriteLine("Unread Messages: " + unread.Length);

                //Crawl through the inbox and parse unread subject lines, then move those email objects to a folder
                if (unread.Length > 0)
                {
                    //Add the unread emails to a collection
                    for (int i = 0; i < unread.Length; i++)
                    {
                        Message msg = inbox.Fetch.MessageObject(unread[i]);
                        //explicitly mark as read for each email obejct
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
                        inbox.MoveMessage(item, Credential.processedFolder);
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

        #region markAsRead()
        /// <summary>
        /// Method only fetches the Headerobject from unread emails and marks the email object as READ EXPLCITLY
        /// </summary>

        public void markAsRead()
        {
            Imap4Client client = new Imap4Client();
            List<Header> unreadHeaderList = new List<Header>();

            try
            {
                //Authenticate
                client.ConnectSsl(Credential.outlookImapHost, Credential.outlookImapPort);
                client.Login(Credential.outlookUserName, Credential.outlookPassword);

                //Stage the enviornment
                Mailbox inbox = client.SelectMailbox(Credential.inboxFolder);
                int[] unread = inbox.Search(Credential.statusUnseen);
                Console.WriteLine("Unread Messages"+unread.Length);

                if (unread.Length>0)
                {
                    //iterate and add the unread header object into the list
                    for (int i = 0; i < unread.Length; i++)
                    {
                        Header unreadHeader = (inbox.Fetch.HeaderObject(unread[i]));
                        Console.WriteLine(unreadHeader.ConfirmRead); //Need to mark a message as read
                        unreadHeaderList.Add(unreadHeader);
                    }

                    foreach (var header in unreadHeaderList)
                    {
                        Console.WriteLine(header.Subject);
                    }
                }
                else
                {
                    Console.WriteLine("No unread E-mail");
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
                unreadHeaderList.Clear();
            }
        }
        #endregion
    }
}


/*                                                      Design Queries
                                            Fetching HeaderObject vs MessageObject

Header object when fecthed is quicker, but needs to be marked as read explicitly and then moved to the processed folder
Message object will be needed in some point, |Even though the library marks it as read| we need to state explicitly to mark it as read| 

 */

//iEnumarable
//Delegates
//security



//Mailbox inbox = imap.SelectMailbox("inbox");
//if (inbox.MessageCount > 0)
//{
//Header header = inbox.Fetch.HeaderObject(1);

//this.AddLogEntry(string.Format("Subject: {0} From :{1} ", header.Subject, header.From.Email));
//}