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
                Mailbox inbox = imap.SelectMailbox("inbox");
                int[] unread = inbox.Search("UNSEEN");
                Console.WriteLine("Unread Messages: " + unread.Length);

                if (unread.Length > 0)
                {
                    //Loop to load the unread subject lines into a list<Message>
                    //Messages will be marked as read
                    for (int i = 1; i <= unread.Length; i++)
                    {
                        Message unreadMsg = (inbox.Fetch.MessageObject(i));
                        unreadList.Add(unreadMsg);
                    }
                    //foreach to print the subject lines
                    foreach (Message item in unreadList)
                    {
                        Console.WriteLine(item.Subject);
                    }
                }
                else
                {
                    Console.WriteLine("No unread mail");
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
                client.ConnectSsl(Credential.outlookImapHost, Credential.outlookImapPort);
                client.Login(Credential.outlookUserName, Credential.outlookPassword);

                //Stage the enviornment
                Mailbox inbox = client.SelectMailbox("INBOX");
                int[] unread = inbox.Search("UNSEEN");
                Console.WriteLine("Unread Messages: " + unread.Length);

                if (unread.Length > 0)
                {
                    //iterate and store each unread email object into a list
                    for (int n = 1; n <= unread.Length; n++)
                    {
                        Message newMessage = inbox.Fetch.MessageObject(n);
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
                    Console.WriteLine("Unread Messages: " + unread.Length);
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

        #region CreateMailBOX()
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
                Mailbox inbox = client.SelectMailbox("CREATED");
                Console.WriteLine(inbox.MessageCount);

                //Array of ALL email objects in selected mailbox
                int[] ids = inbox.Search("ALL");

                //iterate and move each message to a different folder
                foreach (var id in ids)
                {
                    inbox.MoveMessage(id, "INBOX");
                }

                var mailsUndeleted = client.SelectMailbox("CREATED");
                Console.WriteLine("Moved Messages");
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


        public void moveUnreadEmails()
        {
            Imap4Client client = new Imap4Client();
            List<Message> unreadList = new List<Message>();

            try
            {
                //Authenticate 
                client.ConnectSsl(Credential.outlookImapHost, Credential.outlookImapPort);
                client.Login(Credential.outlookUserName, Credential.outlookPassword);

                //Stage the mailbox
                Mailbox inbox = client.SelectMailbox("INBOX");
                int[] unread = inbox.Search("UNSEEN");
                Console.WriteLine("Unread Messages: "+ unread.Length);

                //If there are unread messages in the inbox
                if (unread.Length>0)
                {
                    //store unread message objects into a list
                    for (int i = 1; i <= unread.Length; i++)
                    {
                        Message unreadMessage = inbox.Fetch.MessageObject(i);
                        unreadList.Add(unreadMessage);
                    }

                    //PRINT subjects in the list we assume has unread messages
                    foreach (var msg in unreadList)
                    {
                        Console.WriteLine(msg.Subject);
                        Console.WriteLine(msg.BodyText.Text);
                    }

                    foreach (var item in unread)
                    {
                        inbox.MoveMessage(item, "PROCESSED");
                    }

                }
                else
                {
                    Console.WriteLine("Unread Messages: " + unread.Length);
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
    }
}


//11/15/16 toDO
//Mark as read manually
//copy and delete AND FLAG
//Move Messages to a different folder once parsed




/*Design realted issues*/
//The header and the subject will both be needed at some point, so it is better to just fetch the entire message
//It would be faster to just fetch the header object and parse the subject line, but that would require use to mark it as read manually and then move to the processed folder
//iEnumarable