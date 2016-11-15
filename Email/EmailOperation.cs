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

            try
            {
                imap.ConnectSsl(Credential.outlookImapHost, Credential.outlookImapPort);
                imap.Login(Credential.outlookUserName, Credential.outlookPassword);

                //setup Enviornment
                Mailbox inbox = imap.SelectMailbox("inbox");
                int[] unread = inbox.Search("UNSEEN");
                Console.WriteLine("Unread Messages: " + unread.Length);

                //List to store ONLY the unread subject lines
                List<Message> unreadList = new List<Message>();

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
                Mailbox inbox = client.SelectMailbox("inbox");
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
                    Console.WriteLine("No Unread emails found");
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

    }
}

//11/15/16 toDO
//Mark as read manually
//copy and delete
//Move Messages to a different folder once parsed
