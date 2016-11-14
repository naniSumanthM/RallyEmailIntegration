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
                imap.ConnectSsl(Credential.outlookImapHost, 993);
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
                Console.WriteLine("Imap-Exception");
            }
            catch (Exception)
            {
                Console.WriteLine("Exception");
            }
            finally
            {
                imap.Disconnect();
            }
        }
        #endregion

        #region fetchEmail()
        /// <summary>
        /// Method that scans an inbox and parses only the unread subject lines
        /// </summary>

        public void fetchUnreadSubjectLines()
        {
            //Authenticate with the Outlook Server
            Imap4Client imap = new Imap4Client();
            imap.ConnectSsl(Credential.outlookImapHost, 993);
            imap.Login(Credential.outlookUserName, Credential.outlookPassword);

            //setup Enviornment
            Mailbox inbox = imap.SelectMailbox("inbox");
            int[] unread = inbox.Search("UNSEEN");
            Console.WriteLine("Unread Messages: " + unread.Length);
            
            List<string> unreadList = new List<string>();
            //MessageCollection mc = new MessageCollection();

            if (unread.Length > 0)
            {
                for (int i = 1; i <= unread.Length; i++)
                {
                    Header header = (inbox.Fetch.HeaderObject(i));
                    Console.WriteLine(header.Subject);

                    //only printing - but not storing them in a list.

                    //Message msg = (inbox.Fetch.MessageObject(i));
                    //unreadList.Add(msg.ToString());
                    ////Console.WriteLine(msg.Subject);
                }
            }
            else
            {
                Console.WriteLine("No unread mail");
            }

            //foreach (var item in unreadList)
            //{
            //    Console.WriteLine(item.ToString());
            //}
        }
        #endregion

        #region getMessages()
        public void GetMessages()
        {
            Imap4Client imap = new Imap4Client();
            try
            {
                //Connect and Authenticate
                Imap4Client client = new Imap4Client();
                client.ConnectSsl("imap-mail.outlook.com", 993);
                client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

                //Stage the enviornment
                Mailbox inbox = client.SelectMailbox("inbox");
                MessageCollection mc = new MessageCollection();

                //iterate through the messages
                for (int n = 1; n < inbox.MessageCount + 1; n++)
                {
                    Message newMessage = inbox.Fetch.MessageObject(n);
                    Header getSubject = inbox.Fetch.HeaderObject(n);
                    //mc.Add(newMessage);
                    Console.WriteLine(getSubject.Subject);
                    Console.WriteLine(newMessage.BodyText.Text);

                    //Mailbox processed = new Mailbox();
                    //processed.
                }

                //returns the object form
                //foreach (var item in mc)
                //{
                //    Console.WriteLine(item);
                //}
            }
            catch (Imap4Exception)
            {
                Console.WriteLine("Exception IMap");
            }
            catch (Exception)
            {
                Console.WriteLine("EXception");
            }
            finally
            {
                if (imap.IsConnected)
                {
                    imap.Disconnect();
                }
            }
        }
        #endregion

        #region CreateMailBOX()

        public void createMailBox()
        {
            try
            {
                //Connect and Authenticate
                Imap4Client client = new Imap4Client();
                client.ConnectSsl(Credential.outlookImapHost, 993);
                client.Login(Credential.outlookUserName, Credential.outlookPassword);

                //create mailbox
                client.CreateMailbox("Mailbox-A");
                Console.WriteLine("Created Mailbox");
            }
            catch (Imap4Exception)
            {
                Console.WriteLine("Imap Exception");
            }
            catch (Exception)
            {
                Console.WriteLine("Exception");
            }
        }
        #endregion

    }
}

//Able to get the subjects
//Able to get the body

//Need to move the messages to a folder once processed?
//If the messages are moved, then the program does not search for unwanted email.
//The list will be a short iteration