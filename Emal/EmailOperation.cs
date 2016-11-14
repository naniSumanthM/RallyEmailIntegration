using System;
using ActiveUp.Net.Mail;
using System.Collections.Generic;

namespace Email
{
    class EmailOperation
    {
        #region referenceCode
        public void getEmail()
        {
            Imap4Client imap = new Imap4Client();
            imap.ConnectSsl("imap-mail.outlook.com", 993);
            imap.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            Mailbox inbox = imap.SelectMailbox("inbox");
            MessageCollection mc = new MessageCollection();

            for (int i = 1; i <= inbox.MessageCount; i++)
            {
                Message newMsg = (inbox.Fetch.MessageObject(i));
                //Console.WriteLine(newMsg.Subject);
                mc.Add(newMsg);
            }
            Console.WriteLine("Start of foreach");
            foreach (var inboxItem in mc)
            {
                Console.WriteLine(inboxItem);
            }
            Console.WriteLine("End of foreach");

            #region comment
            //if (inbox.MessageCount > 0)
            //{
            //    Header header = inbox.Fetch.HeaderObject(1);
            //    Console.WriteLine(header.Subject);
            //}
            #endregion
        }
        #endregion

        #region fetchEmail()
        /// <summary>
        /// Method that scans an inbox and parses only the unread subject lines
        /// </summary>

        public void fetchEmail()
        {
            //Connect and Authenticate
            Imap4Client imap = new Imap4Client();
            imap.ConnectSsl("imap-mail.outlook.com", 993);
            imap.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            //setup Enviornment
            Mailbox inbox = imap.SelectMailbox("inbox");
            int[] unread = inbox.Search("UNSEEN");
            List<string> unreadList = new List<string>();
            MessageCollection mc = new MessageCollection();

            if (unread.Length > 0)
            {
                for (int i = 0; i < unread.Length; i++)
                {
                    Message msg = inbox.Fetch.MessageObject(unread[i]);
                    unreadList.Add(msg.ToString());
                    //mc.Add(msg);
                    //Console.WriteLine(msg.Subject);
                }
            }
            else
            {
                Console.WriteLine("No unread mail");
            }
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

        #region MailboxFolder

        public void createMailBox()
        {
            //try
            //{
            //Connect and Authenticate
            Imap4Client client = new Imap4Client();
            client.ConnectSsl("imap-mail.outlook.com", 993);
            client.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            //create mailbox
            //client.CreateMailbox("Processed");
            //Console.WriteLine("Created Mailbox");

            Mailbox inbox = client.SelectMailbox("inbox");
            MessageCollection mc = new MessageCollection();

            //iterate through the messages
            for (int i = 0; i < inbox.MessageCount; i++)
            {
                Message newMessage = inbox.Fetch.MessageObject(i);
                //inbox.MoveMessage(i, "Processed");
                mc.Add(newMessage);
                //Console.WriteLine(newMessage.BodyText.Text);
                //Header getSubject = inbox.Fetch.HeaderObject(n);
                //Console.WriteLine(getSubject.Subject);
            }

            foreach (var item in mc)
            {
                Console.WriteLine(item);
            }

            Console.WriteLine("Messages Moved");

            //}
            //catch (Imap4Exception)
            //{
            //    Console.WriteLine("Imap Exception");
            //}
            //catch (Exception)
            //{
            //    Console.WriteLine("Exception");
            //}

        }
        #endregion

    }
}

//Able to get the subjects
//Able to get the body

//Need to move the messages to a folder once processed?
//If the messages are moved, then the program does not search for unwanted email.
//The list will be a short iteration