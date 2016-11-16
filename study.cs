        #region fetchEmail()
        /// <summary>
        /// Method that scans an inbox and parses only the unread subject lines
        /// </summary>

        public void fetchUnreadSubjectLines()
        {
            //Authenticate with the Outlook Server
            Imap4Client imap = new Imap4Client();
            imap.ConnectSsl(Credential.outlookImapHost, 993);
            imap.Login("sumanthmaddirala@outlook.com", "iYmcmb24");

            //setup Enviornment
            Mailbox inbox = imap.SelectMailbox("inbox");
            int[] unread = inbox.Search("UNSEEN");
            Console.WriteLine("Unread Messages: " + unread.Length);
            
            List<string> unreadList = new List<string>();
            List<Header> unreadHeder = new List<Header>();
            //MessageCollection mc = new MessageCollection();

            if (unread.Length > 0)
            {
                for (int i = 1; i <= unread.Length; i++)
                {
                    //Message msg = (inbox.Fetch.MessageObject(i));
                    //unreadList.Add(msg.ToString());
                    //Console.WriteLine(msg.Subject);

                    Header h = (inbox.Fetch.HeaderObject(i));
                    unreadHeder.Add(h);
                    //Console.WriteLine(h.Subject);
                    
                }

                //loop through the unread header
                //the ideal thing would be to fetch the header and the subject, but then again - its like reading email inside Rally
                foreach (Header item in unreadHeder)
                {
                    Console.WriteLine(item.Subject);
                }
            }
            else
            {
                Console.WriteLine("No unread mail");
            }

        }