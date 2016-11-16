
        [Test, Ignore("Manual tests")]
        public void move_inbox_messages_gmail()
        {
            var _selectedMailBox = "INBOX";
            using (var _clientImap4 = new Imap4Client())
            {
                _clientImap4.ConnectSsl(_imapServerAddress, _imapPort);
                _clientImap4.LoginFast(_imapLogin, _imapPassword);
                
                _clientImap4.CreateMailbox("Processed");

                var mails = _clientImap4.SelectMailbox(_selectedMailBox);
                var ids = mails.Search("ALL");
                foreach (var id in ids)
                {
                    mails.MoveMessage(id, "Processed");
                }
                var mailsUndeleted = _clientImap4.SelectMailbox(_selectedMailBox);
                _clientImap4.Disconnect();
            }
        }
