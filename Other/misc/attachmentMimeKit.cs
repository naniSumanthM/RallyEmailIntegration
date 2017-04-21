using (var client = new ImapClient())
{
    var credentials = new NetworkCredential("myID", "myPassword");
    var uri = new Uri("imaps://imap.gmail.com");

    using (var cancel = new CancellationTokenSource())
    {
        client.Connect(uri, cancel.Token);
        client.Authenticate(credentials, cancel.Token);

        var inbox = client.Inbox;
        inbox.Open(FolderAccess.ReadOnly, cancel.Token);

        int MailList = 10;
        int ListCount = Math.Min(inbox.Count, MailList);

        System.IO.Directory.CreateDirectory(@"d:\mailAttachmentsTemp\");
        for (int i = inbox.Count - 1; i > inbox.Count - ListCount; i--)
        {
            var message = inbox.GetMessage(i, cancel.Token);

            if (message.Attachments.Count() > 0)
            {
                Console.WriteLine("Have Attachments!");                            
                System.IO.DirectoryInfo dir = System.IO.Directory.CreateDirectory(@"d:\mailAttachmentsTemp\" + message.MessageId);
                foreach (var mp in message.Attachments)
                {
                    using (var stream = System.IO.File.Create(dir.FullName + @"\" + mp.FileName))
                    {
                        mp.ContentObject.DecodeTo(stream);
                        stream.Close();
                    }
                }
            }
        }
        client.Disconnect(true, cancel.Token);
    }
}