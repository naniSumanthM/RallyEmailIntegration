public void getAtttachments()
{
    using (var client = new ImapClient ()) {
        client.Connect(Constant.OutlookImapHost, Constant.ImapPort, SecureSocketOptions.SslOnConnect);
        client.Authenticate(Constant.GoogleUserName, Constant.GenericPassword);

        client.Inbox.Open(FolderAccess.ReadWrite);
        IList<UniqueId> uids = client.Inbox.Search(SearchQuery.All);
        foreach (UniqueId uid in uids) {
            MimeMessage message = client.Inbox.GetMessage(uid);

            foreach (MimeEntity attachment in message.Attachments) {
                // literally copied & pasted from the FAQ:
                string fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;

                if (string.IsNullOrEmpty (fileName)) {
                    // This attachment doesn't have a filename, I guess we'll skip it...
                }

                string path = Path.Combine ("C:\\Users\\maddirsh\\Desktop", fileName);

                // also literally copied and pasted from the FAQ:
                using (var stream = File.Create (path)) {
                    if (attachment is MessagePart) {
                        var rfc822 = (MessagePart) attachment;
            
                        rfc822.Message.WriteTo (stream);
                    } else {
                        var part = (MimePart) attachment;
            
                        part.ContentObject.DecodeTo (stream);
                    }
                }
            }
        }
    }
}