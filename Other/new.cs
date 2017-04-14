foreach (var message in items)
{
    foreach (var attachment in message.BodyParts)
    {
        MimeKit.MimePart mime = (MimeKit.MimePart)client.Inbox.GetBodyPart(message.UniqueId, attachment);
        string fileName = mime.FileName;

        if (string.IsNullOrEmpty(fileName))
        {
            fileName = string.Format("unnamed-{0}", ++unnamed);
        }

        FormatOptions options = FormatOptions.Default.Clone();
        options.ParameterEncodingMethod = ParameterEncodingMethod.Rfc2047;

        using (FileStream stream = File.Create(Path.Combine("C:\\Users\\maddirsh\\Desktop\\MimeKit\\", fileName)))
        {
            mime.ContentObject.DecodeTo(stream);
        }

        Console.WriteLine("End");
    }
}