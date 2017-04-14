var items = client.Inbox.Fetch (0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.BodyStructure | MessageSummaryItems.Envelope);

foreach (var item in items) 
{
    // here's how to get the subject
    string subject = item.Envelope.Subject;

    // here's how to get the text body
    BodyPart part = item.TextBody;

    // here's how to get both attachments and inline attachments
    foreach (var attachment in item.BodyParts) 
    {
        MimeEntity entity = folder.GetBodyPart (item.UniqueId, attachment);
        // to save the content, it works exactly the same as in the GetMessage example
        //need to save this locally
    }
}