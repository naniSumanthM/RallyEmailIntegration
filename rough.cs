mark as read
copy
delete 
retreive specific messages
unreadHeader.ReceivedDate


  Mailbox inbox = imap.SelectMailbox("inbox");
	if (inbox.MessageCount > 0)
	{
		Header header = inbox.Fetch.HeaderObject(1);
		this.AddLogEntry(string.Format("Subject: {0} From :{1} ", header.Subject, header.From.Email));
	}
	
	this.AddLogEntry(string.Format("Imap4 Error: {0}", iex.Message));
	this.AddLogEntry(string.Format("Failed: {0}", ex.Message));
	
	
	//Collection of mailboxes
	List<Mailbox> mailboxCollection = new List<Mailbox>();
	
	
	for(int i=1; i<x+1; i++ )
	{
		//code goes here
	}
	
	for(int i=1; i<x; i++)
	{
		//code goes here
	}
	
	for(int i=0; i<=x; i++)
	{
		//code goes here
	}