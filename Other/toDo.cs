mark as read
copy
delete 
//ToDO:=
	retreive specific messages
	unreadHeader.ReceivedDate - to sort by date received

  //messageObject fetch Example
  Mailbox inbox = imap.SelectMailbox("inbox");
	if (inbox.MessageCount > 0)
	{
		Header header = inbox.Fetch.HeaderObject(1);
		this.AddLogEntry(string.Format("Subject: {0} From :{1} ", header.Subject, header.From.Email));
	}
	
	//throwing exceptions
	this.AddLogEntry(string.Format("Imap4 Error: {0}", iex.Message));
	this.AddLogEntry(string.Format("Failed: {0}", ex.Message));
	
	
	//Collection of mailboxes - for a more realisic situation
	List<Mailbox> mailboxCollection = new List<Mailbox>();
	
	
	//Unread Message loop
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
	
	//EnsureOutlookIsAuthenticated() - FIX
	    #region codeSnippet
        //Imap Object
        
        //Properties
        //public string outlookUsername { get; set; }
        //public string outlookPassword { get; set; }

        //Default Constructor
        //public EmailOperation(string OutlookUserName, string OutlookPassword)
        //{
        //    _imap = new Imap4Client();
        //    this.outlookUsername = OutlookUserName;
        //    this.outlookPassword = OutlookPassword;
        //    //this.EnsureOutlookIsAuthenticated();
        //}

        //Outlook Authentication --(Throws Exception)
        //private void EnsureOutlookIsAuthenticated()
        //{
        //    try
        //    {
        //        _imap.ConnectSsl(Credential.outlookImapHost, Credential.outlookImapPort);
        //        _imap.Login(this.outlookUsername, this.outlookPassword);
        //    }
        //    catch (SocketException)
        //    {
        //        throw new SocketException();
        //    }
        //}
        #endregion
	