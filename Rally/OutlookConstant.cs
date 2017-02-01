namespace Rally
{
    class OutlookConstant
    {
        /// <summary>
        /// Outlook.cs holds the constants required to use the Imap4Client libary
        /// </summary>

        public const int OutlookPort = 993;
        public const string OutlookHost = "imap-mail.outlook.com";
        public const string OutlookInboxFolder = "INBOX";
        public const string OutlookUnseenMessages = "UNSEEN";
        public const string OutlookSeenMessages = "SEEN";
        public const string OutlookProcessedFolder = "PROCESSED";
        public const string NoSubject = "No Subject";
        public const string NoBody = "No Body";

        //Sensitive - Remove after making repo public
        public const string OutlookUsername = "sumanthmaddirala@outlook.com";
        public const string OutlookPassword = "iYmcmb24";

    }
}
