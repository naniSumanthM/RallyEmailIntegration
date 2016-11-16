using System;

namespace Email
{
    class Demo
    {
        static void Main(string[] args)
        {
            EmailOperation operation = new EmailOperation();
            //operation.createMailBox();
            //operation.getAllEmail();
            operation.fetchUnreadSubjectLines();
            //operation.FetchUnreadEmails();
            //operation.moveMessages();
            //operation.move_inbox_messages_gmail();

            //operation.moveUnreadEmails();

            Console.ReadLine();
        }
    }
}
