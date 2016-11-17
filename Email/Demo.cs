using System;

namespace Email
{
    class Demo
    {
        static void Main(string[] args)
        {
            EmailOperation operation = new EmailOperation();

            //operation.FetchUnreadEmails();
            //operation.fetchUnreadSubjectLines();  
            //operation.moveMessages();
            //operation.moveUnreadEmail();

            operation.markAsRead();

            Console.ReadLine();
        }
    }
}
