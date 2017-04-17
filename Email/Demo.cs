using System;

namespace Email
{
    class Demo
    {
        public static void Main(string[] args)
        {
            EmailOperation operation = new EmailOperation();
            operation.DownloadAttachmentsFileIoWay();

            Console.ReadLine();
        }
    }
}
