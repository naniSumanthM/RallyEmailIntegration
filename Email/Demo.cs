﻿using System;

namespace Email
{
    class Demo
    {
        static void Main(string[] args)
        {
            EmailOperation operation = new EmailOperation();
            operation.moveUnreadEmail();

            Console.ReadLine();
        }
    }
}
