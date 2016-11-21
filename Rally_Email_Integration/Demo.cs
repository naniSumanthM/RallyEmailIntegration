using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rally_Email_Integration
{
    class Demo
    {
        static void Main(string[] args)
        {
            Program p = new Program(RallyField.RuserName, RallyField.Rpassword);
            p.syncUserStories();

            Console.ReadLine();
        }
    }
}
