using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation operation = new RallyOperation(RALLY.RallyUserName, RALLY.RallyPassword);
            operation.GetProjectAdmins(RALLYQUERY.WorkspaceUcit);
            Console.ReadLine();
        }
    }
}




