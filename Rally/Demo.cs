using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation operation = new RallyOperation(RALLY.RallyUserName, RALLY.RallyPassword);
            operation.SyncThroughLabels(RALLYQUERY.WorkspaceUcit);
            
            Console.ReadLine();
        }
    }
}




