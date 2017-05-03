using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation operation = new RallyOperation(RallyConstant.RallyUserName, RallyConstant.RallyPassword);
            operation.SyncThroughLabels(RallyQueryConstant.WorkspaceUcit);
            
            Console.ReadLine();
        }
    }
}




