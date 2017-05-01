using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation ro = new RallyOperation(RallyConstant.RallyUserName, RallyConstant.RallyPassword);
            ro.SyncThroughMultipleLabels();

            Console.ReadLine();
        }
    }
}




