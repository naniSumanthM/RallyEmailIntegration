using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyTest rallySync = new RallyTest(RallyConstant.UserName, RallyConstant.Password, OutlookConstant.OutlookUsername, OutlookConstant.OutlookPassword);
            //rallySync.SyncUserStories(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);
            rallySync.helloQ();


            //RallyOperation o = new RallyOperation(RallyConstant.UserName, RallyConstant.Password);
            //o.Sync(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            Console.ReadLine();
        }
    }
}




