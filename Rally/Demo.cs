using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            //RallyOperation operation = new RallyOperation(RallyConstant.UserName, RallyConstant.Password);
            //operation.Sync(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            Sync rallySync = new Sync(RallyConstant.UserName, RallyConstant.Password, OutlookConstant.OutlookUsername, OutlookConstant.OutlookPassword);

            rallySync.SyncUserStories(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            Console.ReadLine();
        }
    }
}
