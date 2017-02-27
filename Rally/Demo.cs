using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            Sync rallySync = new Sync(RallyConstant.UserName, RallyConstant.Password, OutlookConstant.OutlookUsername, OutlookConstant.OutlookPassword);
            rallySync.SyncUserStories(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            //RallyOperation o = new RallyOperation(RallyConstant.UserName, RallyConstant.Password);
            //o.downlodInlineAttachments(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            Console.ReadLine();
        }
    }
}




