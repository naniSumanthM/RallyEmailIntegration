using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            Sync rallySync = new Sync(RallyConstant.UserName, RallyConstant.Password, "sumanth083@gmail.com", "iYmcmb24$");
            rallySync.SyncUserStories(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            Console.ReadLine();
        }
    }
}




