using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            //RallyOperation operation = new RallyOperation(RallyConstant.UserName, RallyConstant.Password);
            //operation.Sync(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            RallyIntegration rallyIntegration = new RallyIntegration(RallyConstant.UserName, RallyConstant.Password, EmailConstant.GoogleUsername, EmailConstant.GenericPassword);
            rallyIntegration.SyncUserStories(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            Console.ReadLine();
        }
    }
}




