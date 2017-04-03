using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyIntegration rallyRallyIntegration = new RallyIntegration(RallyConstant.UserName, RallyConstant.Password, EmailConstant.GoogleUsername, EmailConstant.GenericPassword);
            rallyRallyIntegration.SyncUserStories(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            Console.ReadLine();
        }
    }
}




