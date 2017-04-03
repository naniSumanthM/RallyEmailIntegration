using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyIntegration rallyRallyIntegration = new RallyIntegration(RallyConstant.UserName, RallyConstant.Password, "sumanth083@gmail.com", "iYmcmb24$");
            rallyRallyIntegration.SyncUserStories(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            Console.ReadLine();
        }
    }
}




