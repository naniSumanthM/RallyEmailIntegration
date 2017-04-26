using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            #region System.Mail.NET
            //RallyIntegration rallyIntegration = new RallyIntegration(RallyConstant.RallyUserName, RallyConstant.RallyPassword, EmailConstant.GoogleUsername, EmailConstant.GenericPassword);
            //rallyIntegration.SyncUserStories(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            #endregion

            RallyIntegrationUsingMimeKit r = new RallyIntegrationUsingMimeKit(RallyConstant.RallyUserName, RallyConstant.RallyPassword, EmailConstant.GoogleUsername, EmailConstant.GenericPassword);
            r.SyncUsingMimeKit(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            Console.ReadLine();
        }
    }
}




