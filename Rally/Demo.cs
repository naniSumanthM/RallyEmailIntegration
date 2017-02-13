using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation operation = new RallyOperation(RallyConstant.UserName, RallyConstant.Password);
            //RallyIntegration integration = new RallyIntegration(RallyConstant.UserName, RallyConstant.Password);

            operation.Sync(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);  
            //operation.getIterations(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);


            operation.CreateUserStory("Hi", "Test with Feature", RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject, RallyQueryConstant.RallyUserJostte, RallyQueryConstant.Iteration);
            Console.ReadLine();
        }
    }
}
