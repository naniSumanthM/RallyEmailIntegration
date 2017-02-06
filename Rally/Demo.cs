using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation operation = new RallyOperation(RallyConstant.UserName, RallyConstant.Password);

            //operation.Sync(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);  
            //operation.CreateUserStory("IterationStory",
            //                           "We will pass in an already created iteration from Rally", 
            //                           RallyQueryConstant.WorkspaceZScratch, 
            //                           RallyQueryConstant.ScrumTeamSampleProject, 
            //                           RallyQueryConstant.RallyUserJostte, 
            //                           RallyQueryConstant.Iteration, 
            //                           "2");


            operation.getIterations(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject);

            Console.ReadLine();
        }
    }
}
