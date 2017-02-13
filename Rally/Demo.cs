using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation operation = new RallyOperation(RallyConstant.UserName, RallyConstant.Password);

<<<<<<< HEAD
            operation.CreateUserStory(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject,"new user story", "testing with iteration", RallyQueryConstant.RallyUserJostte);
           
=======
            operation.CreateUserStory(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject, "ABC");
            
             
>>>>>>> origin/master
            Console.ReadLine();
        }
    }
}
