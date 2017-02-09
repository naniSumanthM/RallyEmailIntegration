using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation operation = new RallyOperation(RallyConstant.UserName, RallyConstant.Password);

            operation.CreateUserStory(RallyQueryConstant.WorkspaceZScratch, RallyQueryConstant.ScrumTeamSampleProject, "ABC");
            
             
            Console.ReadLine();
        }
    }
}
