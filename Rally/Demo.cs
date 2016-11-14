using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation operation = new RallyOperation(RallyField.userName, RallyField.password);

            //operation.CreateUserStory("ListUS", "iterate through list", QueryField.WS_zScratch, QueryField.ST_SampleProject, QueryField.USER_Jostte, QueryField.IT_Iteration, "2");
            //operation.getWorkspaces();
            //operation.getScrumTeams();
            //operation.getUserStories(QueryField.WS_UCIT, QueryField.ST_lotteryWinners);
            //operation.getUSTA(QueryField.WS_UCIT, QueryField.ST_lotteryWinners);
            //operation.CreateTask("RefinedTask", "No more duplication of Authenticate", QueryField.USER_Jostte, "1", QueryField.US_9);
            operation.CreateUserStoryFromList(QueryField.WS_zScratch, QueryField.ST_SampleProject);


            Console.ReadLine();
        }
    }
}
