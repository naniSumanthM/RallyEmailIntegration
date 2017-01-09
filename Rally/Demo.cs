using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation operation = new RallyOperation(RallyField.userName, RallyField.password);

            #region: Test Method Calls
            //operation.CreateUserStory("ListUS", "iterate through list", QueryField.WS_zScratch, QueryField.ST_SampleProject, QueryField.USER_Jostte, QueryField.IT_Iteration, "2");
            //operation.getWorkspaces();
            //operation.getScrumTeams();
            //operation.getUserStories(QueryField.WS_UCIT, QueryField.ST_lotteryWinners);
            //operation.getUSTA(QueryField.WS_UCIT, QueryField.ST_lotteryWinners);
            //operation.CreateTask("RefinedTask", "No more duplication of Authenticate", QueryField.USER_Jostte, "1", QueryField.US_9);
            //operation.UserStorySyncRefined(QueryField.WS_zScratch, QueryField.ST_SampleProject);
            //operation.UserStorySync(QueryField.WS_zScratch, QueryField.ST_SampleProject);
            #endregion

            operation.userStoryWithMultipleAttachments(QueryField.WS_zScratch, QueryField.ST_SampleProject, "multiAttachemnt3", "The attachment needs to be created for how many ever files there are in the folder");
            Console.ReadLine();
        }
    }
}
