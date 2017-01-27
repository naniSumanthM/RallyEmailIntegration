using System;

namespace Rally
{
    class Demo
    {
        static void Main(string[] args)
        {
            RallyOperation operation = new RallyOperation(RallyField.userName, RallyField.password);

            //operation.getWorkspaces();
            //operation.getScrumTeams();
            //operation.getUserStories(QueryField.WS_UCIT, QueryField.ST_lotteryWinners);
            //operation.getUSTA(QueryField.WS_UCIT, QueryField.ST_lotteryWinners);
            //operation.CreateTask("RefinedTask", "No more duplication of Authenticate", QueryField.USER_Jostte, "1", QueryField.US_9);
            //operation.addAttachmentsEliminateDuplicates(QueryField.WS_zScratch, QueryField.ST_SampleProject, "Duplicates allowed");
            //operation.addAttachmentsEliminateDuplicatesWithSimilarBase64Strings(QueryField.WS_zScratch, QueryField.ST_SampleProject, "Filter Duplicate Attachments");           
            //operation.SyncUserStories(QueryField.WS_zScratch, QueryField.ST_SampleProject);
            //operation.SyncUserStoriesAndLeaveMessageAsUnread(QueryField.WS_zScratch, QueryField.ST_SampleProject);
            //operation.CreateUserStory("ListUS","Description testing with images that maybe copied and pasted", QueryField.WS_zScratch, QueryField.ST_SampleProject, QueryField.USER_Jostte, QueryField.IT_Iteration, "2");
            //operation.syncUserStoriesWithAttachments();
            //operation.downloadAttachments();
            operation.Sync(QueryField.WS_zScratch, QueryField.ST_SampleProject);  

            Console.ReadLine();
        }
    }
}
