using System;

namespace Rally
{
    class QueryField
    {
        ///<summary>
        ///This class maintains all the fields that can be queried 
        ///Readonly runtime constant
        /// </summary>

        public const string reference = "_ref";
        public const string referenceObject = "_refObjectName";
        public const string workspaces = "Workspaces";
        public const string projects = "Projects";
        public const string lastUpdatDate = "LastUpdateDate";
        public const string dateGreaterThan = "2012-01-01";
        public const string usMessage = "User Story Owner:";
        public const string wsMessage = "Workspace: ";
        public const string taskName = "Task Name: ";
        public const string taskOwner = "Task Owner: ";
        public const string taskState = "Task State: ";
        public const string taskEstimate = "Task Estimate: ";
        public const string taskDescription = "Task Description: ";
        public const string webExceptionMessage = "Web Exception Hit";
        public const string taskMessage = "Tasks Not Found";

        //Teams and Projects
        public const string WS_zScratch = "/workspace/36903994748";
        public const string WS_UCIT = "/workspace/33809253647";
        public const string ST_lotteryWinners = "/project/57640961096";
        public const string ST_SampleProject = "/project/36903994832";
        public const string USER_Jostte = "/user/33809253641";
        public const string IT_Iteration = "iteration/70430073500";
        public const string US_9 = "/hierarchicalrequirement/70836533324";

    }
}
