using System;

namespace Rally_Email_Integration
{
    class RallyField
    {
        /// <summary>
        /// Class that stores the fields relating to Rally enviornment
        /// CRUD operations can be performed on these fields
        /// </summary>


        public static readonly string workSpace = "Workspace";
        public static readonly string project = "Project";
        public static readonly string owner = "Owner";
        public static readonly string iteration = "Iteration";
        public static readonly string description = "Description";
        public static readonly string estimate = "Estimate";
        public static readonly string planEstimate = "PlanEstimate";
        public static readonly string hierarchicalRequirement = "HierarchicalRequirement";
        public static readonly string workProduct = "WorkProduct";
        public static readonly string smallTasks = "tasks";
        public static readonly string capitalTasks = "Tasks";
        public static readonly string state = "State";
        public static readonly string formattedID = "FormattedID";
        public const string nameForWSorUSorTA = "Name";

        //constants for the userCredentials- how to completly hide them.
        public static readonly string RuserName = "maddirsh@mail.uc.edu";
        public static readonly string Rpassword = "iYmcmb24";
        public const string serverID = "https://rally1.rallydev.com";
        public const bool projectScopeUp = true;
        public const bool projectScopeDown = true;

        //authenticate
        public static bool allowSSO = false;

        //Outlook 
        public const string outlookUSName = "sumanthmaddirala@outlook.com";
        public const string outlookPassword = "iYmcmb24";
        public const string outlookImapHost = "imap-mail.outlook.com";
        public const int outlookImapPort = 993;

        //Teams and Projects
        public const string WS_zScratch = "/workspace/36903994748";
        public const string WS_UCIT = "/workspace/33809253647";
        public const string ST_lotteryWinners = "/project/57640961096";
        public const string ST_SampleProject = "/project/36903994832";
        public const string USER_Jostte = "/user/33809253641";

    }
}
