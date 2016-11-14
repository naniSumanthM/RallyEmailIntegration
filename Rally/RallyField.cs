using System;

namespace Rally
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
        public static readonly string userName = "maddirsh@mail.uc.edu";
        public static readonly string password = "iYmcmb24";
        public const string serverID = "https://rally1.rallydev.com";
        public const bool projectScopeUp = true;
        public const bool projectScopeDown = true;

        //authenticate
        public static bool allowSSO = false;

    }
}
