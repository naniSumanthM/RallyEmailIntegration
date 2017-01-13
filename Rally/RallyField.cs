using System;

namespace Rally
{
    class RallyField
    {
        /// <summary>
        /// Class that stores the fields relating to Rally enviornment
        /// CRUD operations can be performed on these fields
        /// </summary>

        public const string workSpace = "Workspace";
        public const string project = "Project";
        public const string owner = "Owner";
        public const string iteration = "Iteration";
        public const string description = "Description";
        public const string estimate = "Estimate";
        public const string planEstimate = "PlanEstimate";
        public const string hierarchicalRequirement = "HierarchicalRequirement";
        public const string workProduct = "WorkProduct";
        public const string smallTasks = "tasks";
        public const string capitalTasks = "Tasks";
        public const string state = "State";
        public const string formattedID = "FormattedID";
        public const string nameForWSorUSorTA = "Name";
        public const string artifact = "Artifact";
        public const string size = "Size";
        public const string content = "Content";
        public const string contentType = "ContentType";
        public const string attachment = "Attachment";
        public const string attachmentContent = "AttachmentContent";

        //Authentication fields:
        public const string serverID = "https://rally1.rallydev.com";
        public const bool projectScopeUp = true;
        public const bool projectScopeDown = true;
        public static bool allowSSO = false;

        //Login fields
        public const string userName = "maddirsh@mail.uc.edu";
        public const string password = "iYmcmb24";

        //Attachment Constants
        public const string emailAttachment = "Email Attachment";

    }
}
