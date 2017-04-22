using System;

namespace Rally
{
    class RallyConstant
    {
        /// <summary>
        /// RallyField.cs holds the constants needed to leverage the Rally.NET API
        /// </summary>

        //Server Constants
        public const string ServerId = "https://rally1.rallydev.com";
        public const bool ProjectScopeUp = true;
        public const bool ProjectScopeDown = true;
        public const bool AllowSso = false;

        //API Constants
        public const string WorkSpace = "Workspace";
        public const string Project = "Project";
        public const string Owner = "Owner";
        public const string Iteration = "Iteration";
        public const string Description = "Description";
        public const string Estimate = "Estimate";
        public const string PlanEstimate = "PlanEstimate";
        public const string HierarchicalRequirement = "HierarchicalRequirement";
        public const string WorkProduct = "WorkProduct";
        public const string TasksLowerCase = "tasks";
        public const string TasksUpperCase = "Tasks";
        public const string State = "State";
        public const string FormattedId = "FormattedID";
        public const string Name = "Name";
        public const string Artifact = "Artifact";
        public const string Size = "Size";
        public const string Content = "Content";
        public const string ContentType = "ContentType";
        public const string Attachment = "Attachment";
        public const string PortfolioItem = "PortfolioItem";
        public const string AttachmentContent = "AttachmentContent";
        public const string EmailAttachment = "Email Attachment";
        public const string UserStoryUrlFormat = "https://rally1.rallydev.com/#/detail/userstory/";
        public const string HexColor = "#4ef442";
        public const string SlackApiToken = "https://hooks.slack.com/services/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";
        public const string SlackNotificationText = "*Rally Notification*";
        public const string SlackChannel = "#general";
        public const string SlackUser = "sumanth";
    }
}
