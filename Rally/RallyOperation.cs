﻿using System.Linq;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using MimeKit.Text;
using ServiceStack;
using Slack.Webhooks;
using static System.String;

namespace Rally
{
    #region: System Libraries
    using System;
    using System.IO;
    using RestApi;
    using RestApi.Json;
    using RestApi.Response;
    using System.Collections.Generic;
    using System.Net;
    using System.Drawing;
    #endregion
    class RallyOperation
    {
        RallyRestApi _rallyRestApi;
        public const string ServerName = RALLY.ServerId;
        public string UserName { get; set; }
        public string Password { get; set; }
        public RallyOperation(string userName, string password)
        {
            _rallyRestApi = new RallyRestApi();
            this.UserName = userName;
            this.Password = password;
            this.EnsureRallyIsAuthenticated();
        }
        private void EnsureRallyIsAuthenticated()
        {
            if (this._rallyRestApi.AuthenticationState != RallyRestApi.AuthenticationResult.Authenticated)
            {
                _rallyRestApi.Authenticate(this.UserName, this.Password, ServerName, null, RALLY.AllowSso);
            }
        }

        #region: Query Workspaces

        /// <summary>
        /// Returns all the existing workspaces in Rally
        /// </summary>
        public void GetWorkspaces()
        {
            //Authenticate
            this.EnsureRallyIsAuthenticated();

            //instantiate a DynamicJsonObject obj
            DynamicJsonObject djo = _rallyRestApi.GetSubscription(RALLYQUERY.Workspaces);
            Request workspaceRequest = new Request(djo[RALLYQUERY.Workspaces]);

            try
            {
                //query for the workspaces
                QueryResult returnWorkspaces = _rallyRestApi.Query(workspaceRequest);

                //iterate through the list and return the list of workspaces
                foreach (var value in returnWorkspaces.Results)
                {
                    var workspaceReference = value[RALLYQUERY.Reference];
                    var workspaceName = value[RALLY.Name];
                    Console.WriteLine(RALLYQUERY.WorkspaceMessage + workspaceName);
                }
            }
            catch (WebException)
            {
                Console.WriteLine(RALLYQUERY.WebExceptionMessage);
            }
        }
        #endregion

        #region: Query Scrum Teams
        /// <summary>
        /// Retreives all the scrum teams within the Rally Enviornment
        /// </summary>
        public void GetScrumTeams()
        {
            this.EnsureRallyIsAuthenticated();

            DynamicJsonObject dObj = _rallyRestApi.GetSubscription(RALLYQUERY.Workspaces);

            try
            {
                Request workspaceRequest = new Request(dObj[RALLYQUERY.Workspaces]);
                QueryResult workSpaceQuery = _rallyRestApi.Query(workspaceRequest);

                foreach (var workspace in workSpaceQuery.Results)
                {
                    Request projectRequest = new Request(workspace[RALLYQUERY.Projects]);
                    projectRequest.Fetch = new List<string> { RALLY.Name };

                    //Query for the projects
                    QueryResult projectQuery = _rallyRestApi.Query(projectRequest);
                    foreach (var project in projectQuery.Results)
                    {
                        Console.WriteLine(project[RALLY.Name]);
                    }
                }
            }
            catch (WebException)
            {
                Console.WriteLine(RALLYQUERY.WebExceptionMessage);
            }
        }

        #endregion

        #region: Query User Stories
        /// <summary>
        /// When passes a workspace and a project, all the userstories are returned within that enviornment
        /// Any userstory without an owner is invalid and will not be queried
        /// <definition>Project Scoping: UP - Will exclude projects that are above the default project</definition> 
        /// <definition>Project Scoping: DOWN - Will include child projects</definition> 
        /// </summary>
        /// <param name="workspaceRef"></param>
        /// <param name="projectRef"></param>

        public void GetUserStories(string workspaceRef, string projectRef)
        {
            //Authenticate
            this.EnsureRallyIsAuthenticated();

            //setup the userStoryRequest
            Request userStoryRequest = new Request(RALLY.HierarchicalRequirement);
            userStoryRequest.Workspace = workspaceRef;
            userStoryRequest.Project = projectRef;
            userStoryRequest.ProjectScopeUp = RALLY.ProjectScopeUp;
            userStoryRequest.ProjectScopeDown = RALLY.ProjectScopeDown;

            //fetch data from the story request
            userStoryRequest.Fetch = new List<string>()
            {
                RALLY.FormattedId, RALLY.Name, RALLY.Owner
            };

            try
            {
                //query the items in the list
                userStoryRequest.Query = new Query(RALLYQUERY.LastUpdatDate, Query.Operator.GreaterThan, RALLYQUERY.DateGreaterThan);
                QueryResult userStoryResult = _rallyRestApi.Query(userStoryRequest);

                //iterate through the userStory Collection
                foreach (var userStory in userStoryResult.Results)
                {
                    var userStoryOwner = userStory[RALLY.Owner];
                    if (userStoryOwner != null)
                    {
                        var USOwner = userStoryOwner[RALLYQUERY.ReferenceObject];
                        Console.WriteLine(userStory[RALLY.FormattedId] + ":" + userStory[RALLY.Name] + Environment.NewLine + RALLYQUERY.UserStoryMessage + USOwner + Environment.NewLine);
                    }
                }
            }
            catch (WebException)
            {
                Console.WriteLine(RALLYQUERY.WebExceptionMessage);
            }
        }
        #endregion

        #region: Query User Stories and Tasks
        /// <summary>
        /// When provided with a workspace and a project, the method will return all the user stories along with the tasks and their details
        /// </summary>
        /// <param name="workspaceRef"></param>
        /// <param name="projectRef"></param>

        public void GetUserStoriesWithTasks(string workspaceRef, string projectRef)
        {
            //Authenticate
            this.EnsureRallyIsAuthenticated();

            //stage the request (not using the getters and setters from the Rally Enviornment class
            Request userStoryRequest = new Request(RALLY.HierarchicalRequirement);
            userStoryRequest.Workspace = workspaceRef;
            userStoryRequest.Project = projectRef;
            userStoryRequest.ProjectScopeUp = RALLY.ProjectScopeUp;
            userStoryRequest.ProjectScopeDown = RALLY.ProjectScopeDown;

            //fetch US data in the form of a list
            userStoryRequest.Fetch = new List<string>()
            {
                RALLY.FormattedId, RALLY.Name, RALLY.TasksUpperCase, RALLY.Estimate, RALLY.State, RALLY.Owner
            };

            //Userstory Query
            userStoryRequest.Query = (new Query(RALLYQUERY.LastUpdatDate, Query.Operator.GreaterThan, RALLYQUERY.DateGreaterThan));

            try
            {
                //query for the items in the list
                QueryResult userStoryResult = _rallyRestApi.Query(userStoryRequest);

                //iterate through the query results
                foreach (var userStory in userStoryResult.Results)
                {
                    var userStoryOwner = userStory[RALLY.Owner];
                    if (userStoryOwner != null) //return only US who have an assigned owner
                    {
                        var USOwner = userStoryOwner[RALLYQUERY.ReferenceObject];
                        Console.WriteLine(userStory[RALLY.FormattedId] + ":" + userStory[RALLY.Name]);
                        Console.WriteLine(RALLYQUERY.UserStoryMessage + USOwner);
                    }

                    //Task Request
                    Request taskRequest = new Request(userStory[RALLY.TasksUpperCase]);
                    QueryResult taskResult = _rallyRestApi.Query(taskRequest);
                    if (taskResult.TotalResultCount > 0)
                    {
                        foreach (var task in taskResult.Results)
                        {
                            var taskName = task[RALLY.Name];
                            var owner = task[RALLY.Owner];
                            var taskState = task[RALLY.State];
                            var taskEstimate = task[RALLY.Estimate];
                            //var taskDescription = task[RallyField.description];

                            if (owner != null)
                            {
                                var ownerName = owner[RALLYQUERY.ReferenceObject];
                                Console.WriteLine(RALLYQUERY.TaskName + taskName + Environment.NewLine + RALLYQUERY.TaskOwner + ownerName + Environment.NewLine + RALLYQUERY.TaskState + taskState + Environment.NewLine + RALLYQUERY.TaskEstimate + taskEstimate);
                                //Console.WriteLine(QueryField.taskDescription + taskDescription);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(RALLYQUERY.TaskMessage);
                    }
                }
            }
            catch (WebException)
            {
                Console.WriteLine(RALLYQUERY.WebExceptionMessage);
            }

        }

        #endregion

        #region: Query Iterations
        public void GetIterations(string workspace, string project)
        {
            this.EnsureRallyIsAuthenticated();

            Request iterationRequest = new Request(RALLY.Iteration);
            iterationRequest.Workspace = workspace;
            iterationRequest.Project = project;
            iterationRequest.ProjectScopeUp = RALLY.ProjectScopeUp;
            iterationRequest.ProjectScopeDown = RALLY.ProjectScopeDown;

            try
            {
                iterationRequest.Fetch = new List<string>()
                {
                 RALLY.Name
                };

                iterationRequest.Query = new Query(RALLY.Project, Query.Operator.Equals, RALLYQUERY.ScrumTeamSampleProject);
                QueryResult queryResult = _rallyRestApi.Query(iterationRequest);
                foreach (var iteration in queryResult.Results)
                {
                    Console.WriteLine(iteration[RALLY.Name]);
                }

            }
            catch (WebException e)
            {
                Console.WriteLine(e.Message);
            }
        }
        #endregion

        #region: create Userstory
        /// <summary>
        /// Creates the userstory with a feature or iteration
        /// Both feature and iteration are read only fields
        /// </summary>
        /// <param name="workspace"></param>
        /// <param name="project"></param>
        /// <param name="userstory"></param>
        /// <param name="userstoryDescription"></param>
        /// <param name="userstoryOwner"></param>

        public void CreateUserStory(string workspace, string project, string userstory, string userstoryDescription, string userstoryOwner)
        {
            //authenticate
            this.EnsureRallyIsAuthenticated();

            //DynamicJsonObject
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RALLY.WorkSpace] = workspace;
            toCreate[RALLY.Project] = project;
            toCreate[RALLY.Name] = userstory;
            toCreate[RALLY.Description] = userstoryDescription;
            toCreate[RALLY.Owner] = userstoryOwner;
            toCreate[RALLY.PlanEstimate] = "1";
            toCreate[RALLY.PortfolioItem] = RALLYQUERY.FeatureShareProject;
            //toCreate[RALLY.Iteration] = usIteration;

            try
            {
                CreateResult createUserStory = _rallyRestApi.Create(RALLY.HierarchicalRequirement, toCreate);
                Console.WriteLine("Created Userstory: " + createUserStory.Reference);
            }
            catch (WebException e)
            {
                Console.WriteLine(e.Message);
            }
        }

        #endregion

        #region: Create Task
        /// <summary>
        /// When passed the task specfications along with the us reference, the method will create a task and attach it to an existing userStory.
        /// User story contians the reference to the project and the workspace
        /// </summary>
        /// <param name="taskName"></param>
        /// <param name="taskDescription"></param>
        /// <param name="taskOwner"></param>
        /// <param name="taskEstimate"></param>
        /// <param name="userStoryReference"></param>

        public void CreateTask(string taskName, string taskDescription, string taskOwner, string taskEstimate, string userStoryReference)
        {
            this.EnsureRallyIsAuthenticated();

            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RALLY.Name] = taskName;
            toCreate[RALLY.Description] = taskDescription;
            toCreate[RALLY.Owner] = taskOwner;
            toCreate[RALLY.Estimate] = taskEstimate;
            toCreate[RALLY.WorkProduct] = userStoryReference;

            //create a task and attach it to a userStory
            try
            {
                Console.WriteLine("<<Creating TA>>");
                CreateResult createTask = _rallyRestApi.Create(RALLY.TasksLowerCase, toCreate);
                Console.WriteLine("<<Created TA>>");
            }
            catch (WebException)
            {

                Console.WriteLine(RALLYQUERY.WebExceptionMessage);
            }

        }

        #endregion

        #region: fileToBase64
        public static string FileToBase64(string attachment)
        {
            Byte[] attachmentBytes = File.ReadAllBytes(attachment);
            string base64EncodedString = Convert.ToBase64String(attachmentBytes);
            return base64EncodedString;
        }
        #endregion

        #region: associate attanchment to user story using reference
        ///<summary>
        ///Pushes an attachment file to an existing user story
        /// </summary>

        public void AddAttachmentToUserStory()
        {
            this.EnsureRallyIsAuthenticated();

            string storyReference = "https://rally1.rallydev.com/slm/webservice/v2.0/hierarchicalrequirement/112831486000";

            //Process - Attaching the image to the user story

            // Read In Image Content
            string imageFilePath = "C:\\Users\\maddirsh\\Desktop\\";
            string imageFileName = "HomeController.cs";
            string fullImageFile = imageFilePath + imageFileName;

            // Convert Image to Base64 format
            string imageBase64String = FileToBase64(fullImageFile);

            // Length calculated from Base64String converted back
            var imageNumberBytes = Convert.FromBase64String(imageBase64String).Length;

            // DynamicJSONObject for AttachmentContent
            DynamicJsonObject myAttachmentContent = new DynamicJsonObject();
            myAttachmentContent["Content"] = imageBase64String; //string 

            try
            {
                CreateResult myAttachmentContentCreateResult = _rallyRestApi.Create("AttachmentContent", myAttachmentContent);
                string myAttachmentContentRef = myAttachmentContentCreateResult.Reference;
                Console.WriteLine("Created: " + myAttachmentContentRef);

                // DynamicJSONObject for Attachment Container
                DynamicJsonObject myAttachment = new DynamicJsonObject();
                myAttachment["Artifact"] = storyReference;
                myAttachment["Content"] = myAttachmentContentRef;
                myAttachment["Name"] = "AttachmentFromREST.png";
                myAttachment["Description"] = "Attachment Desc";
                myAttachment["ContentType"] = "image/png";
                myAttachment["Size"] = imageNumberBytes;

                CreateResult myAttachmentCreateResult = _rallyRestApi.Create("Attachment", myAttachment);
            }
            catch (Exception e)
            {
                Console.WriteLine("Unhandled exception occurred: " + e.StackTrace);
                Console.WriteLine(e.Message);
            }
        }
        #endregion

        #region SyncUsingMimeKit
        public void SyncUsingMimeKit(string workspace, string project)
        {
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RALLY.WorkSpace] = workspace;
            toCreate[RALLY.Project] = project;
            DynamicJsonObject attachmentContent = new DynamicJsonObject();
            DynamicJsonObject attachmentContainer = new DynamicJsonObject();
            CreateResult createUserStory;
            CreateResult attachmentContentCreateResult;
            CreateResult attachmentContainerCreateResult;

            string[] allAttachments;
            Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();
            string emailSubject;
            string emailBody;
            string userStoryReference;
            int anotherOne = 0;
            string base64String;
            string attachmentFileName;
            string fileName;

            EnsureRallyIsAuthenticated();

            using (var client = new ImapClient())
            {
                client.ServerCertificateValidationCallback = (s, c, ch, e) => true;
                client.Connect(EMAIL.GoogleImapHost, EMAIL.ImapPort, SecureSocketOptions.SslOnConnect);
                client.AuthenticationMechanisms.Remove(EMAIL.GoogleOAuth);
                client.Authenticate(EMAIL.GoogleUsername, EMAIL.GenericPassword);

                client.Inbox.Open(FolderAccess.ReadWrite);
                IMailFolder inboxFolder = client.GetFolder("Inbox");
                IList<UniqueId> uids = client.Inbox.Search(SearchQuery.All);

                foreach (UniqueId uid in uids)
                {
                    MimeMessage message = inboxFolder.GetMessage(uid);
                    emailSubject = message.Subject;
                    emailBody = message.TextBody;

                    if (emailSubject.IsEmpty())
                    {
                        emailSubject = "<No Subject User Story>";
                    }

                    toCreate[RALLY.Name] = (emailSubject);
                    toCreate[RALLY.Description] = (emailBody);
                    createUserStory = _rallyRestApi.Create(RALLY.HierarchicalRequirement, toCreate);

                    foreach (MimeEntity attachment in message.BodyParts)
                    {
                        string attachmentFile = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                        string attachmentFilePath = Concat(STORAGE.MimeKitAttachmentsDirectory, Path.GetFileName(attachmentFile));

                        if (!IsNullOrWhiteSpace(attachmentFile))
                        {
                            if (File.Exists(attachmentFilePath))
                            {
                                string extension = Path.GetExtension(attachmentFilePath);
                                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(attachmentFilePath);
                                attachmentFile = Format(fileNameWithoutExtension + "-{0}" + "{1}", ++anotherOne,
                                    extension);
                                attachmentFilePath = Path.Combine(STORAGE.MimeKitAttachmentsDirectory,
                                    attachmentFile);
                            }

                            using (var attachmentStream = File.Create(attachmentFilePath))
                            {
                                MimeKit.MimePart part = (MimeKit.MimePart)attachment;
                                part.ContentObject.DecodeTo(attachmentStream);
                            }

                            Console.WriteLine("Downloaded: " + attachmentFile);
                        }
                    }

                    allAttachments = Directory.GetFiles(STORAGE.MimeKitAttachmentsDirectory);
                    foreach (string file in allAttachments)
                    {
                        base64String = FileToBase64(file);
                        attachmentFileName = Path.GetFileName(file);
                        fileName = Empty;

                        if (!(attachmentsDictionary.TryGetValue(base64String, out fileName)))
                        {
                            Console.WriteLine("Added to Dictionary: " + file);
                            attachmentsDictionary.Add(base64String, attachmentFileName);
                        }

                        File.Delete(file);
                    }

                    //for each email message - upload it to Rally
                    foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
                    {
                        try
                        {
                            //create attachment content
                            attachmentContent[RALLY.Content] = attachmentPair.Key;
                            attachmentContentCreateResult = _rallyRestApi.Create(RALLY.AttachmentContent,
                                attachmentContent);
                            userStoryReference = attachmentContentCreateResult.Reference;

                            //create attachment contianer
                            attachmentContainer[RALLY.Artifact] = createUserStory.Reference;
                            attachmentContainer[RALLY.Content] = userStoryReference;
                            attachmentContainer[RALLY.Name] = attachmentPair.Value;
                            attachmentContainer[RALLY.Description] = RALLY.EmailAttachment;
                            attachmentContainer[RALLY.ContentType] = "file/";

                            //Create & associate the attachment
                            attachmentContainerCreateResult = _rallyRestApi.Create(RALLY.Attachment,
                                attachmentContainer);
                        }
                        catch (WebException e)
                        {
                            Console.WriteLine("Attachment: " + e.Message);
                        }
                    }
                    attachmentsDictionary.Clear();

                    Console.WriteLine("User Story: " + message.Subject);
                }
            }
        }

        #endregion

        #region SyncThroughLabels
        /// <summary>
        /// Rally is authenticated in the constructor
        /// </summary>
        /// <param name="workspace"></param>

        public void SyncThroughLabels(string workspace)
        {
            #region variables
            SlackClient _slackClient = new SlackClient(SLACK.SlackApiToken, 100);
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RALLY.WorkSpace] = workspace;
            DynamicJsonObject attachmentContent = new DynamicJsonObject();
            DynamicJsonObject attachmentContainer = new DynamicJsonObject();
            CreateResult createUserStory;
            CreateResult attachmentContentCreateResult;
            CreateResult attachmentContainerCreateResult;
            string[] allAttachments;
            Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();
            string userStorySubject;
            string userStoryDescription;
            string userStoryReference;
            string attachmentReference;
            int anotherOne = 0;
            string base64String;
            string attachmentFileName;
            string fileName;
            string _objectId;
            string _userStoryUrl;
            string _slackAttachmentString;
            string slackChannel;
            #endregion

            using (ImapClient client = new ImapClient())
            {
                AuthenticateWithGoogleImap(client);

                client.Inbox.Open(FolderAccess.ReadWrite);
                IMailFolder parentFolder = client.GetFolder(EMAIL.EnrollmentStudentServicesFolder);
                IMailFolder processedFolder = parentFolder.GetSubfolder(RALLYQUERY.ProcessedEnrollmentStudentServices);

                foreach (IMailFolder childFolder in parentFolder.GetSubfolders())
                {
                    #region Folders
                    if (childFolder.Name.Equals(RALLYQUERY.GmailFolderCatalyst2016))
                    {
                        toCreate[RALLY.Project] = RALLYQUERY.ProjectCatalyst2016;
                        slackChannel = SLACK.Channelcatalyst2016;
                    }
                    else if (childFolder.Name.Equals(RALLYQUERY.GmailFolderHonorsEnhancements))
                    {
                        toCreate[RALLY.Project] = RALLYQUERY.ProjectHonorsEnhancements;
                        slackChannel = SLACK.ChannelHonorsEnhancements;
                    }
                    else if (childFolder.Name.Equals(RALLYQUERY.GmailFolderPalHelp))
                    {
                        toCreate[RALLY.Project] = RALLYQUERY.ProjectPalHelp;
                        slackChannel = SLACK.ChannelPalHelp;
                    }
                    else if (childFolder.Name.Equals(RALLYQUERY.GmailFolderPciAzureTouchNetImplementation))
                    {
                        toCreate[RALLY.Project] = RALLYQUERY.ProjectPciAzureTouchNetImplementation;
                        slackChannel = SLACK.ChannelAzureTouchNet;
                    }
                    else
                    {
                        toCreate[RALLY.Project] = RALLYQUERY.ProjectScrumptious;
                        slackChannel = SLACK.ChannelScrumptious;
                    }
                    #endregion

                    Console.WriteLine(childFolder.Name);
                    childFolder.Open(FolderAccess.ReadWrite);
                    IList<UniqueId> childFolderMsgUniqueIds = childFolder.Search(SearchQuery.NotSeen);

                    if (childFolderMsgUniqueIds.Any())
                    {
                        foreach (UniqueId uid in childFolderMsgUniqueIds)
                        {
                            MimeMessage message = childFolder.GetMessage(uid);
                            userStorySubject = message.Subject;
                            userStoryDescription =
                                         "From: " + message.From +
                                "<br>" + "Date Sent: " + message.Date + "</br>" +
                                "<br>" + "Subject: " + userStorySubject + "</br>" +
                                "<br>" + "Request: " + message.GetTextBody(TextFormat.Plain) + "<br>";

                            if (userStorySubject.IsEmpty())
                            {
                                userStorySubject = "<No Subject User Story>";
                            }

                            toCreate[RALLY.Name] = userStorySubject;
                            toCreate[RALLY.Description] = userStoryDescription;
                            createUserStory = _rallyRestApi.Create(RALLY.HierarchicalRequirement, toCreate);
                            userStoryReference = createUserStory.Reference;

                            #region Download Attachments

                            foreach (MimeEntity attachment in message.BodyParts)
                            {
                                string attachmentFile = attachment.ContentDisposition?.FileName ??
                                                        attachment.ContentType.Name;
                                string attachmentFilePath = Concat(STORAGE.MimeKitAttachmentsDirectoryWork,
                                    Path.GetFileName(attachmentFile));

                                if (!IsNullOrWhiteSpace(attachmentFile))
                                {
                                    if (File.Exists(attachmentFilePath))
                                    {
                                        string extension = Path.GetExtension(attachmentFilePath);
                                        string fileNameWithoutExtension =
                                            Path.GetFileNameWithoutExtension(attachmentFilePath);
                                        attachmentFile = Format(fileNameWithoutExtension + "-{0}" + "{1}", ++anotherOne,
                                            extension);
                                        attachmentFilePath =
                                            Path.Combine(STORAGE.MimeKitAttachmentsDirectoryWork, attachmentFile);
                                    }

                                    using (var attachmentStream = File.Create(attachmentFilePath))
                                    {
                                        MimeKit.MimePart part = (MimeKit.MimePart)attachment;
                                        part.ContentObject.DecodeTo(attachmentStream);
                                    }

                                    Console.WriteLine("Downloaded: " + attachmentFile);
                                }
                            }

                            #endregion

                            #region Process Attachments

                            allAttachments = Directory.GetFiles(STORAGE.MimeKitAttachmentsDirectoryWork);
                            foreach (string file in allAttachments)
                            {
                                base64String = FileToBase64(file);
                                attachmentFileName = Path.GetFileName(file);
                                fileName = Empty;

                                if (!(attachmentsDictionary.TryGetValue(base64String, out fileName)))
                                {
                                    Console.WriteLine("Added to Dictionary: " + file);
                                    attachmentsDictionary.Add(base64String, attachmentFileName);
                                }

                                File.Delete(file);
                            }

                            #endregion

                            #region Upload to Rally
                            foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
                            {
                                try
                                {
                                    //create attachment content
                                    attachmentContent[RALLY.Content] = attachmentPair.Key;
                                    attachmentContentCreateResult = _rallyRestApi.Create(
                                        RALLY.AttachmentContent,
                                        attachmentContent);
                                    attachmentReference = attachmentContentCreateResult.Reference;

                                    //create attachment contianer
                                    attachmentContainer[RALLY.Artifact] = userStoryReference;
                                    attachmentContainer[RALLY.Content] = attachmentReference;
                                    attachmentContainer[RALLY.Name] = attachmentPair.Value;
                                    attachmentContainer[RALLY.Description] = RALLY.EmailAttachment;
                                    attachmentContainer[RALLY.ContentType] = "file/";

                                    //Create & associate the attachment
                                    attachmentContainerCreateResult = _rallyRestApi.Create(RALLY.Attachment,
                                        attachmentContainer);
                                    Console.WriteLine("Uploaded to Rally: " + attachmentPair.Value);
                                }
                                catch (WebException e)
                                {
                                    Console.WriteLine("Attachment: " + e.Message);
                                }
                            }
                            attachmentsDictionary.Clear();

                            #endregion

                            #region See and Move

                            childFolder.SetFlags(uid, MessageFlags.Seen, true);
                            childFolder.MoveTo(uid, processedFolder);

                            #endregion

                            #region Slack

                            if (userStoryReference != null)
                            {
                                _objectId = Ref.GetOidFromRef(userStoryReference);
                                _userStoryUrl = string.Concat(SLACK.UserStoryUrlFormat, _objectId);
                                _slackAttachmentString = string.Format("User Story: <{0} | {1} >", _userStoryUrl, message.Subject);

                                SlackMessage slackMessage = new SlackMessage
                                {
                                    //Channel is set according to the source of the email message folder
                                    Channel = slackChannel,
                                    Text = SLACK.SlackNotificationBanner,
                                    IconEmoji = Emoji.SmallRedTriangle,
                                    Username = SLACK.SlackUser
                                };

                                SlackAttachment slackAttachment = new SlackAttachment
                                {
                                    Fallback = _slackAttachmentString,
                                    Text = _slackAttachmentString,
                                    Color = SLACK.HexColor
                                };

                                slackMessage.Attachments = new List<SlackAttachment> { slackAttachment };
                                _slackClient.Post(slackMessage);
                            }
                            else
                            {
                                throw new NullReferenceException();
                            }

                            #endregion

                            #region Email
                            using (SmtpClient smtpClient = new SmtpClient())
                            {
                                if (!smtpClient.IsAuthenticated)
                                {
                                    AuthenticateWithGoogleSmtp(smtpClient);
                                }

                                //iterate throught the email addresses, to send the emails
                                List<MailboxAddress> emailNoticationList = new List<MailboxAddress>();
                                emailNoticationList.Add(new MailboxAddress("maddirsh@mail.uc.edu"));

                                foreach (var mailboxAddress in emailNoticationList)
                                {
                                    MimeMessage emailNotificationMessage = new MimeMessage();
                                    emailNotificationMessage.From.Add(new MailboxAddress("Rally Integration", EMAIL.GoogleUsername));
                                    emailNotificationMessage.To.Add(mailboxAddress);
                                    emailNotificationMessage.Subject = "Rally Notification: " + userStorySubject;
                                    emailNotificationMessage.Body = new TextPart("plain")
                                    {
                                        Text = "User Story: " + _userStoryUrl
                                    };

                                    smtpClient.Send(emailNotificationMessage);
                                }

                                //disconnect here...
                                //this will make the program connect and disconnect in a loop
                            }
                            #endregion

                            Console.WriteLine(message.Subject + " Created");
                        }
                    }
                    else
                    {
                        Console.WriteLine(childFolder + "-No Unread Messages");
                    }
                }
                Console.WriteLine("Done");
                client.Disconnect(true);
            }
        }

        private static void AuthenticateWithGoogleSmtp(SmtpClient client)
        {
            client.ServerCertificateValidationCallback = (s, c, h, e) => true;
            client.Connect(EMAIL.GoogleSmtpHost, EMAIL.SmtpPort, true);
            client.AuthenticationMechanisms.Remove(EMAIL.GoogleOAuth);
            client.Authenticate(EMAIL.GoogleUsername, EMAIL.GenericPassword);
        }

        private static void AuthenticateWithGoogleImap(ImapClient client)
        {
            #region authentication

            client.ServerCertificateValidationCallback = (s, c, ch, e) => true;
            client.Connect(EMAIL.GoogleImapHost, EMAIL.ImapPort, SecureSocketOptions.SslOnConnect);
            client.AuthenticationMechanisms.Remove(EMAIL.GoogleOAuth);
            client.Authenticate(EMAIL.GoogleUsername, EMAIL.GenericPassword);

            #endregion
        }

        #endregion

        //set project string according to the folder
        //get the List<string> 
        //iterate over that List to get the queried emails
        //send emails to users using smtp client
        #region GetProjectAdmins
        public void GetProjectAdmins(string workspaceRef)
        {
            Request projectAdminRequest = new Request("ProjectPermission");
            projectAdminRequest.Workspace = workspaceRef;
            projectAdminRequest.Fetch = new List<string>() { "User", "EmailAddress" };
            projectAdminRequest.Query = Query.And(
                new Query("Project", Query.Operator.Equals, RALLYQUERY.ProcessedEnrollmentStudentServices),
                new Query("Role", Query.Operator.Equals, "Project Admin"));

            QueryResult pAdminResult = _rallyRestApi.Query(projectAdminRequest);

            Console.WriteLine(pAdminResult.Results.Any()); 
            if (pAdminResult.Results.Any())
            {
                foreach (var admin in pAdminResult.Results)
                {
                    Console.WriteLine(admin["EmailAddress"]);
                }
            }
        }
        #endregion
    }
}

