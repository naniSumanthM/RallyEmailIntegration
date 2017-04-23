using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using ServiceStack;

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
    using ActiveUp.Net.Mail;
    using System.Drawing;
    #endregion
    class RallyOperation
    {
        RallyRestApi _rallyRestApi;
        Imap4Client _imap;
        public const string ServerName = RallyConstant.ServerId;

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
                _rallyRestApi.Authenticate(this.UserName, this.Password, ServerName, null, RallyConstant.AllowSso);
            }
        }

        private void EnsureOutlookIsAuthenticated()
        {
            _imap = new Imap4Client();
            _imap.ConnectSsl(EmailConstant.OutlookHost, EmailConstant.ImapPort);
            _imap.Login(EmailConstant.OutlookUsername, EmailConstant.GenericPassword);
        }

        #region: Query Workspaces

        /// <summary>
        /// Returns all the existing workspaces in Rally
        /// </summary>
        public void getWorkspaces()
        {
            //Authenticate
            this.EnsureRallyIsAuthenticated();

            //instantiate a DynamicJsonObject obj
            DynamicJsonObject djo = _rallyRestApi.GetSubscription(RallyQueryConstant.Workspaces);
            Request workspaceRequest = new Request(djo[RallyQueryConstant.Workspaces]);

            try
            {
                //query for the workspaces
                QueryResult returnWorkspaces = _rallyRestApi.Query(workspaceRequest);

                //iterate through the list and return the list of workspaces
                foreach (var value in returnWorkspaces.Results)
                {
                    var workspaceReference = value[RallyQueryConstant.Reference];
                    var workspaceName = value[RallyConstant.Name];
                    Console.WriteLine(RallyQueryConstant.WorkspaceMessage + workspaceName);
                }
            }
            catch (WebException)
            {
                Console.WriteLine(RallyQueryConstant.WebExceptionMessage);
            }
        }
        #endregion

        #region: Query Scrum Teams
        /// <summary>
        /// Retreives all the scrum teams within the Rally Enviornment
        /// </summary>
        public void getScrumTeams()
        {
            this.EnsureRallyIsAuthenticated();

            //DynamicJSonObject instantion
            DynamicJsonObject dObj = _rallyRestApi.GetSubscription(RallyQueryConstant.Workspaces);

            try
            {
                Request workspaceRequest = new Request(dObj[RallyQueryConstant.Workspaces]);
                QueryResult workSpaceQuery = _rallyRestApi.Query(workspaceRequest);

                foreach (var workspace in workSpaceQuery.Results)
                {
                    Request projectRequest = new Request(workspace[RallyQueryConstant.Projects]);
                    projectRequest.Fetch = new List<String>() { RallyConstant.Name };

                    //Query for the projects
                    QueryResult projectQuery = _rallyRestApi.Query(projectRequest);
                    foreach (var project in projectQuery.Results)
                    {
                        Console.WriteLine(project[RallyConstant.Name]);
                    }
                }
            }
            catch (WebException)
            {
                Console.WriteLine(RallyQueryConstant.WebExceptionMessage);
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

        public void getUserStories(string workspaceRef, string projectRef)
        {
            //Authenticate
            this.EnsureRallyIsAuthenticated();

            //setup the userStoryRequest
            Request userStoryRequest = new Request(RallyConstant.HierarchicalRequirement);
            userStoryRequest.Workspace = workspaceRef;
            userStoryRequest.Project = projectRef;
            userStoryRequest.ProjectScopeUp = RallyConstant.ProjectScopeUp;
            userStoryRequest.ProjectScopeDown = RallyConstant.ProjectScopeDown;

            //fetch data from the story request
            userStoryRequest.Fetch = new List<string>()
            {
                RallyConstant.FormattedId, RallyConstant.Name, RallyConstant.Owner
            };

            try
            {
                //query the items in the list
                userStoryRequest.Query = new Query(RallyQueryConstant.LastUpdatDate, Query.Operator.GreaterThan, RallyQueryConstant.DateGreaterThan);
                QueryResult userStoryResult = _rallyRestApi.Query(userStoryRequest);

                //iterate through the userStory Collection
                foreach (var userStory in userStoryResult.Results)
                {
                    var userStoryOwner = userStory[RallyConstant.Owner];
                    if (userStoryOwner != null)
                    {
                        var USOwner = userStoryOwner[RallyQueryConstant.ReferenceObject];
                        Console.WriteLine(userStory[RallyConstant.FormattedId] + ":" + userStory[RallyConstant.Name] + Environment.NewLine + RallyQueryConstant.UserStoryMessage + USOwner + Environment.NewLine);
                    }
                }
            }
            catch (WebException)
            {
                Console.WriteLine(RallyQueryConstant.WebExceptionMessage);
            }
        }
        #endregion

        #region: Query User Stories and Tasks
        /// <summary>
        /// When provided with a workspace and a project, the method will return all the user stories along with the tasks and their details
        /// </summary>
        /// <param name="workspaceRef"></param>
        /// <param name="projectRef"></param>

        public void getUSTA(string workspaceRef, string projectRef)
        {
            //Authenticate
            this.EnsureRallyIsAuthenticated();

            //stage the request (not using the getters and setters from the Rally Enviornment class
            Request userStoryRequest = new Request(RallyConstant.HierarchicalRequirement);
            userStoryRequest.Workspace = workspaceRef;
            userStoryRequest.Project = projectRef;
            userStoryRequest.ProjectScopeUp = RallyConstant.ProjectScopeUp;
            userStoryRequest.ProjectScopeDown = RallyConstant.ProjectScopeDown;

            //fetch US data in the form of a list
            userStoryRequest.Fetch = new List<string>()
        {
            RallyConstant.FormattedId, RallyConstant.Name, RallyConstant.TasksUpperCase, RallyConstant.Estimate, RallyConstant.State, RallyConstant.Owner
        };

            //Userstory Query
            userStoryRequest.Query = (new Query(RallyQueryConstant.LastUpdatDate, Query.Operator.GreaterThan, RallyQueryConstant.DateGreaterThan));

            try
            {
                //query for the items in the list
                QueryResult userStoryResult = _rallyRestApi.Query(userStoryRequest);

                //iterate through the query results
                foreach (var userStory in userStoryResult.Results)
                {
                    var userStoryOwner = userStory[RallyConstant.Owner];
                    if (userStoryOwner != null) //return only US who have an assigned owner
                    {
                        var USOwner = userStoryOwner[RallyQueryConstant.ReferenceObject];
                        Console.WriteLine(userStory[RallyConstant.FormattedId] + ":" + userStory[RallyConstant.Name]);
                        Console.WriteLine(RallyQueryConstant.UserStoryMessage + USOwner);
                    }

                    //Task Request
                    Request taskRequest = new Request(userStory[RallyConstant.TasksUpperCase]);
                    QueryResult taskResult = _rallyRestApi.Query(taskRequest);
                    if (taskResult.TotalResultCount > 0)
                    {
                        foreach (var task in taskResult.Results)
                        {
                            var taskName = task[RallyConstant.Name];
                            var owner = task[RallyConstant.Owner];
                            var taskState = task[RallyConstant.State];
                            var taskEstimate = task[RallyConstant.Estimate];
                            //var taskDescription = task[RallyField.description];

                            if (owner != null)
                            {
                                var ownerName = owner[RallyQueryConstant.ReferenceObject];
                                Console.WriteLine(RallyQueryConstant.TaskName + taskName + Environment.NewLine + RallyQueryConstant.TaskOwner + ownerName + Environment.NewLine + RallyQueryConstant.TaskState + taskState + Environment.NewLine + RallyQueryConstant.TaskEstimate + taskEstimate);
                                //Console.WriteLine(QueryField.taskDescription + taskDescription);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(RallyQueryConstant.TaskMessage);
                    }
                }
            }
            catch (WebException)
            {
                Console.WriteLine(RallyQueryConstant.WebExceptionMessage);
            }

        }

        #endregion

        #region: Query Iterations
        public void getIterations(string workspace, string project)
        {

            this.EnsureRallyIsAuthenticated();

            Request iterationRequest = new Request(RallyConstant.Iteration);
            iterationRequest.Workspace = workspace;
            iterationRequest.Project = project;
            iterationRequest.ProjectScopeUp = RallyConstant.ProjectScopeUp;
            iterationRequest.ProjectScopeDown = RallyConstant.ProjectScopeDown;

            try
            {
                iterationRequest.Fetch = new List<string>()
                {
                 RallyConstant.Name
                };

                iterationRequest.Query = new Query(RallyConstant.Project, Query.Operator.Equals, RallyQueryConstant.ScrumTeamSampleProject);
                QueryResult queryResult = _rallyRestApi.Query(iterationRequest);
                foreach (var iteration in queryResult.Results)
                {
                    Console.WriteLine(iteration[RallyConstant.Name]);
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
            toCreate[RallyConstant.WorkSpace] = workspace;
            toCreate[RallyConstant.Project] = project;
            toCreate[RallyConstant.Name] = userstory;
            toCreate[RallyConstant.Description] = userstoryDescription;
            toCreate[RallyConstant.Owner] = userstoryOwner;
            toCreate[RallyConstant.PlanEstimate] = "1";
            toCreate[RallyConstant.PortfolioItem] = RallyQueryConstant.FeatureShareProject;
            //toCreate[RallyConstant.Iteration] = usIteration;

            try
            {
                CreateResult createUserStory = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);
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
            toCreate[RallyConstant.Name] = taskName;
            toCreate[RallyConstant.Description] = taskDescription;
            toCreate[RallyConstant.Owner] = taskOwner;
            toCreate[RallyConstant.Estimate] = taskEstimate;
            toCreate[RallyConstant.WorkProduct] = userStoryReference;

            //create a task and attach it to a userStory
            try
            {
                Console.WriteLine("<<Creating TA>>");
                CreateResult createTask = _rallyRestApi.Create(RallyConstant.TasksLowerCase, toCreate);
                Console.WriteLine("<<Created TA>>");
            }
            catch (WebException)
            {

                Console.WriteLine(RallyQueryConstant.WebExceptionMessage);
            }

        }

        #endregion

        #region: RallyIntegration User Stories through unread email
        //testing a list of userstories
        public void SyncUserStories(string usWorkspace, string usProject)
        {
            //Authenticate with Rally
            this.EnsureRallyIsAuthenticated();

            //List to add the unreadMessages
            List<Message> unreadMessageList = new List<Message>();
            unreadMessageList.Capacity = 25;

            //Set up the US
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyConstant.WorkSpace] = usWorkspace;
            toCreate[RallyConstant.Project] = usProject;

            Console.WriteLine("Starting...");
            try
            {
                //Authenticate with Imap
                Imap4Client imap = new Imap4Client();
                imap.ConnectSsl(EmailConstant.OutlookHost, EmailConstant.ImapPort);
                imap.Login(EmailConstant.OutlookUsername, EmailConstant.GenericPassword);

                //setup Imap enviornment
                Mailbox inbox = imap.SelectMailbox(EmailConstant.InboxFolder);
                int[] unread = inbox.Search(EmailConstant.UnseenMessages);
                Console.WriteLine("Unread Messages: " + unread.Length);

                if (unread.Length > 0)
                {
                    //fetch all the messages and populate the unreadMessageList with items
                    for (int i = 0; i < unread.Length; i++)
                    {
                        Message msg = inbox.Fetch.MessageObject(unread[i]);
                        unreadMessageList.Add(msg);
                    }

                    //Create a Rally user story, with a description for each unread email message
                    for (int i = 0; i < unreadMessageList.Count; i++)
                    {
                        toCreate[RallyConstant.Name] = (unreadMessageList[i].Subject);
                        toCreate[RallyConstant.Description] = (unreadMessageList[i].BodyText.Text);
                        CreateResult cr = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);
                    }

                    //Move Fetched Messages into the processed folder
                    //Maybe mark them as unread, if a developer wants to still examine an email for further clarity
                    foreach (var item in unread)
                    {
                        inbox.MoveMessage(item, EmailConstant.ProcessedFolder);
                    }
                }
                else
                {
                    Console.WriteLine("Unread Email Messages Not Found");
                }
                Console.WriteLine("Finished...");
            }
            catch (WebException)
            {
                Console.WriteLine(RallyQueryConstant.WebExceptionMessage);
            }
        }
        #endregion

        #region: SyncUserStoriesAndLeaveMessageAsUnread
        ///<summary>
        ///After each email item is moved to the processed folder, it has to be marked as unread for future reference
        ///</summary>   

        public void SyncUserStoriesAndLeaveMessageAsUnread(string usWorkspace, string usProject)
        {
            //Authenticate with Rally
            this.EnsureRallyIsAuthenticated();

            //List to add the unreadMessages
            List<Message> unreadMessageList = new List<Message>();
            unreadMessageList.Capacity = 25;

            //Set up the US
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyConstant.WorkSpace] = usWorkspace;
            toCreate[RallyConstant.Project] = usProject;

            Console.WriteLine("Start");
            try
            {
                //Authenticate with Imap
                Imap4Client imap = new Imap4Client();
                imap.ConnectSsl(EmailConstant.OutlookHost, EmailConstant.ImapPort);
                imap.Login(EmailConstant.OutlookUsername, EmailConstant.GenericPassword);

                //setup Imap enviornment
                Mailbox inbox = imap.SelectMailbox(EmailConstant.InboxFolder);
                int[] unread = inbox.Search(EmailConstant.UnseenMessages);
                Console.WriteLine("Unread Messages: " + unread.Length);
                FlagCollection markAsUnreadFlag = new FlagCollection();

                if (unread.Length > 0)
                {
                    //fetch all the messages and add to the unreadMessageList
                    for (int i = 0; i < unread.Length; i++)
                    {
                        Message msg = inbox.Fetch.MessageObject(unread[i]);
                        unreadMessageList.Add(msg);
                    }

                    //Create a Rally user story along with the description found from the email
                    for (int i = 0; i < unreadMessageList.Count; i++)
                    {
                        toCreate[RallyConstant.Name] = (unreadMessageList[i].Subject);
                        toCreate[RallyConstant.Description] = (unreadMessageList[i].BodyText.Text);
                        CreateResult cr = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);
                    }

                    //Move Fetched Messages - Here We are blindly just moving all the messages in the unseen array, assuming they are processed
                    foreach (var item in unread)
                    {
                        markAsUnreadFlag.Add("Seen");
                        inbox.RemoveFlags(item, markAsUnreadFlag); //removing the seen flag on the email object
                        inbox.MoveMessage(item, EmailConstant.ProcessedFolder);
                    }
                    //TODO: Safer to write another loop and iterate over the procesed folder, but that will crawl
                    //the entire inbox and mark the already read items as unread.
                    //Need to find a way to say "mark these newly added items as unread - (index[i], recentlyAdded);
                }
                else
                {
                    Console.WriteLine("Unread Email Not-Found!");
                }
                Console.WriteLine("End");
            }
            catch (WebException)
            {
                Console.WriteLine(RallyQueryConstant.WebExceptionMessage);
            }
        }
        #endregion

        #region: fileToBase64
        public static string fileToBase64(string attachment)
        {
            Byte[] attachmentBytes = File.ReadAllBytes(attachment);
            string base64EncodedString = Convert.ToBase64String(attachmentBytes);
            return base64EncodedString;
        }
        #endregion

        #region: add an attachment to a an existing user story reference
        ///<summary>
        ///Pushes an attachment file to an existing user story
        /// </summary>

        public void addAttachmentToUS(string usWorkspace, string usProject)
        {
            this.EnsureRallyIsAuthenticated();

            string storyReference = "https://rally1.rallydev.com/slm/webservice/v2.0/hierarchicalrequirement/70836533324";

            //Process - Attaching the image to the user story

            // Read In Image Content
            String imageFilePath = "C:\\Users\\maddirsh\\Desktop\\";
            String imageFileName = "web.png";
            String fullImageFile = imageFilePath + imageFileName;

            // Convert Image to Base64 format
            string imageBase64String = fileToBase64(fullImageFile);

            // Length calculated from Base64String converted back
            var imageNumberBytes = Convert.FromBase64String(imageBase64String).Length;

            // DynamicJSONObject for AttachmentContent
            DynamicJsonObject myAttachmentContent = new DynamicJsonObject();
            myAttachmentContent["Content"] = imageBase64String; //string 

            try
            {
                CreateResult myAttachmentContentCreateResult = _rallyRestApi.Create("AttachmentContent", myAttachmentContent);
                String myAttachmentContentRef = myAttachmentContentCreateResult.Reference;
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

        #region: create userStory with single attachment
        ///<summary>
        ///Method that creates a user story with an attachment (takes only 1 png attachment)
        /// </summary>

        public void createUsWithAttachment(string workspace, string project, string userStoryName, string userStoryDescription)
        {
            //Authentication
            this.EnsureRallyIsAuthenticated();

            //UserStory Setup
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyConstant.WorkSpace] = workspace;
            toCreate[RallyConstant.Project] = project;
            toCreate[RallyConstant.Name] = userStoryName;
            toCreate[RallyConstant.Description] = userStoryDescription;

            //get the image reference - assume that this is where the image lives in respect to the path after being pulled from outlook
            String imageFilePath = "C:\\Users\\maddirsh\\Desktop\\";
            String imageFileName = "webException.png";
            String fullImageFile = imageFilePath + imageFileName;
            Image myImage = Image.FromFile(fullImageFile);

            // Convert Image to Base64 format
            string imageBase64String = fileToBase64(fullImageFile);

            // Length calculated from Base64String converted back
            var imageNumberBytes = Convert.FromBase64String(imageBase64String).Length;

            // DynamicJSONObject for AttachmentContent
            DynamicJsonObject myAttachmentContent = new DynamicJsonObject();
            myAttachmentContent[RallyConstant.Content] = imageBase64String;

            try
            {
                //create user story
                CreateResult createUserStory = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);

                //create attachment
                CreateResult myAttachmentContentCreateResult = _rallyRestApi.Create(RallyConstant.AttachmentContent, myAttachmentContent);
                String myAttachmentContentRef = myAttachmentContentCreateResult.Reference;

                // DynamicJSONObject for Attachment Container
                DynamicJsonObject myAttachment = new DynamicJsonObject();
                myAttachment[RallyConstant.Artifact] = createUserStory.Reference;
                myAttachment[RallyConstant.Content] = myAttachmentContentRef;
                myAttachment[RallyConstant.Name] = "fileName.png"; //method to get the fileName from the attached documents
                myAttachment[RallyConstant.Description] = "Email Attachment";
                myAttachment[RallyConstant.ContentType] = "image/png"; //Method to identify the fileType.java
                myAttachment[RallyConstant.Size] = imageNumberBytes;

                //create & associate the attachment
                CreateResult myAttachmentCreateResult = _rallyRestApi.Create(RallyConstant.Attachment, myAttachment);
                Console.WriteLine("Created User Story: " + createUserStory.Reference);
            }
            catch (WebException e)
            {
                Console.WriteLine(e.Message);
            }
        }

        #endregion

        #region : create userStory with a collection of diverse attachment types
        /// <summary>
        /// Method to upload a diverse set of attachments to a user story
        /// </summary>
        /// <param name="workspace"></param>
        /// <param name="project"></param>
        /// <param name="userstoryName"></param>

        public void addAttachments(string workspace, string project, string userstoryName)
        {
            //Dictionary Object to hold each attachments base64EncodedString and its fileName
            Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();

            //Objects to support attachment specifics
            DynamicJsonObject attachmentContent = new DynamicJsonObject();
            DynamicJsonObject attachmentContainer = new DynamicJsonObject();

            //Objects that helps create the a) user story, b) attachment content, c) attachment container
            CreateResult createUserStory;
            CreateResult attachmentContentCreateResult;
            CreateResult attachmentContainerCreateResult;

            //base 64 conversion variables
            string[] attachmentPaths = Directory.GetFiles("C:\\Users\\maddirsh\\Desktop\\diverseAttachments");
            string base64EncodedString;
            string attachmentFileName;
            string attachmentContentReference = "";

            //Rally Authentication
            this.EnsureRallyIsAuthenticated();

            //User story creation and set up
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyConstant.WorkSpace] = workspace;
            toCreate[RallyConstant.Project] = project;
            toCreate[RallyConstant.Name] = userstoryName;
            createUserStory = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);

            //iterate over each filePath and a) convert to base 64, b) get the fileName.extension, c)Add to the dictionary object
            foreach (string attachment in attachmentPaths)
            {
                //Base 64 conversion process
                base64EncodedString = fileToBase64(attachment);
                attachmentFileName = Path.GetFileName(attachment);

                //Populate the Dictionary
                attachmentsDictionary.Add(base64EncodedString, attachmentFileName);
            }

            //iterate over the populated dictionary and upload each attachment to the respective user story
            foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
            {
                try
                {
                    //create attachment content
                    attachmentContent[RallyConstant.Content] = attachmentPair.Key;
                    attachmentContentCreateResult = _rallyRestApi.Create(RallyConstant.AttachmentContent, attachmentContent);
                    attachmentContentReference = attachmentContentCreateResult.Reference;

                    //create attachment contianer
                    attachmentContainer[RallyConstant.Artifact] = createUserStory.Reference;
                    attachmentContainer[RallyConstant.Content] = attachmentContentReference;
                    attachmentContainer[RallyConstant.Name] = attachmentPair.Value;
                    attachmentContainer[RallyConstant.Description] = RallyConstant.EmailAttachment;
                    attachmentContainer[RallyConstant.ContentType] = "file/";
                    //attachmentContainer[RallyField.size] = Omitted

                    //Create & associate the attachment
                    attachmentContainerCreateResult = _rallyRestApi.Create(RallyConstant.Attachment, attachmentContainer);
                    Console.WriteLine("Created User Story: " + createUserStory.Reference);
                }
                catch (WebException e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        #endregion

        #region: avoidDuplicateAttachments
        ///<summary>
        ///Link attachments to a user story, ELIMINATING THE IDEA OF DUPLICATE FILES.
        ///When someone does attach the same file twice, the download will not accept fileNames with the same fileName
        ///Dictionary <fileName.extension...fileBase64String>
        ///Problem with this method is that k.txt and k(1).txt will have the same base64Strings, so we really want to filter duplicate attachments before pushing to rally
        /// </summary>

        public void addAttachmentsEliminateDuplicates(string workspace, string project, string userstoryName)
        {
            //Dictionary Object to hold each attachments base64EncodedString and its fileName
            Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();

            //Objects to support attachment specifics
            DynamicJsonObject attachmentContent = new DynamicJsonObject();
            DynamicJsonObject attachmentContainer = new DynamicJsonObject();

            //Objects that helps create the a) user story, b) attachment content, c) attachment container
            CreateResult createUserStory;
            CreateResult attachmentContentCreateResult;
            CreateResult attachmentContainerCreateResult;

            //base 64 conversion variables
            string attachmentFilePath = "C:\\Users\\maddirsh\\Desktop\\diverseAttachments";
            string[] attachmentPaths = Directory.GetFiles(attachmentFilePath);
            string base64EncodedString;
            string attachmentFileName;
            string attachmentContentReference = "";

            //Rally Authentication
            this.EnsureRallyIsAuthenticated();

            //User story creation and set up
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyConstant.WorkSpace] = workspace;
            toCreate[RallyConstant.Project] = project;
            toCreate[RallyConstant.Name] = userstoryName;
            createUserStory = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);

            //iterate over each filePath and a) convert to base 64, b) get the fileName.extension, c)Add to the dictionary object
            foreach (string attachment in attachmentPaths)
            {
                //Base 64 conversion process
                attachmentFileName = Path.GetFileName(attachment);
                base64EncodedString = fileToBase64(attachment);

                //Populate the Dictionary
                if (attachmentsDictionary.ContainsKey(attachmentFileName))
                {
                    Console.WriteLine("exists, so not adding");
                    //trying to get a value for a key that does not exist
                }
                else
                {
                    Console.WriteLine("dOES NOT exist");
                    attachmentsDictionary.Add(attachmentFileName, base64EncodedString);
                    Console.WriteLine("ADDING");
                }
            }

            //iterate over the populated dictionary and upload each attachment to the respective user story
            foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
            {
                try
                {
                    //create attachment content
                    attachmentContent[RallyConstant.Content] = attachmentPair.Value;
                    attachmentContentCreateResult = _rallyRestApi.Create(RallyConstant.AttachmentContent, attachmentContent);
                    attachmentContentReference = attachmentContentCreateResult.Reference;

                    //create attachment contianer
                    attachmentContainer[RallyConstant.Artifact] = createUserStory.Reference;
                    attachmentContainer[RallyConstant.Content] = attachmentContentReference;
                    attachmentContainer[RallyConstant.Name] = attachmentPair.Key;
                    attachmentContainer[RallyConstant.Description] = RallyConstant.EmailAttachment;
                    attachmentContainer[RallyConstant.ContentType] = "file/";

                    //Create & associate the attachment
                    attachmentContainerCreateResult = _rallyRestApi.Create(RallyConstant.Attachment, attachmentContainer);
                    Console.WriteLine("Created User Story: " + createUserStory.Reference);
                }
                catch (WebException e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        #endregion

        #region: addAttachmentsEliminateDuplicatesWithSimilarBase64Strings
        /// <summary>
        /// Use tryGetValue to identify duplicate file content with different file names
        /// (string base64EncodedString, string fileName)
        /// </summary>

        public void addAttachmentsEliminateDuplicatesWithSimilarBase64Strings(string workspace, string project, string userstoryName)
        {
            //Dictionary Object to hold each attachments base64EncodedString and its fileName
            Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();

            //Objects to support attachment specifics
            DynamicJsonObject attachmentContent = new DynamicJsonObject();
            DynamicJsonObject attachmentContainer = new DynamicJsonObject();

            //Objects that helps create the a) user story, b) attachment content, c) attachment container
            CreateResult createUserStory;
            CreateResult attachmentContentCreateResult;
            CreateResult attachmentContainerCreateResult;

            //base 64 conversion variables
            string attachmentFilePath = "C:\\Users\\maddirsh\\Desktop\\diverseAttachments";
            string[] attachmentPaths = Directory.GetFiles(attachmentFilePath);
            string base64EncodedString;
            string attachmentFileName;
            string attachmentContentReference = "";
            int attachmentCount = 0;

            //Rally Authentication
            this.EnsureRallyIsAuthenticated();

            //User story creation and set up
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyConstant.WorkSpace] = workspace;
            toCreate[RallyConstant.Project] = project;
            toCreate[RallyConstant.Name] = userstoryName;
            createUserStory = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);

            foreach (string attachment in attachmentPaths)
            {
                //Base 64 conversion process
                attachmentFileName = Path.GetFileName(attachment);
                base64EncodedString = fileToBase64(attachment);
                var filename = string.Empty;

                if (!(attachmentsDictionary.TryGetValue(base64EncodedString, out filename)))
                {
                    //base64EncodedString does not exist
                    //Populate the dictionary
                    attachmentsDictionary.Add(base64EncodedString, attachmentFileName);
                }
                else
                {
                    //Does exist so do not populate
                    Console.WriteLine("Duplicate file exists for: " + attachmentFileName);
                }
            }

            //iterate over the populated dictionary and upload each attachment to the user story created above
            foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
            {
                try
                {
                    //create attachment content
                    attachmentContent[RallyConstant.Content] = attachmentPair.Key;
                    attachmentContentCreateResult = _rallyRestApi.Create(RallyConstant.AttachmentContent, attachmentContent);
                    attachmentContentReference = attachmentContentCreateResult.Reference;

                    //create attachment contianer
                    attachmentContainer[RallyConstant.Artifact] = createUserStory.Reference;
                    attachmentContainer[RallyConstant.Content] = attachmentContentReference;
                    attachmentContainer[RallyConstant.Name] = attachmentPair.Value;
                    attachmentContainer[RallyConstant.Description] = RallyConstant.EmailAttachment;
                    attachmentContainer[RallyConstant.ContentType] = "file/";

                    //Create & associate the attachment
                    attachmentContainerCreateResult = _rallyRestApi.Create(RallyConstant.Attachment, attachmentContainer);
                    attachmentCount++;
                }
                catch (WebException e)
                {
                    Console.WriteLine(e.Message);
                }
            }

            Console.WriteLine(attachmentCount + " Attachments Pushed...");
        }

        #endregion

        #region :embeddedImages
        /// <summary>
        /// Method to pull images that could have been copied & pasted, instead of attaching
        /// </summary>
        public void downlodInlineAttachments(string workspace, string project)
        {
            //Mail Variables
            List<Message> unreadMsgCollection = new List<Message>();
            Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();
            unreadMsgCollection.Capacity = 25;
            string[] inlineAttachmentsPath;

            //Rally Variables
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyConstant.WorkSpace] = workspace;
            toCreate[RallyConstant.Project] = project;
            DynamicJsonObject attachmentContent = new DynamicJsonObject();
            DynamicJsonObject attachmentContainer = new DynamicJsonObject();
            CreateResult createUserStory;
            CreateResult attachmentContentCreateResult;
            CreateResult attachmentContainerCreateResult;
            string userStoryReference;

            //Authentication
            this.EnsureOutlookIsAuthenticated();
            this.EnsureRallyIsAuthenticated();

            var inbox = _imap.SelectMailbox(EmailConstant.InboxFolder);
            var unread = inbox.Search(EmailConstant.UnseenMessages);
            Console.WriteLine("Unread Messages: " + unread.Length);

            if (unread.Length > 0)
            {
                //Pupulate the unread Message List
                for (int i = 0; i < unread.Length; i++)
                {
                    Message msg = inbox.Fetch.MessageObject(unread[i]);
                    unreadMsgCollection.Add(msg);
                }

                for (int i = 0; i < unreadMsgCollection.Count; i++)
                {
                    toCreate[RallyConstant.Name] = (unreadMsgCollection[i].Subject);
                    toCreate[RallyConstant.Description] = (unreadMsgCollection[i].BodyText.Text);
                    createUserStory = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);

                    foreach (MimePart embedded in unreadMsgCollection[i].EmbeddedObjects)
                    {
                        var fileName = embedded.ContentName;
                        var binary = embedded.BinaryContent;
                        File.WriteAllBytes(StorageConstant.InlineImageDirectory + fileName, binary); //downloads one file from the email

                        //} //the images can all be downloaded once, but if they clash with the fileNames that are attached, they will fail
                        //mechanism for identifying duplicate file names with unique base64 string

                        //Which is always expected to be a .png extension
                        Console.WriteLine("Downloaded: " + fileName);

                        //only 1 file from the email is downloaded here at once, so we continue this proecedure for the number of emails there are in the mailbox
                        inlineAttachmentsPath = Directory.GetFiles(StorageConstant.InlineImageDirectory);

                        foreach (var file in inlineAttachmentsPath)
                        {
                            //convert to base 64
                            string base64String = fileToBase64(file);
                            string attachmentFileName = Path.GetFileName(file);
                            var emptyFileString = string.Empty;

                            Console.WriteLine("Adding to Dictionary: " + attachmentFileName);

                            if (!(attachmentsDictionary.TryGetValue(base64String, out fileName)))
                            {
                                attachmentsDictionary.Add(base64String, attachmentFileName);
                            }

                            //once the dictionary is populated, clear the file for the next email object iteration
                            File.Delete(file);
                        }

                        //now that the dictionary is populated for each inline image, we push to Rally

                        foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
                        {
                            //create attachment content
                            attachmentContent[RallyConstant.Content] = attachmentPair.Key;
                            attachmentContentCreateResult = _rallyRestApi.Create(RallyConstant.AttachmentContent, attachmentContent);
                            userStoryReference = attachmentContentCreateResult.Reference;

                            //create attachment contianer
                            attachmentContainer[RallyConstant.Artifact] = createUserStory.Reference;
                            attachmentContainer[RallyConstant.Content] = userStoryReference;
                            attachmentContainer[RallyConstant.Name] = attachmentPair.Value;
                            attachmentContainer[RallyConstant.Description] = RallyConstant.EmailAttachment;
                            attachmentContainer[RallyConstant.ContentType] = "file/";

                            //Create & associate the attachment
                            attachmentContainerCreateResult = _rallyRestApi.Create(RallyConstant.Attachment, attachmentContainer);
                        }
                        //clear the dictionary for each parse mime part object, not for each email
                        //unlike the other example where ALL the images are parsed from a given email object in one go for one iteration
                        attachmentsDictionary.Clear();
                    }
                }
                Console.WriteLine("Created " + unread.Length + " User stories with Inline Images");
            }
            else
            {
                Console.WriteLine("No Unread Messages");
            }
        }
        #endregion

        #region: SyncRallyUserStories
        public void Sync(string workspace, string project)
        {
            //Email variables
            List<Message> unreadMsgCollection = new List<Message>();
            unreadMsgCollection.Capacity = 25;
            Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();

            //Rally variables
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyConstant.WorkSpace] = workspace;
            toCreate[RallyConstant.Project] = project;
            DynamicJsonObject attachmentContent = new DynamicJsonObject();
            DynamicJsonObject attachmentContainer = new DynamicJsonObject();
            CreateResult createUserStory;
            CreateResult attachmentContentCreateResult;
            CreateResult attachmentContainerCreateResult;

            //base 64 conversion variables
            string base64String;
            string attachmentFileName;
            string[] attachmentPaths;
            string userStoryReference;

            try
            {
                //Authentication
                EnsureRallyIsAuthenticated();
                EnsureOutlookIsAuthenticated();

                //Setup Imap enviornment
                Mailbox inbox = _imap.SelectMailbox(EmailConstant.InboxFolder);
                int[] unread = inbox.Search(EmailConstant.UnseenMessages);
                FlagCollection markAsUnreadFlag = new FlagCollection();

                if (unread.Length > 0)
                {
                    Console.WriteLine("Syncing: " + unread.Length + " Unread Messages");
                    //fetch and populate unreadMsgCollection with unread Message Objects
                    for (int i = 0; i < unread.Length; i++)
                    {
                        Message msg = inbox.Fetch.MessageObject(unread[i]);
                        unreadMsgCollection.Add(msg);
                    }

                    //Iterate through the collection 1) Create the user story 2) Check for attachments 3) Convert attachments to base 64 4)Delete attachments once pushed 
                    for (int i = 0; i < unreadMsgCollection.Count; i++)
                    {
                        //stage the user story
                        if (unreadMsgCollection[i].Subject.Equals(""))
                        {
                            unreadMsgCollection[i].Subject = EmailConstant.NoSubject;
                        }
                        toCreate[RallyConstant.Name] = (unreadMsgCollection[i].Subject);
                        toCreate[RallyConstant.Description] = (unreadMsgCollection[i].BodyText.Text);
                        createUserStory = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);

                        //check to see if message has attachments & then store them
                        if (unreadMsgCollection[i].Attachments.Count > 0)
                        {
                            //Do all the attachments from the email object get stored here?? _No it has to complete one iteration
                            unreadMsgCollection[i].Attachments.StoreToFolder(StorageConstant.AttachmentsDirectory);
                        }

                        //reference the path where the attachments live for the [ith] message
                        attachmentPaths = Directory.GetFiles(StorageConstant.AttachmentsDirectory);

                        //Convert each attachment to base64, populate the map, and move the file
                        foreach (var file in attachmentPaths)
                        {
                            //Converting attachments to base 64
                            base64String = fileToBase64(file);
                            attachmentFileName = Path.GetFileName(file);
                            var fileName = string.Empty;

                            //populate the dictionary - eliminate adding duplicate files
                            if (!(attachmentsDictionary.TryGetValue(base64String, out fileName)))
                            {
                                attachmentsDictionary.Add(base64String, attachmentFileName);
                            }

                            Console.WriteLine("Uploading: " + file);
                            File.Delete(file);
                        }

                        //Stage the attachment
                        foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
                        {
                            try
                            {
                                //create attachment content
                                attachmentContent[RallyConstant.Content] = attachmentPair.Key;
                                attachmentContentCreateResult = _rallyRestApi.Create(RallyConstant.AttachmentContent, attachmentContent);
                                userStoryReference = attachmentContentCreateResult.Reference;

                                //create attachment contianer
                                attachmentContainer[RallyConstant.Artifact] = createUserStory.Reference;
                                attachmentContainer[RallyConstant.Content] = userStoryReference;
                                attachmentContainer[RallyConstant.Name] = attachmentPair.Value;
                                attachmentContainer[RallyConstant.Description] = RallyConstant.EmailAttachment;
                                attachmentContainer[RallyConstant.ContentType] = "file/";

                                //Create & associate the attachment
                                attachmentContainerCreateResult = _rallyRestApi.Create(RallyConstant.Attachment, attachmentContainer);
                            }
                            catch (WebException e)
                            {
                                Console.WriteLine("Attachment: " + e.Message);
                            }
                        }
                        attachmentsDictionary.Clear();
                    }

                    //Move Fetched Messages to Processed Folder, and mark them as unread()
                    foreach (var item in unread)
                    {
                        markAsUnreadFlag.Add(EmailConstant.SeenMessages);
                        inbox.RemoveFlags(item, markAsUnreadFlag);
                        inbox.MoveMessage(item, EmailConstant.ProcessedFolder);
                    }

                    Console.WriteLine("Created " + unread.Length + " User Stories");
                }
                else
                {
                    Console.WriteLine("No Unread Messages Found...");
                }
            }
            catch (Imap4Exception i)
            {
                Console.WriteLine(i.Message);
            }
            catch (IOException i)
            {
                Console.WriteLine(i.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                _imap.Disconnect();
            }

        }
        #endregion

        #region SyncUsingMimeKit
        public void SyncUsingMimeKit(string workspace, string project)
        {
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyConstant.WorkSpace] = workspace;
            toCreate[RallyConstant.Project] = project;
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
                client.Connect(EmailConstant.GoogleHost, EmailConstant.ImapPort, SecureSocketOptions.SslOnConnect);
                client.AuthenticationMechanisms.Remove(EmailConstant.GoogleOAuth);
                client.Authenticate(EmailConstant.GoogleUsername, EmailConstant.GenericPassword);

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

                    toCreate[RallyConstant.Name] = (emailSubject);
                    toCreate[RallyConstant.Description] = (emailBody);
                    createUserStory = _rallyRestApi.Create(RallyConstant.HierarchicalRequirement, toCreate);

                    foreach (MimeEntity attachment in message.BodyParts)
                    {
                        string attachmentFile = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                        string attachmentFilePath = String.Concat(StorageConstant.MimeKitAttachmentsDirectory, Path.GetFileName(attachmentFile));

                        if (!string.IsNullOrWhiteSpace(attachmentFile))
                        {
                            if (File.Exists(attachmentFilePath))
                            {
                                string extension = Path.GetExtension(attachmentFilePath);
                                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(attachmentFilePath);
                                attachmentFile = string.Format(fileNameWithoutExtension + "-{0}" + "{1}", ++anotherOne,
                                    extension);
                                attachmentFilePath = Path.Combine(StorageConstant.MimeKitAttachmentsDirectory,
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

                    allAttachments = Directory.GetFiles(StorageConstant.MimeKitAttachmentsDirectory);
                    foreach (string file in allAttachments)
                    {
                        base64String = fileToBase64(file);
                        attachmentFileName = Path.GetFileName(file);
                        fileName = string.Empty;

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
                            attachmentContent[RallyConstant.Content] = attachmentPair.Key;
                            attachmentContentCreateResult = _rallyRestApi.Create(RallyConstant.AttachmentContent,
                                attachmentContent);
                            userStoryReference = attachmentContentCreateResult.Reference;

                            //create attachment contianer
                            attachmentContainer[RallyConstant.Artifact] = createUserStory.Reference;
                            attachmentContainer[RallyConstant.Content] = userStoryReference;
                            attachmentContainer[RallyConstant.Name] = attachmentPair.Value;
                            attachmentContainer[RallyConstant.Description] = RallyConstant.EmailAttachment;
                            attachmentContainer[RallyConstant.ContentType] = "file/";

                            //Create & associate the attachment
                            attachmentContainerCreateResult = _rallyRestApi.Create(RallyConstant.Attachment,
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
    }
}
