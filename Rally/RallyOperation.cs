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
        RallyRestApi _api;
        Imap4Client imap;
        public const string ServerName = RallyField.ServerId;

        //properties
        public string UserName { get; set; }

        public string Password { get; set; }

        //constructor
        public RallyOperation(string userName, string password)
        {
            _api = new RallyRestApi();
            this.UserName = userName;
            this.Password = password;
            this.EnsureRallyIsAuthenticated();
        }

        private void EnsureRallyIsAuthenticated()
        {
            if (this._api.AuthenticationState != RallyRestApi.AuthenticationResult.Authenticated)
            {
                _api.Authenticate(this.UserName, this.Password, ServerName, null, RallyField.AllowSso);
            }
        }

        public void EnsureOutlookIsAuthenticated()
        {
            imap = new Imap4Client();
            imap.ConnectSsl(Outlook.OutlookHost, Outlook.OutlookPort);
            imap.Login(Outlook.OutlookUsername, Outlook.OutlookPassword);
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
            DynamicJsonObject djo = _api.GetSubscription(QueryField.workspaces);
            Request workspaceRequest = new Request(djo[QueryField.workspaces]);

            try
            {
                //query for the workspaces
                QueryResult returnWorkspaces = _api.Query(workspaceRequest);

                //iterate through the list and return the list of workspaces
                foreach (var value in returnWorkspaces.Results)
                {
                    var workspaceReference = value[QueryField.reference];
                    var workspaceName = value[RallyField.Name];
                    Console.WriteLine(QueryField.wsMessage + workspaceName);
                }
            }
            catch (WebException)
            {
                Console.WriteLine(QueryField.webExceptionMessage);
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
            DynamicJsonObject dObj = _api.GetSubscription(QueryField.workspaces);

            try
            {
                Request workspaceRequest = new Request(dObj[QueryField.workspaces]);
                QueryResult workSpaceQuery = _api.Query(workspaceRequest);

                foreach (var workspace in workSpaceQuery.Results)
                {
                    Request projectRequest = new Request(workspace[QueryField.projects]);
                    projectRequest.Fetch = new List<String>() { RallyField.Name };

                    //Query for the projects
                    QueryResult projectQuery = _api.Query(projectRequest);
                    foreach (var project in projectQuery.Results)
                    {
                        Console.WriteLine(project[RallyField.Name]);
                    }
                }
            }
            catch (WebException)
            {
                Console.WriteLine(QueryField.webExceptionMessage);
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
            Request userStoryRequest = new Request(RallyField.HierarchicalRequirement);
            userStoryRequest.Workspace = workspaceRef;
            userStoryRequest.Project = projectRef;
            userStoryRequest.ProjectScopeUp = RallyField.ProjectScopeUp;
            userStoryRequest.ProjectScopeDown = RallyField.ProjectScopeDown;

            //fetch data from the story request
            userStoryRequest.Fetch = new List<string>()
        {
            RallyField.FormattedId, RallyField.Name, RallyField.Owner
        };

            try
            {
                //query the items in the list
                userStoryRequest.Query = new Query(QueryField.lastUpdatDate, Query.Operator.GreaterThan, QueryField.dateGreaterThan);
                QueryResult userStoryResult = _api.Query(userStoryRequest);

                //iterate through the userStory Collection
                foreach (var userStory in userStoryResult.Results)
                {
                    var userStoryOwner = userStory[RallyField.Owner];
                    if (userStoryOwner != null)
                    {
                        var USOwner = userStoryOwner[QueryField.referenceObject];
                        Console.WriteLine(userStory[RallyField.FormattedId] + ":" + userStory[RallyField.Name] + Environment.NewLine + QueryField.usMessage + USOwner + Environment.NewLine);
                    }
                }
            }
            catch (WebException)
            {
                Console.WriteLine(QueryField.webExceptionMessage);
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
            Request userStoryRequest = new Request(RallyField.HierarchicalRequirement);
            userStoryRequest.Workspace = workspaceRef;
            userStoryRequest.Project = projectRef;
            userStoryRequest.ProjectScopeUp = RallyField.ProjectScopeUp;
            userStoryRequest.ProjectScopeDown = RallyField.ProjectScopeDown;

            //fetch US data in the form of a list
            userStoryRequest.Fetch = new List<string>()
        {
            RallyField.FormattedId, RallyField.Name, RallyField.TasksUpperCase, RallyField.Estimate, RallyField.State, RallyField.Owner
        };

            //Userstory Query
            userStoryRequest.Query = (new Query(QueryField.lastUpdatDate, Query.Operator.GreaterThan, QueryField.dateGreaterThan));

            try
            {
                //query for the items in the list
                QueryResult userStoryResult = _api.Query(userStoryRequest);

                //iterate through the query results
                foreach (var userStory in userStoryResult.Results)
                {
                    var userStoryOwner = userStory[RallyField.Owner];
                    if (userStoryOwner != null) //return only US who have an assigned owner
                    {
                        var USOwner = userStoryOwner[QueryField.referenceObject];
                        Console.WriteLine(userStory[RallyField.FormattedId] + ":" + userStory[RallyField.Name]);
                        Console.WriteLine(QueryField.usMessage + USOwner);
                    }

                    //Task Request
                    Request taskRequest = new Request(userStory[RallyField.TasksUpperCase]);
                    QueryResult taskResult = _api.Query(taskRequest);
                    if (taskResult.TotalResultCount > 0)
                    {
                        foreach (var task in taskResult.Results)
                        {
                            var taskName = task[RallyField.Name];
                            var owner = task[RallyField.Owner];
                            var taskState = task[RallyField.State];
                            var taskEstimate = task[RallyField.Estimate];
                            //var taskDescription = task[RallyField.description];

                            if (owner != null)
                            {
                                var ownerName = owner[QueryField.referenceObject];
                                Console.WriteLine(QueryField.taskName + taskName + Environment.NewLine + QueryField.taskOwner + ownerName + Environment.NewLine + QueryField.taskState + taskState + Environment.NewLine + QueryField.taskEstimate + taskEstimate);
                                //Console.WriteLine(QueryField.taskDescription + taskDescription);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(QueryField.taskMessage);
                    }
                }
            }
            catch (WebException)
            {
                Console.WriteLine(QueryField.webExceptionMessage);
            }

        }

        #endregion

        #region: Create User Story
        /// <summary>
        /// Method Creates a UserStory in Rally, according to the data passed in the parameters
        /// </summary>
        /// <param name="usName"></param>
        /// <param name="usDescription"></param>
        /// <param name="usWorkspace"></param>
        /// <param name="usProject"></param>
        /// <param name="usOwner"></param>
        /// <param name="usIteration"></param>
        /// <param name="usPlanEstimate"></param>

        public void CreateUserStory(string usName, string usDescription, string usWorkspace, string usProject, string usOwner, string usIteration, string usPlanEstimate)
        {
            //authenticate
            this.EnsureRallyIsAuthenticated();

            //DynamicJsonObject
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyField.Name] = usName;
            toCreate[RallyField.Description] = usDescription;
            toCreate[RallyField.WorkSpace] = usWorkspace;
            toCreate[RallyField.Project] = usProject;
            toCreate[RallyField.Owner] = usOwner;
            toCreate[RallyField.Iteration] = usIteration;
            toCreate[RallyField.PlanEstimate] = usPlanEstimate;

            //use try and catch to create push a US to the specific workspace within a specific project
            try
            {
                Console.WriteLine("<<Creating US>>");
                CreateResult createUserStory = _api.Create(RallyField.HierarchicalRequirement, toCreate);
                Console.WriteLine("<<Created US>>");
            }
            catch (WebException)
            {
                Console.WriteLine(QueryField.webExceptionMessage);
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
            toCreate[RallyField.Name] = taskName;
            toCreate[RallyField.Description] = taskDescription;
            toCreate[RallyField.Owner] = taskOwner;
            toCreate[RallyField.Estimate] = taskEstimate;
            toCreate[RallyField.WorkProduct] = userStoryReference;

            //create a task and attach it to a userStory
            try
            {
                Console.WriteLine("<<Creating TA>>");
                CreateResult createTask = _api.Create(RallyField.TasksLowerCase, toCreate);
                Console.WriteLine("<<Created TA>>");
            }
            catch (WebException)
            {

                Console.WriteLine(QueryField.webExceptionMessage);
            }

        }

        #endregion

        #region: Create US through List
        //testing a list of userstories
        public void CreateUserStoryFromList(string usWorkspace, string usProject)
        {
            //authenticate
            this.EnsureRallyIsAuthenticated();

            List<string> usList = new List<string>();
            usList.Add("Item 1");
            usList.Add("Item 2");
            usList.Add("Item 3");
            usList.Add("Item 4");

            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyField.WorkSpace] = usWorkspace;
            toCreate[RallyField.Project] = usProject;

            Console.WriteLine("Start");
            try
            {
                #region foreach
                //foreach (var item in usList)
                //{
                //    toCreate[RallyField.nameForWSorUSorTA] = usList[i];
                //    CreateResult cr = _api.Create("HierarchicalRequirement", toCreate);
                //}
                #endregion

                for (int i = 0; i < usList.Count; i++)
                {
                    toCreate[RallyField.Name] = usList[i];
                    CreateResult cr = _api.Create(RallyField.HierarchicalRequirement, toCreate);
                }
            }
            catch (WebException)
            {
                Console.WriteLine(QueryField.webExceptionMessage);
            }
            Console.WriteLine("End");
        }
        #endregion

        #region: Sync User Stories through unread email
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
            toCreate[RallyField.WorkSpace] = usWorkspace;
            toCreate[RallyField.Project] = usProject;

            Console.WriteLine("Starting...");
            try
            {
                //Authenticate with Imap
                Imap4Client imap = new Imap4Client();
                imap.ConnectSsl(Outlook.OutlookHost, Outlook.OutlookPort);
                imap.Login(Outlook.OutlookUsername, Outlook.OutlookPassword);

                //setup Imap enviornment
                Mailbox inbox = imap.SelectMailbox(Outlook.OutlookInboxFolder);
                int[] unread = inbox.Search(Outlook.OutlookUnseenMessages);
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
                        toCreate[RallyField.Name] = (unreadMessageList[i].Subject);
                        toCreate[RallyField.Description] = (unreadMessageList[i].BodyText.Text);
                        CreateResult cr = _api.Create(RallyField.HierarchicalRequirement, toCreate);
                    }

                    //Move Fetched Messages into the processed folder
                    //Maybe mark them as unread, if a developer wants to still examine an email for further clarity
                    foreach (var item in unread)
                    {
                        inbox.MoveMessage(item, Outlook.OutlookProcessedFolder);
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
                Console.WriteLine(QueryField.webExceptionMessage);
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
            toCreate[RallyField.WorkSpace] = usWorkspace;
            toCreate[RallyField.Project] = usProject;

            Console.WriteLine("Start");
            try
            {
                //Authenticate with Imap
                Imap4Client imap = new Imap4Client();
                imap.ConnectSsl(Outlook.OutlookHost, Outlook.OutlookPort);
                imap.Login(Outlook.OutlookUsername, Outlook.OutlookPassword);

                //setup Imap enviornment
                Mailbox inbox = imap.SelectMailbox(Outlook.OutlookInboxFolder);
                int[] unread = inbox.Search(Outlook.OutlookUnseenMessages);
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
                        toCreate[RallyField.Name] = (unreadMessageList[i].Subject);
                        toCreate[RallyField.Description] = (unreadMessageList[i].BodyText.Text);
                        CreateResult cr = _api.Create(RallyField.HierarchicalRequirement, toCreate);
                    }

                    //Move Fetched Messages - Here We are blindly just moving all the messages in the unseen array, assuming they are processed
                    foreach (var item in unread)
                    {
                        markAsUnreadFlag.Add("Seen");
                        inbox.RemoveFlags(item, markAsUnreadFlag); //removing the seen flag on the email object
                        inbox.MoveMessage(item, Outlook.OutlookProcessedFolder);
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
                Console.WriteLine(QueryField.webExceptionMessage);
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
                CreateResult myAttachmentContentCreateResult = _api.Create("AttachmentContent", myAttachmentContent);
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

                CreateResult myAttachmentCreateResult = _api.Create("Attachment", myAttachment);
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
            toCreate[RallyField.WorkSpace] = workspace;
            toCreate[RallyField.Project] = project;
            toCreate[RallyField.Name] = userStoryName;
            toCreate[RallyField.Description] = userStoryDescription;

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
            myAttachmentContent[RallyField.Content] = imageBase64String;

            try
            {
                //create user story
                CreateResult createUserStory = _api.Create(RallyField.HierarchicalRequirement, toCreate);

                //create attachment
                CreateResult myAttachmentContentCreateResult = _api.Create(RallyField.AttachmentContent, myAttachmentContent);
                String myAttachmentContentRef = myAttachmentContentCreateResult.Reference;

                // DynamicJSONObject for Attachment Container
                DynamicJsonObject myAttachment = new DynamicJsonObject();
                myAttachment[RallyField.Artifact] = createUserStory.Reference;
                myAttachment[RallyField.Content] = myAttachmentContentRef;
                myAttachment[RallyField.Name] = "fileName.png"; //method to get the fileName from the attached documents
                myAttachment[RallyField.Description] = "Email Attachment";
                myAttachment[RallyField.ContentType] = "image/png"; //Method to identify the fileType.java
                myAttachment[RallyField.Size] = imageNumberBytes;

                //create & associate the attachment
                CreateResult myAttachmentCreateResult = _api.Create(RallyField.Attachment, myAttachment);
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
            toCreate[RallyField.WorkSpace] = workspace;
            toCreate[RallyField.Project] = project;
            toCreate[RallyField.Name] = userstoryName;
            createUserStory = _api.Create(RallyField.HierarchicalRequirement, toCreate);

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
                    attachmentContent[RallyField.Content] = attachmentPair.Key;
                    attachmentContentCreateResult = _api.Create(RallyField.AttachmentContent, attachmentContent);
                    attachmentContentReference = attachmentContentCreateResult.Reference;

                    //create attachment contianer
                    attachmentContainer[RallyField.Artifact] = createUserStory.Reference;
                    attachmentContainer[RallyField.Content] = attachmentContentReference;
                    attachmentContainer[RallyField.Name] = attachmentPair.Value;
                    attachmentContainer[RallyField.Description] = RallyField.EmailAttachment;
                    attachmentContainer[RallyField.ContentType] = "file/";
                    //attachmentContainer[RallyField.size] = Omitted

                    //Create & associate the attachment
                    attachmentContainerCreateResult = _api.Create(RallyField.Attachment, attachmentContainer);
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
            toCreate[RallyField.WorkSpace] = workspace;
            toCreate[RallyField.Project] = project;
            toCreate[RallyField.Name] = userstoryName;
            createUserStory = _api.Create(RallyField.HierarchicalRequirement, toCreate);

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
                    attachmentContent[RallyField.Content] = attachmentPair.Value;
                    attachmentContentCreateResult = _api.Create(RallyField.AttachmentContent, attachmentContent);
                    attachmentContentReference = attachmentContentCreateResult.Reference;

                    //create attachment contianer
                    attachmentContainer[RallyField.Artifact] = createUserStory.Reference;
                    attachmentContainer[RallyField.Content] = attachmentContentReference;
                    attachmentContainer[RallyField.Name] = attachmentPair.Key;
                    attachmentContainer[RallyField.Description] = RallyField.EmailAttachment;
                    attachmentContainer[RallyField.ContentType] = "file/";

                    //Create & associate the attachment
                    attachmentContainerCreateResult = _api.Create(RallyField.Attachment, attachmentContainer);
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
            toCreate[RallyField.WorkSpace] = workspace;
            toCreate[RallyField.Project] = project;
            toCreate[RallyField.Name] = userstoryName;
            createUserStory = _api.Create(RallyField.HierarchicalRequirement, toCreate);

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
                    attachmentContent[RallyField.Content] = attachmentPair.Key;
                    attachmentContentCreateResult = _api.Create(RallyField.AttachmentContent, attachmentContent);
                    attachmentContentReference = attachmentContentCreateResult.Reference;

                    //create attachment contianer
                    attachmentContainer[RallyField.Artifact] = createUserStory.Reference;
                    attachmentContainer[RallyField.Content] = attachmentContentReference;
                    attachmentContainer[RallyField.Name] = attachmentPair.Value;
                    attachmentContainer[RallyField.Description] = RallyField.EmailAttachment;
                    attachmentContainer[RallyField.ContentType] = "file/";

                    //Create & associate the attachment
                    attachmentContainerCreateResult = _api.Create(RallyField.Attachment, attachmentContainer);
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

        #region: scenraio
        //even if duplicates are sent via email, windows does not download the duplicate
        //x.txt != x(1).txt
        //What if via sender and receiver communication the file is passed back and forth after making some changes
        //if the file in the first place does not exist in the attachmentPath then we need not worry about creating folders for each attachment type
        // we are mirroring the process of pulling emails and moviong to a processed folder
        
        public void downloadAttachments()
        {
            EnsureOutlookIsAuthenticated();
            Mailbox inbox = imap.SelectMailbox(Outlook.OutlookInboxFolder);
            int[] unreadIDs = inbox.Search(Outlook.OutlookUnseenMessages);
            int unreadMessagesLength = unreadIDs.Length;
            string[] attachmentPaths; 
            string destinationDirectory = SyncConstant.attachmentsProcessedDirectory;

            for (int i = 0; i < unreadMessagesLength; i++)
            {
                Message unreadMessageObject = inbox.Fetch.MessageObject(unreadIDs[i]);
                Console.WriteLine("Subject: " + unreadMessageObject.Subject);

                if (unreadMessageObject.Attachments.Count > 0)
                {
                    unreadMessageObject.Attachments.StoreToFolder(SyncConstant.attachmentsDirectory);
                }
            }
            Console.WriteLine("Break");

            attachmentPaths = Directory.GetFiles(SyncConstant.attachmentsDirectory);

            foreach (var file in attachmentPaths)
            {
                var fileName = Path.GetFileName(file);
                Console.WriteLine("Moving File: " + fileName);
                File.Move(file, Path.Combine(destinationDirectory, fileName));
            }
        }
        #endregion

        #region: Sync Rally 
        public void Sync(String worskpace, string project)
        {
            //Email variables
            List<Message> unreadMsgCollection = new List<Message>();
            unreadMsgCollection.Capacity = 25;
            Dictionary<string, string> attachmentsDictionary = new Dictionary<string, string>();

            //Rally variables
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyField.WorkSpace] = worskpace;
            toCreate[RallyField.Project] = project;
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
                Mailbox inbox = imap.SelectMailbox(Outlook.OutlookInboxFolder);
                int[] unread = inbox.Search(Outlook.OutlookUnseenMessages);
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
                            unreadMsgCollection[i].Subject = Outlook.NoSubject;
                        }
                            toCreate[RallyField.Name] = (unreadMsgCollection[i].Subject);
                            toCreate[RallyField.Description] = (unreadMsgCollection[i].BodyText.Text);
                            createUserStory = _api.Create(RallyField.HierarchicalRequirement, toCreate);

                        //check to see if message has attachments & then store them
                        if (unreadMsgCollection[i].Attachments.Count > 0)
                        {
                            //Do all the attachments from the email object get stored here?? _No it has to complete one iteration
                            unreadMsgCollection[i].Attachments.StoreToFolder(SyncConstant.attachmentsDirectory);
                        }

                        //reference the path where the attachments live for the [ith] message
                        attachmentPaths = Directory.GetFiles(SyncConstant.attachmentsDirectory);

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

                            //At this point the dictionary is populated, so we move each file to another directory - making each attachment unique to its related email object
                            //The directory we moved to might contain a file with the same text, since there is not use for the file after its uploaded to Rally, we can purge them
                            //Console.WriteLine("Moving File: " + attachmentFileName);
                            //File.Move(file, Path.Combine(SyncConstant.attachmentsProcessedDirectory, attachmentFileName));
                            //might seal "file already exists error" 
                            Console.WriteLine("Processing... "+file);
                            File.Delete(file); 
                        }

                        //Stage the attachment
                        foreach (KeyValuePair<string, string> attachmentPair in attachmentsDictionary)
                        {
                            try
                            {
                                //create attachment content
                                attachmentContent[RallyField.Content] = attachmentPair.Key;
                                attachmentContentCreateResult = _api.Create(RallyField.AttachmentContent, attachmentContent);
                                userStoryReference = attachmentContentCreateResult.Reference;

                                //create attachment contianer
                                attachmentContainer[RallyField.Artifact] = createUserStory.Reference;
                                attachmentContainer[RallyField.Content] = userStoryReference;
                                attachmentContainer[RallyField.Name] = attachmentPair.Value;
                                attachmentContainer[RallyField.Description] = RallyField.EmailAttachment;
                                attachmentContainer[RallyField.ContentType] = "file/";

                                //Create & associate the attachment
                                attachmentContainerCreateResult = _api.Create(RallyField.Attachment, attachmentContainer);
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
                        markAsUnreadFlag.Add(Outlook.OutlookSeenMessages);
                        inbox.RemoveFlags(item, markAsUnreadFlag);
                        inbox.MoveMessage(item, Outlook.OutlookProcessedFolder);
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
                imap.Disconnect();
            }

        }
        #endregion

    }
}


