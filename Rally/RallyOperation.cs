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
        public const string ServerName = RallyField.serverID;

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
                _api.Authenticate(this.UserName, this.Password, ServerName, null, RallyField.allowSSO);
            }
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
                    var workspaceName = value[RallyField.nameForWSorUSorTA];
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
                    projectRequest.Fetch = new List<String>() { RallyField.nameForWSorUSorTA };

                    //Query for the projects
                    QueryResult projectQuery = _api.Query(projectRequest);
                    foreach (var project in projectQuery.Results)
                    {
                        Console.WriteLine(project[RallyField.nameForWSorUSorTA]);
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
            Request userStoryRequest = new Request(RallyField.hierarchicalRequirement);
            userStoryRequest.Workspace = workspaceRef;
            userStoryRequest.Project = projectRef;
            userStoryRequest.ProjectScopeUp = RallyField.projectScopeUp;
            userStoryRequest.ProjectScopeDown = RallyField.projectScopeDown;

            //fetch data from the story request
            userStoryRequest.Fetch = new List<string>()
        {
            RallyField.formattedID, RallyField.nameForWSorUSorTA, RallyField.owner
        };

            try
            {
                //query the items in the list
                userStoryRequest.Query = new Query(QueryField.lastUpdatDate, Query.Operator.GreaterThan, QueryField.dateGreaterThan);
                QueryResult userStoryResult = _api.Query(userStoryRequest);

                //iterate through the userStory Collection
                foreach (var userStory in userStoryResult.Results)
                {
                    var userStoryOwner = userStory[RallyField.owner];
                    if (userStoryOwner != null)
                    {
                        var USOwner = userStoryOwner[QueryField.referenceObject];
                        Console.WriteLine(userStory[RallyField.formattedID] + ":" + userStory[RallyField.nameForWSorUSorTA] + Environment.NewLine + QueryField.usMessage + USOwner + Environment.NewLine);
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
            Request userStoryRequest = new Request(RallyField.hierarchicalRequirement);
            userStoryRequest.Workspace = workspaceRef;
            userStoryRequest.Project = projectRef;
            userStoryRequest.ProjectScopeUp = RallyField.projectScopeUp;
            userStoryRequest.ProjectScopeDown = RallyField.projectScopeDown;

            //fetch US data in the form of a list
            userStoryRequest.Fetch = new List<string>()
        {
            RallyField.formattedID, RallyField.nameForWSorUSorTA, RallyField.capitalTasks, RallyField.estimate, RallyField.state, RallyField.owner
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
                    var userStoryOwner = userStory[RallyField.owner];
                    if (userStoryOwner != null) //return only US who have an assigned owner
                    {
                        var USOwner = userStoryOwner[QueryField.referenceObject];
                        Console.WriteLine(userStory[RallyField.formattedID] + ":" + userStory[RallyField.nameForWSorUSorTA]);
                        Console.WriteLine(QueryField.usMessage + USOwner);
                    }

                    //Task Request
                    Request taskRequest = new Request(userStory[RallyField.capitalTasks]);
                    QueryResult taskResult = _api.Query(taskRequest);
                    if (taskResult.TotalResultCount > 0)
                    {
                        foreach (var task in taskResult.Results)
                        {
                            var taskName = task[RallyField.nameForWSorUSorTA];
                            var owner = task[RallyField.owner];
                            var taskState = task[RallyField.state];
                            var taskEstimate = task[RallyField.estimate];
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
            toCreate[RallyField.nameForWSorUSorTA] = usName;
            toCreate[RallyField.description] = usDescription;
            toCreate[RallyField.workSpace] = usWorkspace;
            toCreate[RallyField.project] = usProject;
            toCreate[RallyField.owner] = usOwner;
            toCreate[RallyField.iteration] = usIteration;
            toCreate[RallyField.planEstimate] = usPlanEstimate;

            //use try and catch to create push a US to the specific workspace within a specific project
            try
            {
                Console.WriteLine("<<Creating US>>");
                CreateResult createUserStory = _api.Create(RallyField.hierarchicalRequirement, toCreate);
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
            toCreate[RallyField.nameForWSorUSorTA] = taskName;
            toCreate[RallyField.description] = taskDescription;
            toCreate[RallyField.owner] = taskOwner;
            toCreate[RallyField.estimate] = taskEstimate;
            toCreate[RallyField.workProduct] = userStoryReference;

            //create a task and attach it to a userStory
            try
            {
                Console.WriteLine("<<Creating TA>>");
                CreateResult createTask = _api.Create(RallyField.smallTasks, toCreate);
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
            toCreate[RallyField.workSpace] = usWorkspace;
            toCreate[RallyField.project] = usProject;

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
                    toCreate[RallyField.nameForWSorUSorTA] = usList[i];
                    CreateResult cr = _api.Create(RallyField.hierarchicalRequirement, toCreate);
                }
            }
            catch (WebException)
            {
                Console.WriteLine(QueryField.webExceptionMessage);
            }
            Console.WriteLine("End");
        }
        #endregion

        #region: Create US'z through unread Mail Messages
        //testing a list of userstories
        public void UserStorySync(string usWorkspace, string usProject)
        {
            //Authenticate with Rally
            this.EnsureRallyIsAuthenticated();

            //List to add the unreadMessages
            List<Message> unreadMessageList = new List<Message>();
            unreadMessageList.Capacity = 25;

            //Set up the US
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyField.workSpace] = usWorkspace;
            toCreate[RallyField.project] = usProject;

            Console.WriteLine("Start");
            try
            {
                //Authenticate with Imap
                Imap4Client imap = new Imap4Client();
                imap.ConnectSsl(Outlook.outlookHost, Outlook.outlookPort);
                imap.Login(Outlook.outlookUsername, Outlook.outlookPassword);

                //setup Imap enviornment
                Mailbox inbox = imap.SelectMailbox(Outlook.outlookInboxFolder);
                int[] unread = inbox.Search(Outlook.outlookUnread);
                Console.WriteLine("Unread Messages: " + unread.Length);

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
                        toCreate[RallyField.nameForWSorUSorTA] = (unreadMessageList[i].Subject);
                        toCreate[RallyField.description] = (unreadMessageList[i].BodyText.Text);
                        CreateResult cr = _api.Create(RallyField.hierarchicalRequirement, toCreate);
                    }

                    //Move Fetched Messages
                    foreach (var item in unread)
                    {
                        inbox.MoveMessage(item, Outlook.outlookProcessedFolder);
                    }
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

        #region: Refined UserStory Sync    
        ///<summary>
        ///After each email item is moved to the processed folder, it has to be marked as unread for future reference
        /// </summary>   
        /// 
        public void UserStorySyncRefined(string usWorkspace, string usProject)
        {
            //Authenticate with Rally
            this.EnsureRallyIsAuthenticated();

            //List to add the unreadMessages
            List<Message> unreadMessageList = new List<Message>();
            unreadMessageList.Capacity = 25;

            //Set up the US
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyField.workSpace] = usWorkspace;
            toCreate[RallyField.project] = usProject;

            Console.WriteLine("Start");
            try
            {
                //Authenticate with Imap
                Imap4Client imap = new Imap4Client();
                imap.ConnectSsl(Outlook.outlookHost, Outlook.outlookPort);
                imap.Login(Outlook.outlookUsername, Outlook.outlookPassword);

                //setup Imap enviornment
                Mailbox inbox = imap.SelectMailbox(Outlook.outlookInboxFolder);
                int[] unread = inbox.Search(Outlook.outlookUnread);
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
                        toCreate[RallyField.nameForWSorUSorTA] = (unreadMessageList[i].Subject);
                        toCreate[RallyField.description] = (unreadMessageList[i].BodyText.Text);
                        CreateResult cr = _api.Create(RallyField.hierarchicalRequirement, toCreate);
                    }

                    //Move Fetched Messages
                    foreach (var item in unread)
                    {
                        markAsUnreadFlag.Add("Seen");
                        inbox.RemoveFlags(item, markAsUnreadFlag);
                        inbox.MoveMessage(item, Outlook.outlookProcessedFolder);
                    }
                    //TODO: Safer to write another loop and iterate over the procesed folder, but that will crawl
                    //the entire inbox and mark the read items as unread.
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

        #region: imageToBase64

        // Converts image to Base 64 Encoded string
        public static string imageToBase64(Image image, System.Drawing.Imaging.ImageFormat format)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, format);
                // Convert Image to byte[]
                byte[] imageBytes = ms.ToArray();

                // Convert byte[] to Base64 String
                string base64String = Convert.ToBase64String(imageBytes);

                return base64String;
            }
        }
        #endregion

        #region: add an attachment to an existing user story
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
            Image myImage = Image.FromFile(fullImageFile);

            // Convert Image to Base64 format
            string imageBase64String = imageToBase64(myImage, System.Drawing.Imaging.ImageFormat.Png);

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

        #region: create userStory with attachment
        ///<summary>
        ///Method that creates a user story with an attachment (takes only 1 png attachment)
        /// </summary>

        public void createUsWithAttachment(string workspace, string project, string userStoryName, string userStoryDescription)
        {
            //Authentication
            this.EnsureRallyIsAuthenticated();

            //UserStory Setup
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyField.workSpace] = workspace;
            toCreate[RallyField.project] = project;
            toCreate[RallyField.nameForWSorUSorTA] = userStoryName;
            toCreate[RallyField.description] = userStoryDescription;

            //get the image reference - assume that this is where the image lives in respect to the path after being pulled from outlook
            String imageFilePath = "C:\\Users\\maddirsh\\Desktop\\";
            String imageFileName = "webException.png";
            String fullImageFile = imageFilePath + imageFileName;
            Image myImage = Image.FromFile(fullImageFile);

            // Convert Image to Base64 format
            string imageBase64String = imageToBase64(myImage, System.Drawing.Imaging.ImageFormat.Png);

            // Length calculated from Base64String converted back
            var imageNumberBytes = Convert.FromBase64String(imageBase64String).Length;

            // DynamicJSONObject for AttachmentContent
            DynamicJsonObject myAttachmentContent = new DynamicJsonObject();
            myAttachmentContent[RallyField.content] = imageBase64String;

            try
            {
                //create user story
                CreateResult createUserStory = _api.Create(RallyField.hierarchicalRequirement, toCreate);

                //create attachment
                CreateResult myAttachmentContentCreateResult = _api.Create(RallyField.attachmentContent, myAttachmentContent);
                String myAttachmentContentRef = myAttachmentContentCreateResult.Reference;

                // DynamicJSONObject for Attachment Container
                DynamicJsonObject myAttachment = new DynamicJsonObject();
                myAttachment[RallyField.artifact] = createUserStory.Reference;
                myAttachment[RallyField.content] = myAttachmentContentRef;
                myAttachment[RallyField.nameForWSorUSorTA] = "AttachmentFromREST.png";
                myAttachment[RallyField.description] = "Email Attachment";
                myAttachment[RallyField.contentType] = "image/png"; //Method to identify the fileType.java
                myAttachment[RallyField.size] = imageNumberBytes;

                //create & associate the attachment
                CreateResult myAttachmentCreateResult = _api.Create(RallyField.attachment, myAttachment);
                Console.WriteLine("Created User Story: "+ createUserStory.Reference);
            }
            catch (WebException e)
            {
                Console.WriteLine(e.Message);
            }
        }

        #endregion

        #region: multipleAttachments
        /// <summary>
        /// Method that can attach multiple attachments to a user story
        /// </summary>
        /// <param name="workspace"></param>
        /// <param name="project"></param>
        /// <param name="userStoryName"></param>
        /// <param name="userStoryDescription"></param>


        public void userStoryWithMultipleAttachments(string workspace, string project, string userStoryName, string userStoryDescription)
        {
            //Authentication
            this.EnsureRallyIsAuthenticated();

            //UserStory Setup
            DynamicJsonObject toCreate = new DynamicJsonObject();
            toCreate[RallyField.workSpace] = workspace;
            toCreate[RallyField.project] = project;
            toCreate[RallyField.nameForWSorUSorTA] = userStoryName;
            toCreate[RallyField.description] = userStoryDescription;

            Dictionary<string, int> imageD = new Dictionary<string, int>();
            Image myImage;
            string base64ImageString;
            int imageByteSize = 0;
            string myAttachmentContentRef;
            DynamicJsonObject myAttachmentContent = new DynamicJsonObject();
            DynamicJsonObject myAttachmentContainer = new DynamicJsonObject();

            string[] filePaths = Directory.GetFiles("C:\\Users\\maddirsh\\Desktop\\RallyAttachments");

            foreach (string filePath in filePaths)
            {
                myImage = Image.FromFile(filePath);
                base64ImageString = imageToBase64(myImage, System.Drawing.Imaging.ImageFormat.Png);
                imageByteSize = Convert.FromBase64String(base64ImageString).Length;
                imageD.Add(base64ImageString, imageByteSize);
                Console.WriteLine(filePath);
            }

            //Create US
            CreateResult createUserStory = _api.Create(RallyField.hierarchicalRequirement, toCreate);
  
            foreach (KeyValuePair<string, int> imagepair in imageD)
            {
                try
                {
                    myAttachmentContent[RallyField.content] = imagepair.Key;

                    CreateResult myAttachmentContentCreateResult = _api.Create(RallyField.attachmentContent, myAttachmentContent);
                    myAttachmentContentRef = myAttachmentContentCreateResult.Reference;

                    myAttachmentContainer[RallyField.artifact] = createUserStory.Reference;
                    myAttachmentContainer[RallyField.content] = myAttachmentContentRef;
                    myAttachmentContainer[RallyField.nameForWSorUSorTA] = "AttachmentFromREST.png";
                    myAttachmentContainer[RallyField.description] = "Email Attachment";
                    myAttachmentContainer[RallyField.contentType] = "image/png";
                    myAttachmentContainer[RallyField.size] = imagepair.Value;

                    //create & associate the attachment
                    CreateResult myAttachmentCreateResult = _api.Create(RallyField.attachment, myAttachmentContainer);
                    Console.WriteLine("Created User Story: " + createUserStory.Reference);

                }
                catch (WebException e)
                {
                    Console.WriteLine(e.Message);
                }
            }

            
        }
        #endregion
    }
}

