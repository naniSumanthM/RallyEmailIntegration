public void CreateUserStory(string usName,
                            string usDescription,
                            string usWorkspace,
                            string usProject,
                            string usOwner,
                            string usIteration,
                            string usPlanEstimate)
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
