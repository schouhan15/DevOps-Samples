
#### QueryTFS - A simple .Net console app (4.7.2 Runtime Required) to run and save work-item query results across all projects in all collections.

##### Pre-Requisites: 

* The account running QueryTFS.exe must be a TFS/Azure DevOps Server Administrator.
* Supports only Azure DevOps Server


##### How-To: 

* Update the work-item query you want to run across all projects in the code, restore nuget packages, build and run! 

##### Note:

* The code delimits ',' as default, if you have ',' in your work-item query result (example, title [as a customer, I ...] you'd have to modify the code accordingly 

More info on the classes/methods used, please refer Taylor's excellent [post](https://blogs.msdn.microsoft.com/taylaf/2010/01/26/retrieve-the-list-of-team-project-collections-from-tfs-2010-client-apis/)
