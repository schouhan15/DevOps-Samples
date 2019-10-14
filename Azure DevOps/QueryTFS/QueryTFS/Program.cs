using System;
using System.Collections.ObjectModel;
using System.IO;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using WorkItem = Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem;

namespace QueryTFS
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get the Server URL
            Console.WriteLine("************************************************************************************************************************");
            Console.Write("Enter the Server URL:\t");
            Uri configurationServerUri = new Uri(Console.ReadLine());
            Console.WriteLine("\n************************************************************************************************************************");


            TfsConfigurationServer configurationServer =
                    TfsConfigurationServerFactory.GetConfigurationServer(configurationServerUri);

            CatalogNode configurationServerNode = configurationServer.CatalogNode;

            // Query the children of the configuration server node for all of the team project collection nodes
            ReadOnlyCollection<CatalogNode> tpcNodes = configurationServerNode.QueryChildren(
                    new Guid[] { CatalogResourceTypes.ProjectCollection },
                    false,
                    CatalogQueryOptions.None);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Success : Connected to the Server");
            Console.ResetColor();

            Console.WriteLine("\n************************************************************************************************************************");
            Console.Write("Enter the CSV File Path (Press Enter for default D:\\file.csv)  :\t");

            string fileInput = Console.ReadLine();
            Console.WriteLine("\n************************************************************************************************************************");

            string filePath;

            if (string.IsNullOrEmpty(fileInput))
            {
                filePath = @"D:\file.csv";
            }
            else
            {
                filePath = fileInput;
            }

            //delete if the file exists

            if (File.Exists(filePath))
            {
                try
                {
                    File.Delete(filePath);
                }
                catch
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine("Unable to remove the file, check if the file is open/in use\n");
                    Environment.Exit(1);
                }
            }

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Created a new file at {0}\n", filePath);
            Console.ResetColor();

            // initialize counter
            int count = 0;

            foreach (CatalogNode tpcNode in tpcNodes)
            {
                // Use tpcNode.Resource to get the details for each team project collection.
                ServiceDefinition tpcServiceDefinition = tpcNode.Resource.ServiceReferences["Location"];

                ILocationService configLocationService = configurationServer.GetService<ILocationService>();
                Uri tpcUri = new Uri(configLocationService.LocationForCurrentConnection(tpcServiceDefinition));

                // Actually connect to the team project collection
                TfsTeamProjectCollection tpc = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(tpcUri);

                CatalogNode teamProjectCollection = tpcNodes[count];

                Console.WriteLine("================================================================\n");
                Console.WriteLine("Currently in Collection '" + tpcNodes[count].Resource.DisplayName + "' :");
                Console.WriteLine();

                //    Get the Team Projects that belong to the Team Project Collection
                ReadOnlyCollection<CatalogNode> teamProjects = teamProjectCollection.QueryChildren(
                   new Guid[] { CatalogResourceTypes.TeamProject },
                   false,
                   CatalogQueryOptions.None);


                foreach (CatalogNode teamProject in teamProjects)
                {
                    string teamProj = teamProject.Resource.DisplayName;

                    // initialise an empty string
                    string info = String.Empty;

                    Console.WriteLine("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
                    Console.WriteLine("Currently in Project '" + teamProject.Resource.DisplayName + "' :");

                    //Microsoft.TeamFoundation.TeamFoundationServiceUnavailableException


                    WorkItemStore workItemStore = null;
                    try
                    {
                        workItemStore = new WorkItemStore(tpc);
                    }
                    catch (Microsoft.TeamFoundation.TeamFoundationServiceUnavailableException ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("Collection '{0}' seems to be offline, please check the status of the collection", tpcNodes[count].Resource.DisplayName);
                        Console.ResetColor();
                        continue;
                    }
                    // write the wiql query here

                    if (workItemStore != null)
                    {
                        try
                        {
                            string WitQuery = "SELECT [System.Id]," +
                                                "[System.WorkItemType]," +
                                                "[System.Title]," +
                                                "[System.AssignedTo]," +
                                                "[System.State]," +
                                                "[System.Tags]" +
                                                "FROM workitems " +
                                                "WHERE [System.TeamProject] = '" + teamProj + "' " +
                                                "AND[System.WorkItemType] <> '' " +
                                                "AND[System.State] <> '' ";

                            Query query = new Query(workItemStore, WitQuery);

                            WorkItemCollection wic = query.RunQuery();

                            foreach (WorkItem item in wic)
                            {
                                // add additional entries seperated by ',' 
                                info += String.Format("{0},{1},{2},{3},{4},{5}\n",
                                    teamProj,
                                    item.Id,
                                    item.Title,
                                    item.Fields[CoreField.AssignedTo].Value,
                                    item.State, item.Tags);
                            }


                            //Console.WriteLine(info);

                            //add to csv

                            File.AppendAllText(filePath, info);

                            Console.ForegroundColor = ConsoleColor.Cyan;
                            Console.WriteLine("Updated query output to the file {0}\n", filePath);
                            Console.ResetColor();


                        }
                        catch (Microsoft.TeamFoundation.WorkItemTracking.Client.ValidationException ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine("An error occured while running the query:\n" + ex.Message);
                            Console.ResetColor();

                        }

                    }
                }
                Console.WriteLine("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\n");

                count++;
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("file {0} Successfully Updated with Query Results!\n", filePath);
            Console.ResetColor();
            Console.WriteLine("Press 'Enter' to exit\n");
            Console.Read();

        }
    }
}
