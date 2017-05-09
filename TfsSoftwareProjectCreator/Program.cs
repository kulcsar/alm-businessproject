using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using TfsSoftwareProjectCreator.Build;
using TfsSoftwareProjectCreator.Entities;
using TfsSoftwareProjectCreator.Excel;
using TfsSoftwareProjectCreator.Group;
using TfsSoftwareProjectCreator.Repository;
using TfsSoftwareProjectCreator.Security;
using TfsSoftwareProjectCreator.Team;

namespace TfsSoftwareProjectCreator
{
    /// <summary>
    /// Main program
    /// </summary>
    class Program
    {
        private const string WORKSHEETNAME = "Objects";

        /// <summary>
        /// Entry point
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("TfsSoftwareProjectCreator tool.\n Usage: TfsSoftwareProjectCreator <Excel input file absolute path>\n");
                return;
            }

            // Read argument
            string excelFilePath = args[0];

            // Read content from Excel sheet
            var excelContent = ExcelReader.GetExcelContent(excelFilePath, WORKSHEETNAME);

            // Read structured information from content 
            var businessProject = new BusinessProject(excelContent);
            var workItemAreaIteration = new WorkItemAreaIteration(excelContent);
            var securityGroup = new SecurityGroup(excelContent);
            var tfvcFolders = new TfvcFolder(excelContent);
            var buildDefTemplate = new BuildDefTemplate(excelContent);

            //Create new Team with Areas, Iterations
            var team = CreateTeam(businessProject, workItemAreaIteration);

            //Create new TFS Groups
            var tfsGroups = CreateGroups(businessProject, securityGroup);

            //Add TFS Groups as a member to the TFS Team
            AddTfsTeamMembers(businessProject, team.Identity, tfsGroups);

            //Create TFVC repository folders 
            if (tfvcFolders.Folders.Any())
            {
                CreateTfvcRepositoryFolders(businessProject, tfvcFolders);
            }

            // NEXT: Create GIT repository and folders

            // Create Build definition 
            CreateBuildDefinition(businessProject, buildDefTemplate);

            Console.WriteLine("TfsSoftwareProjectCreator finished. Press Enter to finish...");
            Console.ReadLine();
        }
        
        /// <summary>
        /// Create new Team with Areas, Iterations
        /// </summary>
        /// <param name="businessProject"></param>
        /// <param name="workItemAreaIteration"></param>
        /// <returns></returns>
        private static TeamFoundationTeam CreateTeam(BusinessProject businessProject, 
                                                        WorkItemAreaIteration workItemAreaIteration)
        {
            var teamManager = new TeamManager(businessProject.TeamProjectCollectionUrl,
                                                businessProject.TeamProjectName,
                                                businessProject.BusinessProjectName,
                                                businessProject.BusinessProjectDescription);
            teamManager.CreateWorkItemAreaIteration(workItemAreaIteration);
            return teamManager.CreateTeam();           
        }

        /// <summary>
        /// Create new TFS Security Groups for software project roles
        /// </summary>
        /// <param name="businessProject"></param>
        /// <param name="securityGroup"></param>
        /// <returns></returns>
        private static List<TeamFoundationIdentity> CreateGroups(BusinessProject businessProject, 
                                                                    SecurityGroup securityGroup)
        {
            List<TeamFoundationIdentity> securityGroups = new List<TeamFoundationIdentity>();

            var groupManager = new GroupManager(businessProject.TeamProjectCollectionUrl, businessProject.TeamProjectName);
            foreach (var groupName in securityGroup.Identities)
            {   
                var tfsGroup = groupManager.CreateGroup(groupName);
                securityGroups.Add(tfsGroup);
            }

            return securityGroups;
        }

        /// <summary>
        /// Add members to identity
        /// </summary>
        /// <param name="businessProject"></param>
        /// <param name="identity"></param>
        /// <param name="members"></param>
        private static void AddTfsTeamMembers(BusinessProject businessProject, 
                                                TeamFoundationIdentity identity, 
                                                List<TeamFoundationIdentity> members)
        {
            var securityManager = new SecurityManager(businessProject.TeamProjectCollectionUrl);
            securityManager.AddMembers(identity, members);
        }

        /// <summary>
        /// Create Team Foundation Version Control repository folders
        /// </summary>
        /// <param name="businessProject"></param>
        /// <param name="tfvcFolder"></param>
        private static void CreateTfvcRepositoryFolders(BusinessProject businessProject, TfvcFolder tfvcFolder)
        {
            var respositoryManager = new RepositoryManager(businessProject.TeamProjectCollectionUrl, 
                                                            businessProject.TeamProjectName,
                                                            businessProject.BusinessProjectName);
            respositoryManager.CreateTfvcFolders(tfvcFolder);
        }

        /// <summary>
        /// Create Build definition based on template
        /// </summary>
        /// <param name="businessProject"></param>
        /// <param name="buildDefTemplate"></param>
        private static void CreateBuildDefinition(BusinessProject businessProject, BuildDefTemplate buildDefTemplate)
        {
            var buildManager = new BuildManager(businessProject.TeamProjectCollectionUrl,
                                                businessProject.TeamProjectName,
                                                businessProject.BusinessProjectName);

            buildManager.CreateBuildDefinition(buildDefTemplate);
        }
    }
}
