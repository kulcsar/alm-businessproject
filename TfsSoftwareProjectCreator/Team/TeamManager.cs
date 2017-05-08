using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.ProcessConfiguration.Client;
using Microsoft.TeamFoundation.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TfsSoftwareProjectCreator.Entities;

namespace TfsSoftwareProjectCreator.Team
{
    /// <summary>
    /// Provides TFS Team manipulation functionalities
    /// </summary>
    public class TeamManager
    {
        private readonly string _teamProjectCollectionUrl;
        private readonly string _teamProjectName;
        private readonly string _softwareProjectName;
        private readonly string _softwareProjectDescription;
        private readonly TfsTeamProjectCollection _tfsTeamProjectCollection;
        private readonly ICommonStructureService4 _commonStructureService4;
        private readonly TfsTeamService _tfsTeamService;
        private readonly TeamSettingsConfigurationService _teamSettingsConfigurationService;
        private readonly ProjectInfo _projectInfo;

        public TeamManager(string teamProjectCollectionUrl, string teamProjectName, string softwareProjectName, string softwareProjectDescription)
        {
            _teamProjectCollectionUrl = teamProjectCollectionUrl;
            _teamProjectName = teamProjectName;
            _softwareProjectName = softwareProjectName;
            _softwareProjectDescription = softwareProjectDescription;

            _tfsTeamProjectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(_teamProjectCollectionUrl));
            _tfsTeamProjectCollection.EnsureAuthenticated();
            _commonStructureService4 = _tfsTeamProjectCollection.GetService<ICommonStructureService4>();
            _tfsTeamService = _tfsTeamProjectCollection.GetService<TfsTeamService>();
            _teamSettingsConfigurationService = _tfsTeamProjectCollection.GetService<TeamSettingsConfigurationService>();
            _projectInfo = _commonStructureService4.GetProjectFromName(_teamProjectName);
        }

        /// <summary>
        /// Create Area and Iteration paths
        /// </summary>
        /// <param name="workItemAreaIteration"></param>
        public void CreateWorkItemAreaIteration(WorkItemAreaIteration workItemAreaIteration)
        {
            foreach (var node in workItemAreaIteration.Nodes)
            {
                AddNode(node.Path, _teamProjectName, node.Type);
            }
        }

        /// <summary>
        /// Create Area and Iteration path recursively
        /// </summary>
        /// <param name="elementPath"></param>
        /// <param name="projectName"></param>
        /// <param name="nodeType"></param>
        /// <returns></returns>
        private NodeInfo AddNode(string elementPath, string projectName, string nodeType)
        {
            NodeInfo retVal;

            //Check if this path already exsist
            try
            {
                retVal = _commonStructureService4.GetNodeFromPath(elementPath);
                if (retVal != null)
                {
                    return null; //already exists
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("The following node does not exist"))
                {
                    //just means that this path is not exist and we can continue.
                }
                else
                {
                    throw ex;
                }
            }

            int backSlashIndex = elementPath.LastIndexOf("\\");
            string newPathName = elementPath.Substring(backSlashIndex + 1);
            string newPath = (backSlashIndex == 0 ? string.Empty : elementPath.Substring(0, backSlashIndex));
            NodeInfo previousPath = null;
            try
            {
                previousPath = _commonStructureService4.GetNodeFromPath(newPath);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Invalid path."))
                {
                    //just means that this path is not exist and we can continue.
                    previousPath = null;
                }
                else
                {
                    throw ex;
                }
            }
            if (previousPath == null)
            {
                //call this method to create the parent paths.
                previousPath = AddNode(newPath, projectName, nodeType);
            }
            string newPathUri = _commonStructureService4.CreateNode(newPathName, previousPath.Uri);

            return _commonStructureService4.GetNode(newPathUri);
        }

        /// <summary>
        /// Create TFS Team if not exists
        /// </summary>
        /// <returns></returns>
        public TeamFoundationTeam CreateTeam()
        {
            // Check team already exists
            var teams = _tfsTeamService.QueryTeams(_projectInfo.Uri.ToString());
            var existingTeam = teams.FirstOrDefault(t => t.Name == _softwareProjectName);
            if (existingTeam != null)
            {
                return existingTeam;
            }

            //Create TFS Team
            TeamFoundationTeam team = _tfsTeamService.CreateTeam(
                _projectInfo.Uri.ToString(), _softwareProjectName, _softwareProjectDescription, null);

          
            //Set the IterationPaths and BacklogIterationPath for the TFS Team
            var teamConfiguration = _teamSettingsConfigurationService.GetTeamConfigurations(new[] { team.Identity.TeamFoundationId });
            TeamConfiguration tconfig = teamConfiguration.FirstOrDefault();
            TeamSettings ts = tconfig.TeamSettings;
            ts.IterationPaths = new string[] { $"{_projectInfo.Name}\\{_softwareProjectName}\\Sprint 1" };
            ts.BacklogIterationPath = $"{_projectInfo.Name}\\{_softwareProjectName}";
            TeamFieldValue tfv = new TeamFieldValue();
            tfv.IncludeChildren = true;
            tfv.Value = ts.BacklogIterationPath;
            ts.TeamFieldValues = new TeamFieldValue[] { tfv };
            _teamSettingsConfigurationService.SetTeamSettings(tconfig.TeamId, ts);

            return team;
        }
    }
}
