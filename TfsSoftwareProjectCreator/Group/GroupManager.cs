using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.Server;
using System;
using System.Linq;

namespace TfsSoftwareProjectCreator.Group
{
    /// <summary>
    /// Provides TFS Group manipulation functionalities
    /// </summary>
    public class GroupManager
    {
        private readonly string _teamProjectCollectionUrl;
        private readonly string _teamProjectName;
        private readonly TfsTeamProjectCollection _tfsTeamProjectCollection;
        private readonly ICommonStructureService4 _commonStructureService4;
        private readonly ProjectInfo _projectInfo;
        private readonly IIdentityManagementService _identityManagementService;

        public GroupManager(string teamProjectCollectionUrl, string teamProjectName)
        {
            _teamProjectCollectionUrl = teamProjectCollectionUrl;
            _teamProjectName = teamProjectName;

            _tfsTeamProjectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(_teamProjectCollectionUrl));
            _tfsTeamProjectCollection.EnsureAuthenticated();
            _commonStructureService4 = _tfsTeamProjectCollection.GetService<ICommonStructureService4>();
            _projectInfo = _commonStructureService4.GetProjectFromName(_teamProjectName);
            _identityManagementService = _tfsTeamProjectCollection.GetService<IIdentityManagementService>();
        }

        /// <summary>
        /// Create TFS Group if not exists
        /// </summary>
        /// <param name="groupName"></param>
        /// <returns></returns>
        public TeamFoundationIdentity CreateGroup(string groupName)
        {
            // Check if group already exists
            TeamFoundationIdentity group = _identityManagementService
                .ListApplicationGroups(_projectInfo.Uri, ReadIdentityOptions.IncludeReadFromSource)
                .FirstOrDefault(g => string.Equals(g.DisplayName, groupName, StringComparison.OrdinalIgnoreCase));
            if (group == null)
            {
                // Prepare group name
                var groupNameToCreate = groupName.Replace($"[{_teamProjectName}]\\", "");
                // Group doesn't exist, create one
                var groupDescriptor = _identityManagementService.CreateApplicationGroup(_projectInfo.Uri, groupNameToCreate, null);
                group = _identityManagementService.ReadIdentity(groupDescriptor, MembershipQuery.None, ReadIdentityOptions.IncludeReadFromSource);
            }
            
            return group;
        }
    }
}
