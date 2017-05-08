using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using System;
using System.Collections.Generic;

namespace TfsSoftwareProjectCreator.Security
{
    /// <summary>
    /// Provides TFS Security Group membership manipulation functionalities
    /// </summary>
    public class SecurityManager
    {
        private readonly string _teamProjectCollectionUrl;
        private readonly TfsTeamProjectCollection _tfsTeamProjectCollection;
        private readonly IIdentityManagementService _identityManagementService;

        public SecurityManager(string teamProjectCollectionUrl)
        {
            _teamProjectCollectionUrl = teamProjectCollectionUrl;
            _tfsTeamProjectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(_teamProjectCollectionUrl));
            _tfsTeamProjectCollection.EnsureAuthenticated();
            _identityManagementService = _tfsTeamProjectCollection.GetService<IIdentityManagementService>();
        }

        /// <summary>
        /// Add members into group
        /// </summary>
        /// <param name="identity"></param>
        /// <param name="members"></param>
        public void AddMembers(TeamFoundationIdentity identity, List<TeamFoundationIdentity> members)
        {
            foreach (var member in members)
            {
                try
                {
                    _identityManagementService.AddMemberToApplicationGroup(identity.Descriptor, member.Descriptor);
                }
                catch(Exception ex)
                {
                    Console.WriteLine($"Error happened while adding member to group: {ex.Message}");
                }
            }
        }
    }
}
