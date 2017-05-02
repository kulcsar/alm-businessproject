using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using System;
using System.IO;
using TfsSoftwareProjectCreator.Entities;

namespace TfsSoftwareProjectCreator.Repository
{
    /// <summary>
    /// Provides Team Foundation Version Control manipulation functionalities
    /// </summary>
    public class RepositoryManager
    {
        private readonly string _teamProjectCollectionUrl;
        private readonly string _teamProjectName;
        private readonly string _softwareProjectName;
        private readonly TfsTeamProjectCollection _tfsTeamProjectCollection;
        private readonly VersionControlServer _versionControlServer;

        public RepositoryManager(string teamProjectCollectionUrl, string teamProjectName, string softwareProjectName)
        {
            _teamProjectCollectionUrl = teamProjectCollectionUrl;
            _teamProjectName = teamProjectName;
            _softwareProjectName = softwareProjectName;

            _tfsTeamProjectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(_teamProjectCollectionUrl));
            _tfsTeamProjectCollection.EnsureAuthenticated();
            _versionControlServer = _tfsTeamProjectCollection.GetService<VersionControlServer>();
        }

        /// <summary>
        /// Create TFVC folders
        /// </summary>
        /// <param name="tfvcFolder"></param>
        public void CreateFolders(TfvcFolder tfvcFolder)
        {
            var workspaceName = Guid.NewGuid().ToString();

            // Create a temporary workspace
            var workspace = _versionControlServer.CreateWorkspace(workspaceName,
                                                                  _tfsTeamProjectCollection.AuthorizedIdentity.UniqueName,
                                                                  $"Temporary workspace for project {_softwareProjectName}");

            // Create a temporary local folder for workspace
            string localPath = $"C:\\{workspaceName}";
            Directory.CreateDirectory(localPath);

            // For this workspace, map a server folder to a local folder
            workspace.Map($"$/{_teamProjectName}", localPath);

            // Create directories for software project
            string rootPath = Path.Combine(localPath, _softwareProjectName);
            
            // Always create root folder i.e. $/TeamProjectName/SoftwareProjectName
            Directory.CreateDirectory(rootPath);
            
            // Create folders 
            foreach (var folder in tfvcFolder.Folders)
            {
                var folderPath = RemoveRootPath($"$/{_teamProjectName}/{_softwareProjectName}", folder);
                Directory.CreateDirectory(Path.Combine(rootPath, folderPath));
            }

            // Check in folders
            workspace.PendAdd(rootPath, true);
            var pendingChanges = workspace.GetPendingChanges();            
            int changesetNumber = workspace.CheckIn(pendingChanges, $"Base folder structure created for project: {_softwareProjectName}");

            // Delete the temporary workspace
            workspace.Delete();

            // Create the temporary local folder
            Directory.Delete(localPath, true);
        }

        /// <summary>
        /// Remove prefix from path
        /// </summary>
        /// <param name="rootPath"></param>
        /// <param name="folder"></param>
        /// <returns></returns>
        private string RemoveRootPath(string rootPath, string folder)
        {
            folder = folder.Replace(rootPath, "");
            folder = folder.Replace("/", "\\");
            if (folder.Length > 0 && folder[0] == '\\')
            {
                folder = folder.Substring(1);
            }
            return folder;
        }
    }
}