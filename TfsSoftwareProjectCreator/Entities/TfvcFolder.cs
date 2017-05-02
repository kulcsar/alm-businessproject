using System.Collections.Generic;

namespace TfsSoftwareProjectCreator.Entities
{
    /// <summary>
    /// Represents folders in Team Foundation Version Control system
    /// </summary>
    public class TfvcFolder
    {
        public List<string> Folders { get; private set; }

        public TfvcFolder(object[,] excelContent)
        {
            Folders = GetFolders(excelContent);
        }

        private List<string> GetFolders(object[,] excelContent)
        {
            List<string> folders = new List<string>();

            for (int i = 1; i <= excelContent.GetLength(0); i++)
            {
                string objectType = (string)excelContent[i, 1];
                if (objectType == "TFVC Folder")
                {
                    string folderName = excelContent[i, 2].ToString();
                    folders.Add(folderName);
                }
            }

            return folders;
        }
    }
}