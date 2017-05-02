using System.Collections.Generic;

namespace TfsSoftwareProjectCreator.Entities
{
    /// <summary>
    /// Represents Identities (TFS Security Groups)
    /// </summary>
    public class SecurityGroup
    {
        public List<string> Identities { get; private set; }

        public SecurityGroup(object[,] excelContent)
        {
            Identities = GetIdentities(excelContent);
        }

        private List<string> GetIdentities(object[,] excelContent)
        {
            List<string> groups = new List<string>();

            for (int i = 1; i <= excelContent.GetLength(0); i++)
            {
                string objectType = (string)excelContent[i, 1];
                if (objectType == "Security Group")
                {
                    string groupName = excelContent[i, 2].ToString();
                    groups.Add(groupName);
                }
            }

            return groups;
        }
    }
}