namespace TfsSoftwareProjectCreator.Entities
{
    /// <summary>
    /// Represents base data of Business Project
    /// </summary>
    public class BusinessProject
    {
        private const string TEAM_PROJECT_COLLECTION = "Team Project Collection";
        private const string TEAM_PROJECT = "Team Project";
        private const string BUSINESS_PROJECT_NAME = "Business Project Name";
        private const string BUSINESS_PROJECT_DESCRIPTION = "Business Project Description";

        public string TeamProjectCollectionUrl { get; private set; }
        public string TeamProjectName { get; private set; }
        public string BusinessProjectName { get; private set; }
        public string BusinessProjectDescription { get; private set; }

        public BusinessProject(object[,] excelContent)
        {
            TeamProjectCollectionUrl = GetValue(excelContent, TEAM_PROJECT_COLLECTION);
            TeamProjectName = GetValue(excelContent, TEAM_PROJECT);
            BusinessProjectName = GetValue(excelContent, BUSINESS_PROJECT_NAME);
            BusinessProjectDescription = GetValue(excelContent, BUSINESS_PROJECT_DESCRIPTION);
        }

        private string GetValue(object[,] excelContent, string fieldName)
        {
            for (int i = 1; i <= excelContent.GetLength(0); i++)
            {
                string name = (string)excelContent[i, 1];
                if (name == fieldName)
                {
                    string value = excelContent[i, 2].ToString();
                    return value;
                }
            }

            return null;           
        }
    }
}