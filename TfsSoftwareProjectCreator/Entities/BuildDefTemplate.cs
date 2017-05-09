using System;
using Microsoft.TeamFoundation.Build.WebApi;

namespace TfsSoftwareProjectCreator.Entities
{
    /// <summary>
    /// Represents Build Definition Template  by which new build definition can be created
    /// </summary>
    public class BuildDefTemplate
    {
        public string BuildDefinitionTemplateName { get; private set; }

        public BuildDefTemplate(object[,] excelContent)
        {
            BuildDefinitionTemplateName = GetTemplateName(excelContent);
        }

        private string GetTemplateName(object[,] excelContent)
        {
            for (int i = 1; i <= excelContent.GetLength(0); i++)
            {
                string objectType = (string)excelContent[i, 1];
                if (objectType == "Build Definition Template")
                {
                    return (string)excelContent[i, 2];
                }
            }

            return null;
        }
    }
}