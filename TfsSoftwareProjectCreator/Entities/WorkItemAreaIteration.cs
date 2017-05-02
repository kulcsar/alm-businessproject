using System.Collections.Generic;

namespace TfsSoftwareProjectCreator.Entities
{
    /// <summary>
    /// Represents WorkItem Area\Iteration paths
    /// </summary>
    public class WorkItemAreaIteration
    {
        public List<WorkItemAreaIterationNode> Nodes { get; private set; }
        
        public WorkItemAreaIteration(object[,] excelContent)
        {
            Nodes = GetNodes(excelContent);
        }

        private List<WorkItemAreaIterationNode> GetNodes(object[,] excelContent)
        {
            List<WorkItemAreaIterationNode> nodes = new List<WorkItemAreaIterationNode>();

            for (int i = 1; i <= excelContent.GetLength(0); i++)
            {
                string nodeType = (string)excelContent[i, 1];
                if (nodeType == "Area" || nodeType == "Iteration")
                {
                    string nodePath = excelContent[i, 2].ToString();
                    nodes.Add(new WorkItemAreaIterationNode(nodeType, nodePath));
                }
            }

            return nodes;
        }
    }
}