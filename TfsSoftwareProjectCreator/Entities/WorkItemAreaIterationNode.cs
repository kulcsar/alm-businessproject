namespace TfsSoftwareProjectCreator.Entities
{
    /// <summary>
    /// Represents a node in WorkItem Area\Iteration path
    /// </summary>
    public class WorkItemAreaIterationNode
    {
        public string Type { get; private set; }
        public string Path { get; private set; }

        public WorkItemAreaIterationNode(string type, string path)
        {
            Type = type;
            Path = path;
        }
    }
}
