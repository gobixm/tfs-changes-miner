using System;

namespace TfsMiner
{
    public class ChangedFileRecord
    {
        public string Path { get; set; }
        public int Changeset { get; set; }
        public int WorkItemId { get; set; }
        public string WorkItemTitle { get; set; }
        public string Comment { get; set; }
        public string WorkItemUri { get; set; }
        public string WorkItemType { get; set; }
        public string Module { get; set; }
        public DateTime Date { get; set; }
    }
}