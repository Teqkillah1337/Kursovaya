using System.Collections.Generic;

namespace ComputerPassport
{
    public class ComparisonResult
    {
        public string Title { get; set; }
        public List<string> AddedItems { get; set; } = new List<string>();
        public List<string> RemovedItems { get; set; } = new List<string>();
        public List<string> ChangedItems { get; set; } = new List<string>();
    }
}