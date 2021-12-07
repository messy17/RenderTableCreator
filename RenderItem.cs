namespace RenderTableCreator
{
    internal class RenderItem
    {
        public string ImageName { get; set; }
        public string Description { get; set; }
        public int LineNumber { get; set; }
        public int RefCount { get; set; } // keeps track of the number of instances 
        

        public RenderItem(string _imageName, string _description, int lineNumber)
        {
            ImageName = _imageName;
            Description = _description;
            LineNumber = lineNumber;
            RefCount = 1; 
        }
    }
}
