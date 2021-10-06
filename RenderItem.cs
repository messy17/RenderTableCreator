namespace RenderTableCreator
{
    internal class RenderItem
    {
        public string ImageName { get; set; }
        public string Description { get; set; }
        public int LineNumber { get; set; }

        public RenderItem(string _imageName, string _description, int lineNumber)
        {
            ImageName = _imageName;
            Description = _description;
            LineNumber = lineNumber;
        }
    }
}
