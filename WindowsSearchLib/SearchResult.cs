namespace WindowsSearch
{
    public struct SearchResult
    {
        public string FilePath { get; private set; }

        public string Content { get; private set; }

        public SearchResult(string filePath, string content)
        {
            FilePath = filePath;
            Content = content;
        }
    }
}