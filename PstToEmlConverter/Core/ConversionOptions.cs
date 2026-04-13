namespace PstToEmlConverter.Core
{
    public sealed class ConversionOptions
    {
        public bool IncludeSubfolders { get; set; } = true;
        public bool PreserveFolderStructure { get; set; } = true;
        public bool SkipExistingFiles { get; set; } = false;
        public bool ExportContacts { get; set; } = true;
        public bool ExportCalendar { get; set; } = true;
        public bool ExportTasks { get; set; } = true;
    }
}
