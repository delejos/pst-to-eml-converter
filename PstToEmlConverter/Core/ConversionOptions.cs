using System;
using System.Collections.Generic;
using System.Text;

namespace PstToEmlConverter.Core
{
    public sealed class ConversionOptions
    {
        public bool IncludeSubfolders { get; set; } = true;
        public bool PreserveFolderStructure { get; set; } = true;
        public bool SkipExistingEml { get; set; } = false;
    }
}
