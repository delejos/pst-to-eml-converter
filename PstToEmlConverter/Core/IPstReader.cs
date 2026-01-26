using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace PstToEmlConverter.Core
{
    public interface IPstReader
    {
        // Converts one PST file to EML files in outputDir.
        // Implementations should throw on fatal errors and use progress reporting separately.
        void ConvertPstToEml(string pstPath, string outputDir, ConversionOptions options, CancellationToken token);
    }
}