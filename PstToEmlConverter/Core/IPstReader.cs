using System;
using System.Threading;

namespace PstToEmlConverter.Core
{
    public interface IPstReader
    {
        void ConvertPstToEml(
            string pstPath,
            string outputDir,
            ConversionOptions options,
            IProgress<ConversionProgress> progress,
            CancellationToken token);
    }
}
