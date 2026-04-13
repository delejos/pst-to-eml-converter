using System;
using System.IO;
using System.Threading;

namespace PstToEmlConverter.Core
{
    public sealed class DryRunPstReader : IPstReader
    {
        public void ConvertPstToEml(
            string pstPath,
            string outputDir,
            ConversionOptions options,
            IProgress<ConversionProgress> progress,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            Directory.CreateDirectory(outputDir);

            progress?.Report(new ConversionProgress
            {
                CurrentPst    = Path.GetFileName(pstPath),
                CurrentFolder = "(dry run)",
                TotalItems    = 1,
                ProcessedItems = 1,
            });

            File.WriteAllText(
                Path.Combine(outputDir, Path.GetFileName(pstPath) + ".dryrun.txt"),
                $"Dry-run: PST={pstPath}\nOutput={outputDir}\n" +
                $"Contacts={options.ExportContacts}\nCalendar={options.ExportCalendar}\nTasks={options.ExportTasks}\n");
        }
    }
}
