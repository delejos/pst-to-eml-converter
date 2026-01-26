using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Threading;

namespace PstToEmlConverter.Core
{
    public sealed class DryRunPstReader : IPstReader
    {
        public void ConvertPstToEml(string pstPath, string outputDir, ConversionOptions options, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            // Just create a marker file so we can confirm output wiring works.
            Directory.CreateDirectory(outputDir);
            var marker = Path.Combine(outputDir, Path.GetFileName(pstPath) + ".dryrun.txt");
            File.WriteAllText(marker,
            $@"Dry-run: PST conversion not implemented yet.
            PST: {pstPath}
            Output: {outputDir}
            IncludeSubfolders: {options.IncludeSubfolders}
            PreserveFolderStructure: {options.PreserveFolderStructure}
            SkipExistingEml: {options.SkipExistingEml}
            ");
        }
    }
}