using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace PstToEmlConverter.Core
{
    public static class PstInputResolver
    {
        public static string[] Resolve(string sourcePath, bool isFileMode, bool includeSubfolders, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            sourcePath = sourcePath.Trim();

            if (isFileMode)
            {
                if (!File.Exists(sourcePath)) return Array.Empty<string>();
                if (!sourcePath.EndsWith(".pst", StringComparison.OrdinalIgnoreCase)) return Array.Empty<string>();
                return new[] { Path.GetFullPath(sourcePath) };
            }

            if (!Directory.Exists(sourcePath)) return Array.Empty<string>();

            var opt = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            IEnumerable<string> files = Directory.EnumerateFiles(sourcePath, "*.pst", opt);

            // normalize and stable ordering
            var list = files
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            token.ThrowIfCancellationRequested();
            return list;
        }
    }
}
