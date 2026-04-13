using System;

namespace PstToEmlConverter.Core
{
    public sealed class ConversionProgress
    {
        public string CurrentPst { get; init; } = "";
        public string CurrentFolder { get; init; } = "";
        public string CurrentItem { get; init; } = "";
        public int TotalItems { get; init; }
        public int ProcessedItems { get; init; }
        public int EmailsSaved { get; init; }
        public int ContactsSaved { get; init; }
        public int CalendarSaved { get; init; }
        public int TasksSaved { get; init; }
        public int Failed { get; init; }

        public double Percentage => TotalItems > 0
            ? Math.Clamp(ProcessedItems * 100.0 / TotalItems, 0, 100)
            : 0;

        public int TotalSaved => EmailsSaved + ContactsSaved + CalendarSaved + TasksSaved;
    }
}
