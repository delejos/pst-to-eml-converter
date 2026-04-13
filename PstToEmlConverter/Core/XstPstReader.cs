using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using MimeKit;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using XstReader;
using XstReader.ElementProperties;

namespace PstToEmlConverter.Core
{
    public sealed class XstPstReader : IPstReader
    {
        // Raw MAPI tag IDs for properties not guaranteed to be in the PropertyCanonicalName enum
        private const ushort TagMessageClass = 0x001A; // PidTagMessageClass

        public void ConvertPstToEml(
            string pstPath,
            string outputDir,
            ConversionOptions options,
            IProgress<ConversionProgress> progress,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            Directory.CreateDirectory(outputDir);

            using var xstFile = new XstFile(pstPath);
            var root = xstFile.RootFolder;

            // First pass: count total items so we can report accurate percentages
            int totalItems = CountItems(root, options, token);

            var state = new ProgressState
            {
                TotalItems  = totalItems,
                CurrentPst  = Path.GetFileName(pstPath),
                Progress    = progress,
            };

            ProcessFolder(root, outputDir, outputDir, options, state, token);
        }

        // ── Counting ─────────────────────────────────────────────────────────

        private static int CountItems(XstFolder folder, ConversionOptions options, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            int count = folder.ContentCount;
            foreach (var sub in folder.Folders)
                count += CountItems(sub, options, token);
            return count;
        }

        // ── Folder recursion ─────────────────────────────────────────────────

        private static void ProcessFolder(
            XstFolder folder,
            string rootOutDir,
            string currentOutDir,
            ConversionOptions options,
            ProgressState state,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            string folderOutDir = currentOutDir;
            if (options.PreserveFolderStructure)
            {
                folderOutDir = Path.Combine(currentOutDir, Sanitize(folder.DisplayName));
                Directory.CreateDirectory(folderOutDir);
            }

            state.CurrentFolder = folder.Path ?? folder.DisplayName;
            state.Report();

            foreach (var message in folder.Messages)
            {
                token.ThrowIfCancellationRequested();
                state.ProcessedItems++;

                try
                {
                    ExportMessage(message, folderOutDir, options, state, token);
                }
                catch (Exception ex)
                {
                    state.Failed++;
                    Log(rootOutDir, $"ERROR [{folder.DisplayName}] {message.Subject}: {ex.Message}");
                }

                state.CurrentItem = message.Subject ?? "";
                state.Report();
            }

            foreach (var sub in folder.Folders)
            {
                token.ThrowIfCancellationRequested();
                ProcessFolder(sub, rootOutDir, folderOutDir, options, state, token);
            }
        }

        // ── Message dispatch ─────────────────────────────────────────────────

        private static void ExportMessage(
            XstMessage message,
            string outDir,
            ConversionOptions options,
            ProgressState state,
            CancellationToken token)
        {
            string msgClass = GetMessageClass(message);

            if (msgClass.StartsWith("IPM.Contact", StringComparison.OrdinalIgnoreCase))
            {
                if (!options.ExportContacts) return;
                string? path = UniqueFilePath(outDir, SafeSubject(message.Subject), ".vcf", options.SkipExistingFiles);
                if (path == null) return;
                WriteVCard(message, path);
                state.ContactsSaved++;
            }
            else if (msgClass.StartsWith("IPM.Appointment", StringComparison.OrdinalIgnoreCase) ||
                     msgClass.StartsWith("IPM.Schedule",    StringComparison.OrdinalIgnoreCase))
            {
                if (!options.ExportCalendar) return;
                string? path = UniqueFilePath(outDir, SafeSubject(message.Subject), ".ics", options.SkipExistingFiles);
                if (path == null) return;
                WriteCalendar(message, path, isTask: false);
                state.CalendarSaved++;
            }
            else if (msgClass.StartsWith("IPM.Task", StringComparison.OrdinalIgnoreCase))
            {
                if (!options.ExportTasks) return;
                string? path = UniqueFilePath(outDir, SafeSubject(message.Subject), ".ics", options.SkipExistingFiles);
                if (path == null) return;
                WriteCalendar(message, path, isTask: true);
                state.TasksSaved++;
            }
            else
            {
                // Treat everything else (IPM.Note, IPM.Post, …) as email
                string? path = UniqueFilePath(outDir, SafeSubject(message.Subject), ".eml", options.SkipExistingFiles);
                if (path == null) return;
                WriteEml(message, path);
                state.EmailsSaved++;
            }
        }

        // ── EML via MimeKit ───────────────────────────────────────────────────

        private static void WriteEml(XstMessage message, string path)
        {
            var mime = new MimeMessage();

            // Message-Id
            if (!string.IsNullOrWhiteSpace(message.InternetMessageId))
                mime.MessageId = message.InternetMessageId.Trim('<', '>');

            // Date
            mime.Date = message.Date.HasValue
                ? new DateTimeOffset(message.Date.Value, TimeSpan.Zero)
                : DateTimeOffset.UtcNow;

            // Subject
            mime.Subject = message.Subject ?? "";

            // From
            var senderRecip = message.Recipients[RecipientType.Sender].FirstOrDefault()
                           ?? message.Recipients[RecipientType.SentRepresenting].FirstOrDefault();
            string fromName  = senderRecip?.DisplayName ?? message.From ?? "";
            string fromEmail = senderRecip?.Address ?? "";
            if (!string.IsNullOrWhiteSpace(fromName) || !string.IsNullOrWhiteSpace(fromEmail))
                mime.From.Add(MakeAddress(fromName, fromEmail));

            // To / Cc / Bcc
            foreach (var r in message.Recipients[RecipientType.To])
                mime.To.Add(MakeAddress(r.DisplayName ?? "", r.Address ?? ""));
            foreach (var r in message.Recipients[RecipientType.Cc])
                mime.Cc.Add(MakeAddress(r.DisplayName ?? "", r.Address ?? ""));
            foreach (var r in message.Recipients[RecipientType.Bcc])
                mime.Bcc.Add(MakeAddress(r.DisplayName ?? "", r.Address ?? ""));

            // Build body
            var builder = new BodyBuilder();
            var body = message.Body;
            if (body != null)
            {
                if (body.Format == XstMessageBodyFormat.Html || body.Format == XstMessageBodyFormat.Rtf)
                    builder.HtmlBody = body.Text ?? "";
                else
                    builder.TextBody = body.Text ?? "";
            }

            // Attachments
            foreach (var att in message.Attachments.Where(a => a.IsFile))
            {
                try
                {
                    using var ms = new MemoryStream();
                    att.SaveToStream(ms);
                    ms.Position = 0;
                    builder.Attachments.Add(att.FileNameForSaving ?? "attachment.bin", ms.ToArray());
                }
                catch { /* skip unreadable attachments */ }
            }

            mime.Body = builder.ToMessageBody();

            using var fs = File.Create(path);
            mime.WriteTo(fs);
        }

        // ── vCard (contacts) ──────────────────────────────────────────────────

        private static void WriteVCard(XstMessage message, string path)
        {
            var sb = new StringBuilder();
            sb.AppendLine("BEGIN:VCARD");
            sb.AppendLine("VERSION:3.0");

            string displayName = message.Subject ?? message.DisplayName ?? "";
            sb.AppendLine($"FN:{VCardEscape(displayName)}");

            // Split FN heuristically into given/family name
            var parts = displayName.Split(' ', 2);
            string firstName = parts.Length > 0 ? parts[0] : "";
            string lastName  = parts.Length > 1 ? parts[1] : "";
            sb.AppendLine($"N:{VCardEscape(lastName)};{VCardEscape(firstName)};;;");

            // Try to pull email addresses from named properties (PSETID_Address)
            // We scan all named properties in that property set and look for string values
            // that look like email addresses.
            TryAppendNamedPropsBySet(sb, message,
                new Guid("00062004-0000-0000-C000-000000000046"), "EMAIL;TYPE=INTERNET",
                v => v.Contains('@'));

            // Body → NOTE
            string note = GetBodyText(message);
            if (!string.IsNullOrWhiteSpace(note))
            {
                string escaped = note.Replace("\\", "\\\\")
                                     .Replace("\r\n", "\\n")
                                     .Replace("\n", "\\n")
                                     .Replace(",", "\\,")
                                     .Replace(";", "\\;");
                sb.AppendLine($"NOTE:{escaped}");
            }

            sb.AppendLine($"REV:{DateTime.UtcNow:yyyyMMddTHHmmssZ}");
            sb.AppendLine("END:VCARD");

            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// Scans all named properties in a given PropertySet GUID, appending any string
        /// values that pass the filter as vCard lines (de-duplicated).
        /// </summary>
        private static void TryAppendNamedPropsBySet(
            StringBuilder sb, XstMessage message,
            Guid targetGuid, string vcardProp,
            Func<string, bool> filter)
        {
            try
            {
                var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var prop in message.Properties.Items)
                {
                    if (prop.PropertySet?.Guid() != targetGuid) continue;
                    if (prop.Value is not string val) continue;
                    if (string.IsNullOrWhiteSpace(val)) continue;
                    if (!filter(val)) continue;
                    if (!seen.Add(val)) continue;
                    sb.AppendLine($"{vcardProp}:{VCardEscape(val)}");
                }
            }
            catch { }
        }

        private static string VCardEscape(string s)
            => s.Replace("\\", "\\\\").Replace(",", "\\,").Replace(";", "\\;");

        // ── iCal (calendar / tasks) ───────────────────────────────────────────

        private static void WriteCalendar(XstMessage message, string path, bool isTask)
        {
            var calendar = new Calendar();
            string subject     = message.Subject ?? "";
            string description = GetBodyText(message);
            var date           = message.Date ?? DateTime.UtcNow;

            // Try to find start/end/location via named property scan
            DateTime? namedStart = TryGetDateFromNamedProps(message, new Guid("00062002-0000-0000-C000-000000000046"));
            string location = TryGetStringFromNamedProps(message,
                                  new Guid("00062002-0000-0000-C000-000000000046"),
                                  v => v.Length < 200 && !v.Contains('\n')) ?? "";

            if (isTask)
            {
                var todo = new Todo
                {
                    Summary     = subject,
                    Description = description,
                    DtStamp     = new CalDateTime(DateTime.UtcNow),
                    Start       = new CalDateTime(namedStart ?? date),
                };
                calendar.Todos.Add(todo);
            }
            else
            {
                DateTime start = namedStart ?? date;
                var evt = new CalendarEvent
                {
                    Summary     = subject,
                    Description = description,
                    Location    = location,
                    Start       = new CalDateTime(start),
                    End         = new CalDateTime(start.AddHours(1)),
                    DtStamp     = new CalDateTime(DateTime.UtcNow),
                };
                calendar.Events.Add(evt);
            }

            var serializer = new CalendarSerializer();
            File.WriteAllText(path, serializer.SerializeToString(calendar), Encoding.UTF8);
        }

        private static DateTime? TryGetDateFromNamedProps(XstMessage message, Guid targetGuid)
        {
            try
            {
                foreach (var prop in message.Properties.Items)
                {
                    if (prop.PropertySet?.Guid() != targetGuid) continue;
                    if (prop.Value is DateTime dt && dt > new DateTime(1970, 1, 1))
                        return dt;
                }
            }
            catch { }
            return null;
        }

        private static string? TryGetStringFromNamedProps(XstMessage message, Guid targetGuid, Func<string, bool> filter)
        {
            try
            {
                foreach (var prop in message.Properties.Items)
                {
                    if (prop.PropertySet?.Guid() != targetGuid) continue;
                    if (prop.Value is string s && !string.IsNullOrWhiteSpace(s) && filter(s))
                        return s;
                }
            }
            catch { }
            return null;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string GetMessageClass(XstMessage message)
        {
            try
            {
                return message.Properties[TagMessageClass]?.Value as string ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string GetBodyText(XstMessage message)
        {
            try
            {
                var body = message.Body;
                if (body == null) return "";
                if (body.Format == XstMessageBodyFormat.Html || body.Format == XstMessageBodyFormat.Rtf)
                    return StripHtmlTags(body.Text ?? "");
                return body.Text ?? "";
            }
            catch { return ""; }
        }

        private static string StripHtmlTags(string html)
        {
            if (string.IsNullOrEmpty(html)) return html;
            // Very lightweight strip — good enough for notes/descriptions
            var sb = new StringBuilder(html.Length);
            bool inTag = false;
            foreach (char c in html)
            {
                if (c == '<') { inTag = true; continue; }
                if (c == '>') { inTag = false; continue; }
                if (!inTag) sb.Append(c);
            }
            return sb.ToString().Trim();
        }

        private static MimeKit.MailboxAddress MakeAddress(string name, string email)
        {
            string trimEmail = email.Trim();
            if (!trimEmail.Contains('@'))
                return new MimeKit.MailboxAddress(name.Trim(), $"{name.Trim().Replace(" ", ".")}@unknown");
            return new MimeKit.MailboxAddress(name.Trim(), trimEmail);
        }

        private static string SafeSubject(string? subject)
        {
            string s = subject ?? "item";
            s = Sanitize(s);
            if (s.Length > 80) s = s[..80];
            return s;
        }

        private static string Sanitize(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            name = name.Trim();
            return string.IsNullOrWhiteSpace(name) ? "item" : name;
        }

        /// <summary>
        /// Returns a unique file path, or null if SkipExistingFiles is true and path exists.
        /// </summary>
        private static string? UniqueFilePath(string dir, string baseName, string ext, bool skipExisting)
        {
            Directory.CreateDirectory(dir);
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
            string candidate = Path.Combine(dir, $"{stamp}_{baseName}{ext}");

            if (skipExisting && File.Exists(candidate))
                return null;

            return candidate;
        }

        private static void Log(string outDir, string line)
        {
            try
            {
                File.AppendAllText(
                    Path.Combine(outDir, "conversion_log.txt"),
                    $"[{DateTime.Now:HH:mm:ss}] {line}{Environment.NewLine}");
            }
            catch { }
        }

        // ── Progress state ────────────────────────────────────────────────────

        private sealed class ProgressState
        {
            public string CurrentPst   { get; set; } = "";
            public string CurrentFolder{ get; set; } = "";
            public string CurrentItem  { get; set; } = "";
            public int TotalItems      { get; set; }
            public int ProcessedItems  { get; set; }
            public int EmailsSaved     { get; set; }
            public int ContactsSaved   { get; set; }
            public int CalendarSaved   { get; set; }
            public int TasksSaved      { get; set; }
            public int Failed          { get; set; }
            public IProgress<ConversionProgress> Progress { get; set; } = null!;

            public void Report() => Progress?.Report(new ConversionProgress
            {
                CurrentPst     = CurrentPst,
                CurrentFolder  = CurrentFolder,
                CurrentItem    = CurrentItem,
                TotalItems     = TotalItems,
                ProcessedItems = ProcessedItems,
                EmailsSaved    = EmailsSaved,
                ContactsSaved  = ContactsSaved,
                CalendarSaved  = CalendarSaved,
                TasksSaved     = TasksSaved,
                Failed         = Failed,
            });
        }
    }
}
