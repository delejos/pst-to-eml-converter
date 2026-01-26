using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace PstToEmlConverter.Core
{
    public sealed class OutlookPstReader : IPstReader
    {
        public void ConvertPstToEml(string pstPath, string outputDir, ConversionOptions options, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            Directory.CreateDirectory(outputDir);

            Outlook.Application app = null!;
            Outlook.NameSpace ns = null!;
            Outlook.MAPIFolder storeRoot = null!;

            try
            {
                app = new Outlook.Application();
                ns = app.GetNamespace("MAPI");
                ns.Logon(Type.Missing, Type.Missing, false, false);

                ns.AddStore(pstPath);

                storeRoot = FindStoreRootFolder(ns, pstPath);

                // rootOutDir stays constant; currentOutDir changes as we recurse
                ExportFolderRecursive(ns, storeRoot, outputDir, outputDir, options, token);
            }
            finally
            {
                try
                {
                    if (storeRoot != null)
                        ns.RemoveStore(storeRoot);
                }
                catch { }

                ReleaseCom(storeRoot);
                ReleaseCom(ns);
                ReleaseCom(app);
            }
        }

        private static Outlook.MAPIFolder FindStoreRootFolder(Outlook.NameSpace ns, string pstPath)
        {
            Outlook.Stores stores = null!;
            try
            {
                stores = ns.Stores;
                string target = Path.GetFullPath(pstPath);

                for (int i = 1; i <= stores.Count; i++)
                {
                    Outlook.Store store = null!;
                    try
                    {
                        store = stores[i];
                        string filePath = store.FilePath;

                        if (!string.IsNullOrEmpty(filePath) &&
                            string.Equals(Path.GetFullPath(filePath), target, StringComparison.OrdinalIgnoreCase))
                        {
                            // root folder for this PST store
                            return store.GetRootFolder();
                        }
                    }
                    finally
                    {
                        ReleaseCom(store);
                    }
                }

                throw new InvalidOperationException("PST store not found in NameSpace.Stores after AddStore().");
            }
            finally
            {
                ReleaseCom(stores);
            }
        }

        private static void ExportFolderRecursive(
            Outlook.NameSpace ns,
            Outlook.MAPIFolder folder,
            string rootOutDir,
            string currentOutDir,
            ConversionOptions options,
            CancellationToken token)
        {
            // Always log to the SAME debug file (rootOutDir)
            DebugLine(rootOutDir, $"FOLDER: {folder.FolderPath}");
            token.ThrowIfCancellationRequested();

            // Where THIS folder's files go
            string folderOutDir = currentOutDir;

            if (options.PreserveFolderStructure)
            {
                folderOutDir = Path.Combine(currentOutDir, SanitizeFolderName(folder.Name));
                Directory.CreateDirectory(folderOutDir);
            }

            // 1) Export items in this folder (using Items.GetFirst/GetNext)
            int seen = 0, saved = 0, failed = 0;

            Outlook.Items items = null!;
            try
            {
                items = folder.Items;

                object cur = null!;
                cur = items.GetFirst();

                while (cur != null)
                {
                    token.ThrowIfCancellationRequested();
                    seen++;

                    object next = null!;
                    try
                    {
                        // get next before processing current (safer for COM enumerations)
                        next = items.GetNext();

                        try
                        {
                            Directory.CreateDirectory(folderOutDir);

                            // subject (best-effort)
                            string subject = "item";
                            try
                            {
                                dynamic dd = cur;
                                subject = dd.Subject as string ?? "item";
                            }
                            catch { }

                            subject = SanitizeFolderName(subject);
                            if (subject.Length > 80) subject = subject.Substring(0, 80);

                            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
                            string path = Path.Combine(folderOutDir, $"{stamp}_{subject}.eml");

                            dynamic itemDyn = cur;

                            // 1) Save to MSG (works across Outlook versions)
                            string msgPath = Path.Combine(folderOutDir, $"{stamp}_{subject}.msg");
                            // Try cast to MailItem for rich fields
                            var mail = cur as Outlook.MailItem;
                            if (mail != null)
                            if (mail != null)
                            {
                                string emlPath = Path.Combine(folderOutDir, $"{stamp}_{subject}.eml");
                                WriteEmlFromMailItem(mail, emlPath);
                                saved++;
                            }
                            // 2) Re-open and save as EML (102 may work on reopened MailItem)
                            dynamic reopened = ns.OpenSharedItem(msgPath);
                            try
                            {
                                // overwrite EML path
                                reopened.SaveAs(path, 102); // 102 = RFC822/EML
                            }
                            finally
                            {
                                ReleaseCom(reopened);
                            }

                            saved++;
                        }
                        catch (Exception ex)
                        {
                            failed++;
                            //DebugLine(rootOutDir, $"  Export failed in {folder.FolderPath}: {ex.GetType().Name}: {ex.Message}");
                        }
                    }
                    finally
                    {
                        ReleaseCom(cur);
                    }

                    cur = next;
                }
            }
            catch (Exception ex)
            {
                DebugLine(rootOutDir, $"  ERROR reading Items in {folder.FolderPath}: {ex.GetType().Name}: {ex.Message}");
            }
            finally
            {
                ReleaseCom(items);
            }

            DebugLine(rootOutDir, $"  RESULT: seen={seen}, saved={saved}, failed={failed}");

            // 2) Recurse into subfolders
            Outlook.Folders subFolders = null!;
            try
            {
                subFolders = folder.Folders;

                int subCount = 0;
                try { subCount = subFolders.Count; } catch { }
                DebugLine(rootOutDir, $"  Subfolders.Count = {subCount}");

                for (int i = 1; i <= subFolders.Count; i++)
                {
                    token.ThrowIfCancellationRequested();

                    Outlook.MAPIFolder sub = null!;
                    try
                    {
                        sub = subFolders[i];
                        DebugLine(rootOutDir, $"  -> Subfolder: {sub.Name}");

                        // IMPORTANT: recurse into *sub*, and pass folderOutDir as the new currentOutDir
                        ExportFolderRecursive(ns, sub, rootOutDir, folderOutDir, options, token);
                    }
                    catch (Exception ex)
                    {
                        DebugLine(rootOutDir, $"  ERROR recursing into subfolder[{i}] of {folder.FolderPath}: {ex.GetType().Name}: {ex.Message}");
                    }
                    finally
                    {
                        ReleaseCom(sub);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLine(rootOutDir, $"  ERROR enumerating subfolders in {folder.FolderPath}: {ex.GetType().Name}: {ex.Message}");
            }
            finally
            {
                ReleaseCom(subFolders);
            }
        }

        private static void WriteEmlFromMailItem(Outlook.MailItem mail, string emlPath)
        {
            // Prefer HTML if present, otherwise plain text
            bool hasHtml = false;
            string html = "";
            try
            {
                html = mail.HTMLBody ?? "";
                hasHtml = !string.IsNullOrWhiteSpace(html);
            }
            catch { }

            string text = "";
            try { text = mail.Body ?? ""; } catch { }

            // Headers
            string from = SafeGet(() => mail.SenderEmailAddress) ?? SafeGet(() => mail.SenderName) ?? "";
            string to = SafeGet(() => mail.To) ?? "";
            string cc = SafeGet(() => mail.CC) ?? "";
            string subject = SafeGet(() => mail.Subject) ?? "";
            DateTime sent = SafeGet(() => mail.SentOn) ?? DateTime.Now;

            // Basic RFC 5322 date format
            string dateHeader = sent.ToString("ddd, dd MMM yyyy HH:mm:ss zzz");

            // Attachments?
            int attCount = 0;
            try { attCount = mail.Attachments?.Count ?? 0; } catch { }

            var sb = new System.Text.StringBuilder();

            sb.AppendLine($"Date: {dateHeader}");
            if (!string.IsNullOrWhiteSpace(from)) sb.AppendLine($"From: {EncodeHeader(from)}");
            if (!string.IsNullOrWhiteSpace(to)) sb.AppendLine($"To: {EncodeHeader(to)}");
            if (!string.IsNullOrWhiteSpace(cc)) sb.AppendLine($"Cc: {EncodeHeader(cc)}");
            sb.AppendLine($"Subject: {EncodeHeader(subject)}");
            sb.AppendLine("MIME-Version: 1.0");

            if (attCount <= 0)
            {
                // Single-part
                if (hasHtml)
                {
                    sb.AppendLine(@"Content-Type: text/html; charset=""utf-8""");
                    sb.AppendLine("Content-Transfer-Encoding: quoted-printable");
                    sb.AppendLine();
                    sb.Append(QuotedPrintableEncode(html));
                }
                else
                {
                    sb.AppendLine(@"Content-Type: text/plain; charset=""utf-8""");
                    sb.AppendLine("Content-Transfer-Encoding: quoted-printable");
                    sb.AppendLine();
                    sb.Append(QuotedPrintableEncode(text));
                }

                File.WriteAllText(emlPath, sb.ToString(), System.Text.Encoding.UTF8);
                return;
            }

            // Multipart/mixed with attachments
            string boundary = "====boundary_" + Guid.NewGuid().ToString("N");
            sb.AppendLine($@"Content-Type: multipart/mixed; boundary=""{boundary}""");
            sb.AppendLine();
            sb.AppendLine($"--{boundary}");

            if (hasHtml)
            {
                sb.AppendLine(@"Content-Type: text/html; charset=""utf-8""");
                sb.AppendLine("Content-Transfer-Encoding: quoted-printable");
                sb.AppendLine();
                sb.AppendLine(QuotedPrintableEncode(html));
            }
            else
            {
                sb.AppendLine(@"Content-Type: text/plain; charset=""utf-8""");
                sb.AppendLine("Content-Transfer-Encoding: quoted-printable");
                sb.AppendLine();
                sb.AppendLine(QuotedPrintableEncode(text));
            }

            // Attachments
            Outlook.Attachments atts = null!;
            try
            {
                atts = mail.Attachments;
                for (int i = 1; i <= atts.Count; i++)
                {
                    Outlook.Attachment att = null!;
                    try
                    {
                        att = atts[i];

                        string originalName = att.FileName ?? $"attachment{i}";
                        string safeName = SanitizeFileName(originalName);

                        // Short temp path
                        string temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".bin");

                        try
                        {
                            att.SaveAsFile(temp);
                        }
                        catch (Exception ex)
                        {
                            // Attachment couldn't be saved by Outlook -> skip it
                            // IMPORTANT: do not throw
                            // Optional: you can log it:
                            // DebugLine(rootOutDir, $"  Attachment skipped ({safeName}): {ex.Message}");
                            continue;
                        }

                        byte[] bytes;
                        try
                        {
                            bytes = File.ReadAllBytes(temp);
                        }
                        finally
                        {
                            try { File.Delete(temp); } catch { }
                        }

                        sb.AppendLine($"--{boundary}");
                        sb.AppendLine($@"Content-Type: application/octet-stream; name=""{EscapeQuotes(safeName)}""");
                        sb.AppendLine("Content-Transfer-Encoding: base64");
                        sb.AppendLine($@"Content-Disposition: attachment; filename=""{EscapeQuotes(safeName)}""");
                        sb.AppendLine();
                        sb.AppendLine(Base64WithLineBreaks(bytes));
                    }
                    finally
                    {
                        ReleaseCom(att);
                    }
                }
            }
            finally
            {
                ReleaseCom(atts);
            }

            sb.AppendLine($"--{boundary}--");

            File.WriteAllText(emlPath, sb.ToString(), System.Text.Encoding.UTF8);
        }

        private static string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');

            name = name.Trim();
            if (string.IsNullOrWhiteSpace(name)) name = "attachment.bin";

            // keep it short
            if (name.Length > 120) name = name.Substring(0, 120);
            return name;
        }

        private static T? SafeGet<T>(Func<T> getter) where T : class
        {
            try { return getter(); } catch { return null; }
        }

        private static DateTime? SafeGet(Func<DateTime> getter)
        {
            try { return getter(); } catch { return null; }
        }

        private static string EscapeQuotes(string s) => s.Replace("\"", "'");

        private static string EncodeHeader(string value)
        {
            // Minimal: if non-ascii, encode as UTF-8 Base64 (RFC 2047)
            if (string.IsNullOrEmpty(value)) return "";
            bool nonAscii = false;
            foreach (char c in value) { if (c > 127) { nonAscii = true; break; } }
            if (!nonAscii) return value;

            var bytes = System.Text.Encoding.UTF8.GetBytes(value);
            return $"=?utf-8?B?{Convert.ToBase64String(bytes)}?=";
        }

        private static string QuotedPrintableEncode(string input)
        {
            if (input == null) return "";
            var bytes = System.Text.Encoding.UTF8.GetBytes(input);
            var sb = new System.Text.StringBuilder();

            int lineLen = 0;
            foreach (byte b in bytes)
            {
                // CRLF handling
                if (b == 0x0D) continue; // skip CR, we'll handle LF
                if (b == 0x0A)
                {
                    sb.Append("\r\n");
                    lineLen = 0;
                    continue;
                }

                string chunk;
                if ((b >= 33 && b <= 60) || (b >= 62 && b <= 126))
                {
                    chunk = ((char)b).ToString();
                }
                else if (b == 0x20 || b == 0x09)
                {
                    chunk = ((char)b).ToString();
                }
                else
                {
                    chunk = "=" + b.ToString("X2");
                }

                // soft line breaks at ~75 chars
                if (lineLen + chunk.Length > 75)
                {
                    sb.Append("=\r\n");
                    lineLen = 0;
                }

                sb.Append(chunk);
                lineLen += chunk.Length;
            }

            return sb.ToString();
        }

        private static string Base64WithLineBreaks(byte[] bytes)
        {
            string b64 = Convert.ToBase64String(bytes);
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < b64.Length; i += 76)
            {
                int len = Math.Min(76, b64.Length - i);
                sb.AppendLine(b64.Substring(i, len));
            }
            return sb.ToString();
        }


        private static string SanitizeFolderName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');

            name = name.Trim();
            return string.IsNullOrWhiteSpace(name) ? "Folder" : name;
        }

        private static void ReleaseCom(object? o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                    Marshal.FinalReleaseComObject(o);
            }
            catch { }
        }

        private static void DebugLine(string outDir, string line)
        {
            try
            {
                File.AppendAllText(Path.Combine(outDir, "export_debug.txt"),
                    $"[{DateTime.Now:HH:mm:ss}] {line}{Environment.NewLine}");
            }
            catch { }
        }
    }
}
