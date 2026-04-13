using Microsoft.Win32;
using PstToEmlConverter.Core;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace PstToEmlConverter
{
    public partial class MainWindow : Window
    {
        private CancellationTokenSource? _cts;
        private readonly IPstReader _reader = new XstPstReader();

        public MainWindow()
        {
            InitializeComponent();
            Loaded += (_, __) =>
            {
                AppendLog("Ready. No Outlook installation required.");
                ValidateInputs();
            };
        }

        // ── Source/Dest radio + browse ────────────────────────────────────────

        private void RbFile_Checked(object sender, RoutedEventArgs e)
        {
            AppendLog("Source mode: Single PST file");
            ValidateInputs();
        }

        private void RbFolder_Checked(object sender, RoutedEventArgs e)
        {
            AppendLog("Source mode: Folder with PST files");
            ValidateInputs();
        }

        private void AnyPathChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
            => ValidateInputs();

        private void BtnBrowseSource_Click(object sender, RoutedEventArgs e)
        {
            if (RbFile.IsChecked == true)
            {
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Select a PST file",
                    Filter = "Outlook Data File (*.pst)|*.pst|All files (*.*)|*.*",
                    CheckFileExists = true,
                };
                if (dlg.ShowDialog() == true)
                    TxtSource.Text = dlg.FileName;
            }
            else
            {
                using var dlg = new System.Windows.Forms.FolderBrowserDialog
                {
                    Description = "Select a folder that contains PST files",
                    UseDescriptionForTitle = true,
                };
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    TxtSource.Text = dlg.SelectedPath;
            }
            ValidateInputs();
        }

        private void BtnBrowseDest_Click(object sender, RoutedEventArgs e)
        {
            using var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select destination folder for output",
                UseDescriptionForTitle = true,
            };
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                TxtDest.Text = dlg.SelectedPath;
            ValidateInputs();
        }

        // ── Log ───────────────────────────────────────────────────────────────

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            TxtLog.Clear();
            AppendLog("Log cleared.");
        }

        private void AppendLog(string message)
        {
            if (TxtLog == null) return;

            if (Dispatcher.CheckAccess())
            {
                TxtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
                TxtLog.ScrollToEnd();
                return;
            }

            Dispatcher.BeginInvoke(() =>
            {
                TxtLog?.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
                TxtLog?.ScrollToEnd();
            });
        }

        // ── Start / Cancel ────────────────────────────────────────────────────

        private async void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            ValidateInputs();
            if (!BtnStart.IsEnabled) return;

            // Capture UI state on the UI thread before going async
            bool   isFileMode         = RbFile.IsChecked == true;
            string source             = TxtSource.Text.Trim();
            string dest               = TxtDest.Text.Trim();

            var options = new ConversionOptions
            {
                IncludeSubfolders    = ChkIncludeSubfolders.IsChecked == true,
                PreserveFolderStructure = ChkPreserveStructure.IsChecked == true,
                SkipExistingFiles    = ChkSkipExisting.IsChecked == true,
                ExportContacts       = ChkExportContacts.IsChecked == true,
                ExportCalendar       = ChkExportCalendar.IsChecked == true,
                ExportTasks          = ChkExportTasks.IsChecked == true,
            };

            BtnStart.IsEnabled  = false;
            BtnCancel.IsEnabled = true;
            Progress.Value      = 0;
            TxtCounts.Visibility = Visibility.Collapsed;
            _cts = new CancellationTokenSource();

            var progressHandler = new Progress<ConversionProgress>(OnProgress);

            try
            {
                AppendLog("Starting…");

                await Task.Run(() =>
                {
                    var pstFiles = PstInputResolver.Resolve(
                        source, isFileMode, options.IncludeSubfolders, _cts.Token);

                    int total = pstFiles.Length;
                    Dispatcher.Invoke(() =>
                        AppendLog($"Found {total} PST file(s). Destination: {dest}"));

                    for (int i = 0; i < total; i++)
                    {
                        _cts.Token.ThrowIfCancellationRequested();
                        string pst = pstFiles[i];
                        string pstOut = Path.Combine(dest,
                            SanitizeFolderName(Path.GetFileNameWithoutExtension(pst)));

                        Dispatcher.Invoke(() =>
                            AppendLog($"[{i + 1}/{total}] {Path.GetFileName(pst)}"));

                        try
                        {
                            _reader.ConvertPstToEml(pst, pstOut, options, progressHandler, _cts.Token);
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex)
                        {
                            Dispatcher.Invoke(() =>
                                AppendLog($"ERROR: {Path.GetFileName(pst)}: {ex.Message}"));
                        }
                    }
                }, _cts.Token);

                AppendLog("Done.");
                Progress.Value = 100;
                TxtStatus.Text = "Conversion complete.";
            }
            catch (OperationCanceledException)
            {
                AppendLog("Cancelled.");
                TxtStatus.Text = "Cancelled.";
            }
            catch (Exception ex)
            {
                AppendLog($"ERROR: {ex.Message}");
                TxtStatus.Text = "Error — see log.";
            }
            finally
            {
                BtnCancel.IsEnabled = false;
                BtnStart.IsEnabled  = true;
                _cts?.Dispose();
                _cts = null;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
            AppendLog("Cancelling…");
        }

        // ── Progress callback ─────────────────────────────────────────────────

        private void OnProgress(ConversionProgress p)
        {
            // Already marshalled to UI thread by Progress<T>
            Progress.Value = p.Percentage;

            TxtStatus.Text = string.IsNullOrEmpty(p.CurrentFolder)
                ? p.CurrentPst
                : $"{p.CurrentPst}  ›  {p.CurrentFolder}";

            if (p.TotalItems > 0)
            {
                TxtCounts.Text = $"Emails: {p.EmailsSaved}  " +
                                 $"Contacts: {p.ContactsSaved}  " +
                                 $"Calendar: {p.CalendarSaved}  " +
                                 $"Tasks: {p.TasksSaved}  " +
                                 $"Failed: {p.Failed}  " +
                                 $"({p.ProcessedItems}/{p.TotalItems})";
                TxtCounts.Visibility = Visibility.Visible;

                if (!string.IsNullOrEmpty(p.CurrentItem))
                    AppendLog($"  {p.CurrentFolder}  ›  {p.CurrentItem}");
            }
        }

        // ── Validation ────────────────────────────────────────────────────────

        private void ValidateInputs()
        {
            if (TxtSource == null || TxtDest == null || BtnStart == null || RbFile == null)
                return;

            string source = TxtSource.Text.Trim();
            string dest   = TxtDest.Text.Trim();

            bool sourceOk = RbFile.IsChecked == true
                ? File.Exists(source) && source.EndsWith(".pst", StringComparison.OrdinalIgnoreCase)
                : Directory.Exists(source);

            bool destOk = Directory.Exists(dest);

            BtnStart.IsEnabled = sourceOk && destOk;
        }

        private static string SanitizeFolderName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            name = name.Trim();
            return string.IsNullOrWhiteSpace(name) ? "PST" : name;
        }
    }
}
