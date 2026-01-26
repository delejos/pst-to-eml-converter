using Microsoft.Win32;
using PstToEmlConverter.Core;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace PstToEmlConverter
{
    public partial class MainWindow : Window
    {

        private DispatcherTimer? _panicTimer;
        private CancellationTokenSource? _cts;
        private readonly System.Collections.Generic.List<string> _pendingLogs = new();


        private readonly IPstReader _reader = new PstToEmlConverter.Core.OutlookPstReader();

        public MainWindow()
        {
            InitializeComponent();

            Loaded += (_, __) =>
            {
                AppendLog("Ready.");
                ValidateInputs();
            };
        }


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
        {
            ValidateInputs();
        }

        private void BtnBrowseSource_Click(object sender, RoutedEventArgs e)
        {
            if (RbFile.IsChecked == true)
            {
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Select a PST file",
                    Filter = "Outlook Data File (*.pst)|*.pst|All files (*.*)|*.*",
                    CheckFileExists = true
                };

                if (dlg.ShowDialog() == true)
                    TxtSource.Text = dlg.FileName;
            }
            else
            {
                using var dlg = new System.Windows.Forms.FolderBrowserDialog
                {
                    Description = "Select a folder that contains PST files",
                    UseDescriptionForTitle = true
                };

                var result = dlg.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                    TxtSource.Text = dlg.SelectedPath;
            }

            ValidateInputs(); // ✅ force refresh
        }


        private void BtnBrowseDest_Click(object sender, RoutedEventArgs e)
        {
            using var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select destination folder for EML output",
                UseDescriptionForTitle = true
            };

            var result = dlg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                TxtDest.Text = dlg.SelectedPath;

            ValidateInputs(); // ✅ force refresh
        }


        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            TxtLog.Clear();
            AppendLog("Log cleared.");
        }

        private void StartPanicTimer()
        {
            // Ensure old timer is gone
            _panicTimer?.Stop();
            _panicTimer = null;

            _panicTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(10)
            };

            _panicTimer.Tick += (s, e) =>
            {
                _panicTimer?.Stop();
                _panicTimer = null;

                AppendLog("ℹ️ Do not panic if it looks stuck — Outlook can be slow on large folders.");
            };

            _panicTimer.Start();
        }


        private async void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            ValidateInputs();
            StartPanicTimer();
            if (!BtnStart.IsEnabled) return;

            BtnStart.IsEnabled = false;
            BtnCancel.IsEnabled = true;
            Progress.Value = 0;
            _cts = new CancellationTokenSource();

            // ✅ Capture UI values on UI thread BEFORE Task.Run
            bool isFileMode = RbFile.IsChecked == true;
            string source = TxtSource.Text.Trim();
            string dest = TxtDest.Text.Trim();

            var options = new ConversionOptions
            {
                IncludeSubfolders = true,
                PreserveFolderStructure = true,
                SkipExistingEml = false
            };

            try
            {
                AppendLog("Starting…");

                await RunStaAsync(() => RunJob(isFileMode, source, dest, options, _cts.Token), _cts.Token);

                AppendLog("Done.");
                Progress.Value = 100;
            }
            catch (OperationCanceledException)
            {
                AppendLog("Cancelled.");
            }
            catch (Exception ex)
            {
                AppendLog("ERROR: " + ex.Message);
            }
            finally
            {
                StopPanicTimer();
                BtnCancel.IsEnabled = false;
                BtnStart.IsEnabled = true;
                _cts.Dispose();
                _cts = null;
            }
        }

        private void StopPanicTimer()
        {
            if (_panicTimer != null)
            {
                _panicTimer.Stop();
                _panicTimer = null;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
            AppendLog("Cancel requested…");
        }

        private void RunJob(
            bool isFileMode,
            string source,
            string dest,
            PstToEmlConverter.Core.ConversionOptions options,
            CancellationToken token)
                {

                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    int ok = 0, failed = 0;

                    var pstFiles = PstToEmlConverter.Core.PstInputResolver.Resolve(
                        source,
                        isFileMode,
                        options.IncludeSubfolders,
                        token);

                    int total = pstFiles.Length;
                    AppendLog($"Found {total} PST file(s). Destination: {dest}");

                    for (int i = 0; i < total; i++)
                    {
                        token.ThrowIfCancellationRequested();

                        string pst = pstFiles[i];
                        string pstBaseName = Path.GetFileNameWithoutExtension(pst);
                        string pstOutDir = Path.Combine(dest, SanitizeFolderName(pstBaseName));

                        AppendLog($"[{i + 1}/{total}] Processing: {pst}");

                        try
                        {
                            _reader.ConvertPstToEml(pst, pstOutDir, options, token);
                            ok++;
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex)
                        {
                            failed++;

                            var msg = ex.ToString();
                            AppendLog($"ERROR converting {pst}: {msg}");

                            // Also write to a file so it isn't cut off
                            try
                            {
                                var crashPath = System.IO.Path.Combine(dest, "outlook_error.txt");
                                System.IO.File.WriteAllText(crashPath, msg);
                                AppendLog($"Wrote details to: {crashPath}");
                            }
                            catch { }
                        }

                        int pct = (int)Math.Round((i + 1) * 100.0 / Math.Max(1, total));
                        Dispatcher.Invoke(() => Progress.Value = pct);
                    }

                    sw.Stop();
                    AppendLog($"Summary: {ok} succeeded, {failed} failed, {total} total. Time: {sw.Elapsed}");
        }

        private static string SanitizeFolderName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');

            name = name.Trim();
            return string.IsNullOrWhiteSpace(name) ? "PST" : name;
        }

        private void ValidateInputs()
        {
            if (TxtSource == null || TxtDest == null || BtnStart == null || RbFile == null)
                return;

            string source = TxtSource.Text.Trim();
            string dest = TxtDest.Text.Trim();

            bool sourceOk;
            if (RbFile.IsChecked == true)
                sourceOk = System.IO.File.Exists(source) && source.EndsWith(".pst", StringComparison.OrdinalIgnoreCase);
            else
                sourceOk = System.IO.Directory.Exists(source);

            bool destOk = System.IO.Directory.Exists(dest);

            BtnStart.IsEnabled = sourceOk && destOk;
        }


        private void AppendLog(string message)
        {
            // During startup, controls may not exist yet
            if (TxtLog == null) return;

            if (Dispatcher.CheckAccess())
            {
                TxtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
                TxtLog.ScrollToEnd();
                return;
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (TxtLog == null) return;
                TxtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
                TxtLog.ScrollToEnd();
            }));
        }

        private Task RunStaAsync(Action action, CancellationToken token)
        {
            var tcs = new TaskCompletionSource<object?>();

            var thread = new Thread(() =>
            {
                try
                {
                    token.ThrowIfCancellationRequested();
                    action();
                    tcs.TrySetResult(null);
                }
                catch (OperationCanceledException oce)
                {
                    tcs.TrySetCanceled(oce.CancellationToken);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            });

            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }



    }
}
