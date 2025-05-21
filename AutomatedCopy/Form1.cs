using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace AutomatedCopy
{
    public partial class Form1 : Form
    {
        private TextBox txtExcelPath;
        private Button btnBrowse;
        private Button btnCopyFiles;
        private Button btnPauseResume;
        private Button btnStop;
        private RadioButton rbtnCopy;
        private RadioButton rbtnMove;
        private Label overalProgressbarLabel;
        private ProgressBar progressBarOverall;
        private Label currentProgressbarLabel;
        private ProgressBar progressBarCurrent;
        private RichTextBox rtbLogs;
        private CheckBox chkAutoScroll;
        private TreeView tvSourceTree;
        private Label lblFileCount;
        private Label lblFolderCount;
        private Label lblTotalSize;
        private bool isMoveOperation = false;
        private bool isPaused = false;
        private bool isStopped = false;
        private CancellationTokenSource cancellationTokenSource;
        //private CancellationTokenSource _fileOperationCts;
        //private CancellationTokenSource _scanCts;
        private bool _isScanning = false;

        private List<string> successfulFiles = new List<string>();
        private List<string> failedFiles = new List<string>();
        //private long totalSizeBytes = 0;
        private long processedSizeBytes = 0;
        //private int totalFiles = 0;
        private int processedFiles = 0;
        private int totalOperations = 0;
        private int completedOperations = 0;
        private object progressLock = new object();
        private bool overwriteAll = false;
        private bool skipAll = false;

        public Form1()
        {
            InitializeComponent();
            this.Text = "Automated Copy V4.1";
            this.Name = "Automated Copy";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //_fileOperationCts = new CancellationTokenSource();
            //_scanCts = new CancellationTokenSource();

            InitializeControls();
        }

        private void InitializeControls()
        {
            this.Size = new Size(1000, 700);

            txtExcelPath = new TextBox
            {
                Location = new Point(10, 10),
                Size = new Size(450, 20)
            };
            this.Controls.Add(txtExcelPath);

            btnBrowse = new Button
            {
                Location = new Point(470, 10),
                Size = new Size(100, 23),
                Text = "Browse"
            };
            btnBrowse.Click += BtnBrowse_Click;
            this.Controls.Add(btnBrowse);

            btnCopyFiles = new Button
            {
                Location = new Point(10, 40),
                Size = new Size(100, 23),
                Text = "Start"
            };
            btnCopyFiles.Click += BtnCopyFiles_Click;
            this.Controls.Add(btnCopyFiles);

            btnPauseResume = new Button
            {
                Location = new Point(120, 40),
                Size = new Size(100, 23),
                Text = "Pause",
                Enabled = false
            };
            btnPauseResume.Click += BtnPauseResume_Click;
            this.Controls.Add(btnPauseResume);

            btnStop = new Button
            {
                Location = new Point(230, 40),
                Size = new Size(100, 23),
                Text = "Stop",
                Enabled = false
            };
            btnStop.Click += BtnStop_Click;
            this.Controls.Add(btnStop);

            rbtnCopy = new RadioButton
            {
                Location = new Point(340, 40),
                Size = new Size(80, 23),
                Text = "Copy",
                Checked = true
            };
            rbtnCopy.CheckedChanged += (s, e) => { isMoveOperation = false; };
            this.Controls.Add(rbtnCopy);

            rbtnMove = new RadioButton
            {
                Location = new Point(420, 40),
                Size = new Size(80, 23),
                Text = "Move"
            };
            rbtnMove.CheckedChanged += (s, e) => { isMoveOperation = true; };
            this.Controls.Add(rbtnMove);

            overalProgressbarLabel = new Label
            {
                Location = new Point(10, 80),
                Size = new Size(300, 23),
                Text = "Overall Progress: 0% (0/0 files remaining)"
            };
            this.Controls.Add(overalProgressbarLabel);

            progressBarOverall = new ProgressBar
            {
                Location = new Point(10, 105),
                Size = new Size(560, 23)
            };
            this.Controls.Add(progressBarOverall);

            currentProgressbarLabel = new Label
            {
                Location = new Point(10, 150),
                Size = new Size(300, 23),
                Text = "Current File Progress: 0%"
            };
            this.Controls.Add(currentProgressbarLabel);

            progressBarCurrent = new ProgressBar
            {
                Location = new Point(10, 175),
                Size = new Size(560, 23)
            };
            this.Controls.Add(progressBarCurrent);

            // Tree view for source files
            tvSourceTree = new TreeView
            {
                Location = new Point(580, 10),
                Size = new Size(390, 200),
                CheckBoxes = false
            };
            this.Controls.Add(tvSourceTree);

            // Labels for statistics
            lblFileCount = new Label
            {
                Location = new Point(580, 220),
                Size = new Size(360, 20),
                Text = "Total Files: 0"
            };
            this.Controls.Add(lblFileCount);

            lblFolderCount = new Label
            {
                Location = new Point(580, 240),
                Size = new Size(360, 20),
                Text = "Total Folders: 0"
            };
            this.Controls.Add(lblFolderCount);

            lblTotalSize = new Label
            {
                Location = new Point(580, 260),
                Size = new Size(360, 40),
                Text = "Total Size: 0 bytes (0 B)"
            };
            this.Controls.Add(lblTotalSize);

            rtbLogs = new RichTextBox
            {
                Location = new Point(10, 210),
                Size = new Size(960, 440),
                ReadOnly = true,
                Font = new Font("Arial", 10, FontStyle.Bold),
                WordWrap = true
            };
            this.Controls.Add(rtbLogs);

            chkAutoScroll = new CheckBox
            {
                Text = "Auto-scroll logs",
                Checked = true, // Enabled by default
                Location = new Point(rtbLogs.Left, rtbLogs.Bottom - 25),
                Width = 120,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            this.Controls.Add(chkAutoScroll);

            // Bring the checkbox to front so it's visible
            chkAutoScroll.BringToFront();
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                openFileDialog.Title = "Select an Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcelPath.Text = openFileDialog.FileName;
                    LoadSourceTree(openFileDialog.FileName);
                }
            }
        }

        private void LoadSourceTree(string excelFilePath)
        {
            // Set loading state immediately
            lblFileCount.Text = "Total Files: Loading...";
            lblFolderCount.Text = "Total Folders: Loading...";
            lblTotalSize.Text = "Total Size: Loading...";

            // Run the loading in a background task to keep UI responsive
            Task.Run(() =>
            {
                try
                {
                    tvSourceTree.Invoke(new Action(() => tvSourceTree.Nodes.Clear()));

                    long localTotalSize = 0;
                    int localTotalFiles = 0;
                    int localFolderCount = 0;

                    using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            string sourcePath = worksheet.Cells[row, 1].Text.Trim();

                            if (string.IsNullOrEmpty(sourcePath) || !Directory.Exists(sourcePath))
                                continue;

                            var rootNode = new TreeNode(sourcePath);
                            localFolderCount++;

                            // Process directory and update counts
                            ProcessDirectoryForTree(sourcePath, rootNode, ref localFolderCount, ref localTotalFiles, ref localTotalSize);

                            // Update tree view on UI thread
                            tvSourceTree.Invoke(new Action(() => tvSourceTree.Nodes.Add(rootNode)));
                        }
                    }

                    // Update UI with final values
                    this.Invoke(new Action(() =>
                    {
                        lblFileCount.Text = $"Total Files: {localTotalFiles:N0}";
                        lblFolderCount.Text = $"Total Folders: {localFolderCount:N0}";
                        lblTotalSize.Text = $"Total Size: {localTotalSize:N0} bytes ({FormatSize(localTotalSize)})";
                    }));
                }
                catch (Exception ex)
                {
                    this.Invoke(new Action(() =>
                    {
                        MessageBox.Show($"Error loading source tree: {ex.Message}");
                        lblFileCount.Text = "Total Files: Error";
                        lblFolderCount.Text = "Total Folders: Error";
                        lblTotalSize.Text = "Total Size: Error";
                    }));
                }
            });
        }

        private void ProcessDirectoryForTree(string path, TreeNode parentNode,
                                   ref int folderCount, ref int fileCount, ref long totalSize)
        {
            try
            {
                // Add files
                foreach (string file in Directory.GetFiles(path))
                {
                    var fileNode = new TreeNode(Path.GetFileName(file));
                    parentNode.Nodes.Add(fileNode);

                    try
                    {
                        var fileInfo = new FileInfo(file);
                        totalSize += fileInfo.Length;
                        fileCount++;
                    }
                    catch { }
                }

                // Add subdirectories
                foreach (string directory in Directory.GetDirectories(path))
                {
                    var dirNode = new TreeNode(Path.GetFileName(directory));
                    parentNode.Nodes.Add(dirNode);
                    folderCount++;
                    ProcessDirectoryForTree(directory, dirNode, ref folderCount, ref fileCount, ref totalSize);
                }
            }
            catch { }
        }

        private string FormatSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            int order = 0;
            double size = bytes;

            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }

            return $"{size:0.##} {sizes[order]}";
        }

        private async void BtnCopyFiles_Click(object sender, EventArgs e)
        {
            string excelFilePath = txtExcelPath.Text;

            if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
            {
                MessageBox.Show("Please select a valid Excel file.");
                return;
            }

            btnCopyFiles.Enabled = false;
            btnBrowse.Enabled = false;
            btnPauseResume.Enabled = true;
            btnStop.Enabled = true;
            progressBarOverall.Value = 0;
            progressBarCurrent.Value = 0;
            rtbLogs.Clear();
            successfulFiles.Clear();
            failedFiles.Clear();
            processedSizeBytes = 0;
            processedFiles = 0;
            isPaused = false;
            isStopped = false;
            overwriteAll = false;
            skipAll = false;
            cancellationTokenSource = new CancellationTokenSource();

            try
            {
                await Task.Run(() => ProcessExcelFile(excelFilePath, cancellationTokenSource.Token));
                if (!isStopped)
                {
                    MessageBox.Show(isMoveOperation ? "Files moved successfully!" : "Files copied successfully!");
                }
            }
            catch (Exception ex)
            {
                if (!isStopped)
                {
                    MessageBox.Show($"An error occurred: {ex.Message}");
                }
            }
            finally
            {
                btnCopyFiles.Enabled = true;
                btnBrowse.Enabled = true;
                btnPauseResume.Enabled = false;
                btnStop.Enabled = false;
            }
        }

        private void BtnPauseResume_Click(object sender, EventArgs e)
        {
            isPaused = !isPaused;
            btnPauseResume.Text = isPaused ? "Resume" : "Pause";

            if (isPaused)
            {
                LogMessage("Operation paused");
            }
            else
            {
                LogMessage("Operation resumed");
            }
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            isStopped = true;
            cancellationTokenSource?.Cancel();
            btnPauseResume.Enabled = false;
            LogMessage("Operation stopped by user");
        }

        private void ProcessExcelFile(string excelFilePath, CancellationToken cancellationToken)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                // First pass to count total operations (files + directories)
                totalOperations = 0;
                for (int row = 2; row <= rowCount; row++)
                {
                    string sourcePath = worksheet.Cells[row, 1].Text.Trim();
                    if (!string.IsNullOrEmpty(sourcePath) && Directory.Exists(sourcePath))
                    {
                        totalOperations += CountFilesAndDirectories(sourcePath);
                    }
                }

                completedOperations = 0;

                // Second pass to actually process
                for (int row = 2; row <= rowCount; row++)
                {
                    if (cancellationToken.IsCancellationRequested)
                        break;

                    while (isPaused && !cancellationToken.IsCancellationRequested)
                    {
                        Thread.Sleep(500);
                    }

                    if (cancellationToken.IsCancellationRequested)
                        break;

                    string sourcePath = worksheet.Cells[row, 1].Text.Trim();
                    string targetPath = worksheet.Cells[row, 2].Text.Trim();

                    if (string.IsNullOrEmpty(sourcePath) || string.IsNullOrEmpty(targetPath))
                    {
                        LogMessage($"Row {row} skipped due to empty source or target path.");
                        continue;
                    }

                    LogMessage($"Processing Row {row}: Source='{sourcePath}', Target='{targetPath}'");

                    CopyOrMoveDirectory(sourcePath, targetPath, cancellationToken);
                }
            }

            if (!cancellationToken.IsCancellationRequested)
            {
                LogSummary();
                SaveLogToExcel();
            }
        }

        // Helper method to count files and directories
        private int CountFilesAndDirectories(string path)
        {
            try
            {
                int count = Directory.GetFiles(path).Length;
                foreach (string dir in Directory.GetDirectories(path))
                {
                    count += CountFilesAndDirectories(dir);
                }
                return count + 1; // +1 for the directory itself
            }
            catch
            {
                return 1; // Count the directory even if we can't access its contents
            }
        }

        private void CopyOrMoveDirectory(string sourceDir, string targetDir, CancellationToken cancellationToken)
        {
            if (!Directory.Exists(sourceDir))
            {
                LogMessage($"Source directory does not exist: {sourceDir}");
                return;
            }

            Directory.CreateDirectory(targetDir);

            var files = Directory.GetFiles(sourceDir);
            for (int i = 0; i < files.Length; i++)
            {
                if (cancellationToken.IsCancellationRequested)
                    return;

                while (isPaused && !cancellationToken.IsCancellationRequested)
                {
                    Thread.Sleep(500);
                }

                if (cancellationToken.IsCancellationRequested)
                    return;

                string file = files[i];
                string fileName = Path.GetFileName(file);
                string destFile = Path.Combine(targetDir, fileName);

                try
                {
                    var fileInfo = new FileInfo(file);
                    long fileSize = fileInfo.Length;

                    if (File.Exists(destFile) && !overwriteAll && !skipAll)
                    {
                        var result = ShowFileConflictDialog(fileName);
                        if (result == DialogResult.No) // Skip
                        {
                            LogMessage($"Skipped: {fileName} (user chose to skip)");
                            continue;
                        }
                        else if (result == DialogResult.Yes) // Overwrite
                        {
                            // Continue to overwrite
                        }
                        else if (result == DialogResult.Ignore) // Skip all
                        {
                            skipAll = true;
                            LogMessage($"Skipped: {fileName} (user chose to skip all)");
                            continue;
                        }
                        else if (result == DialogResult.Retry) // Overwrite all
                        {
                            overwriteAll = true;
                            // Continue to overwrite
                        }
                    }

                    if (File.Exists(destFile) && skipAll)
                    {
                        LogMessage($"Skipped: {fileName} (skip all active)");
                        continue;
                    }

                    if (isMoveOperation)
                    {
                        if (File.Exists(destFile)) File.Delete(destFile);
                        File.Move(file, destFile);
                        LogMessage($"Moved: {fileName} to {targetDir}");
                        successfulFiles.Add(destFile);
                    }
                    else
                    {
                        if (File.Exists(destFile) && !overwriteAll) continue;
                        if (File.Exists(destFile)) File.Delete(destFile);
                        File.Copy(file, destFile);
                        LogMessage($"Copied: {fileName} to {targetDir}");
                        successfulFiles.Add(destFile);
                    }

                    processedSizeBytes += fileSize;
                    processedFiles++;

                    // Update progress after each file operation
                    lock (progressLock)
                    {
                        completedOperations++;
                        int progress = (int)((double)completedOperations / totalOperations * 100);
                        UpdateOverallProgress(progress, totalOperations - completedOperations);
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"Failed: {fileName} - Error: {ex.Message}");
                    failedFiles.Add(file);

                    // Still count failed operations in progress
                    lock (progressLock)
                    {
                        completedOperations++;
                        int progress = (int)((double)completedOperations / totalOperations * 100);
                        UpdateOverallProgress(progress, totalOperations - completedOperations);
                    }
                }

                UpdateCurrentProgress((i + 1) * 100 / files.Length);
            }

            foreach (string subDir in Directory.GetDirectories(sourceDir))
            {
                if (cancellationToken.IsCancellationRequested)
                    return;

                string destSubDir = Path.Combine(targetDir, Path.GetFileName(subDir));
                CopyOrMoveDirectory(subDir, destSubDir, cancellationToken);
            }

            // Count the directory itself when done
            lock (progressLock)
            {
                completedOperations++;
                int progress = (int)((double)completedOperations / totalOperations * 100);
                UpdateOverallProgress(progress, totalOperations - completedOperations);
            }

            if (isMoveOperation && Directory.Exists(sourceDir) && Directory.GetFileSystemEntries(sourceDir).Length == 0)
            {
                try
                {
                    Directory.Delete(sourceDir);
                }
                catch { }
            }
        }

        private DialogResult ShowFileConflictDialog(string fileName)
        {
            if (this.InvokeRequired)
            {
                return (DialogResult)this.Invoke(new Func<DialogResult>(() => ShowFileConflictDialog(fileName)));
            }

            using (var form = new Form())
            {
                form.Text = "File Conflict";
                form.Size = new Size(400, 200);
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.StartPosition = FormStartPosition.CenterParent;

                var label = new Label
                {
                    Text = $"The file '{fileName}' already exists. What would you like to do?",
                    Dock = DockStyle.Top,
                    Padding = new Padding(10),
                    AutoSize = true
                };

                var btnOverwrite = new Button { Text = "Overwrite", DialogResult = DialogResult.Yes };
                var btnOverwriteAll = new Button { Text = "Overwrite All", DialogResult = DialogResult.Retry };
                var btnSkip = new Button { Text = "Skip", DialogResult = DialogResult.No };
                var btnSkipAll = new Button { Text = "Skip All", DialogResult = DialogResult.Ignore };

                btnOverwrite.Click += (s, e) => form.Close();
                btnOverwriteAll.Click += (s, e) => form.Close();
                btnSkip.Click += (s, e) => form.Close();
                btnSkipAll.Click += (s, e) => form.Close();

                var flowLayout = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new Padding(10),
                    AutoSize = true
                };

                flowLayout.Controls.Add(btnOverwriteAll);
                flowLayout.Controls.Add(btnOverwrite);
                flowLayout.Controls.Add(btnSkipAll);
                flowLayout.Controls.Add(btnSkip);

                form.Controls.Add(label);
                form.Controls.Add(flowLayout);

                return form.ShowDialog(this);
            }
        }

        private void LogSummary()
        {
            long successfulSize = successfulFiles.Sum(f => {
                try { return new FileInfo(f).Length; } catch { return 0; }
            });
            long failedSize = failedFiles.Sum(f => {
                try { return new FileInfo(f).Length; } catch { return 0; }
            });

            LogMessage("\n=== Summary ===");
            LogMessage($"Total Successful Files: {successfulFiles.Count:N0}");
            LogMessage($"Total Successful Size: {FormatSize(successfulSize)}");
            LogMessage($"Total Failed Files: {failedFiles.Count:N0}");
            LogMessage($"Total Failed Size: {FormatSize(failedSize)}");

            LogMessage("\nSuccessful Files:");
            foreach (var file in successfulFiles)
            {
                LogMessage(file);
            }

            if (failedFiles.Count > 0)
            {
                LogMessage("\nFailed Files:");
                foreach (var file in failedFiles)
                {
                    LogMessage(file);
                }
            }
        }

        private void SaveLogToExcel()
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"CopyLog_{DateTime.Now:yyyyMMddHHmmss}.xlsx");

            using (var package = new ExcelPackage())
            {
                // Summary sheet
                var summarySheet = package.Workbook.Worksheets.Add("Summary");
                summarySheet.Cells[1, 1].Value = "Operation Type";
                summarySheet.Cells[1, 2].Value = isMoveOperation ? "Move" : "Copy";
                summarySheet.Cells[2, 1].Value = "Total Files Processed";
                summarySheet.Cells[2, 2].Value = successfulFiles.Count + failedFiles.Count;
                summarySheet.Cells[3, 1].Value = "Successful Files";
                summarySheet.Cells[3, 2].Value = successfulFiles.Count;
                summarySheet.Cells[4, 1].Value = "Failed Files";
                summarySheet.Cells[4, 2].Value = failedFiles.Count;
                summarySheet.Cells[5, 1].Value = "Start Time";
                summarySheet.Cells[5, 2].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                summarySheet.Cells[6, 1].Value = "End Time";
                summarySheet.Cells[6, 2].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Successful files sheet
                var successSheet = package.Workbook.Worksheets.Add("Successful Files");
                successSheet.Cells[1, 1].Value = "File Path";
                successSheet.Cells[1, 2].Value = "Size";
                successSheet.Cells[1, 3].Value = "Size (Bytes)";

                for (int i = 0; i < successfulFiles.Count; i++)
                {
                    successSheet.Cells[i + 2, 1].Value = successfulFiles[i];
                    try
                    {
                        var fileInfo = new FileInfo(successfulFiles[i]);
                        successSheet.Cells[i + 2, 2].Value = FormatSize(fileInfo.Length);
                        successSheet.Cells[i + 2, 3].Value = fileInfo.Length;
                    }
                    catch
                    {
                        successSheet.Cells[i + 2, 2].Value = "N/A";
                        successSheet.Cells[i + 2, 3].Value = "N/A";
                    }
                }

                // Failed files sheet
                if (failedFiles.Count > 0)
                {
                    var failSheet = package.Workbook.Worksheets.Add("Failed Files");
                    failSheet.Cells[1, 1].Value = "File Path";
                    failSheet.Cells[1, 2].Value = "Error";

                    for (int i = 0; i < failedFiles.Count; i++)
                    {
                        failSheet.Cells[i + 2, 1].Value = failedFiles[i];
                        // Error message would need to be tracked separately for each file
                        failSheet.Cells[i + 2, 2].Value = "Error occurred during operation";
                    }
                }

                package.SaveAs(new FileInfo(logPath));
            }

            LogMessage($"\nLog saved to: {logPath}");
        }

        private void UpdateOverallProgress(int value, int remainingOperations)
        {
            if (progressBarOverall.InvokeRequired)
            {
                progressBarOverall.Invoke(new Action(() =>
                {
                    progressBarOverall.Value = value;
                    overalProgressbarLabel.Text = $"Overall Progress: {value}% ({remainingOperations:N0}/{totalOperations:N0} operations remaining)";
                }));
            }
            else
            {
                progressBarOverall.Value = value;
                overalProgressbarLabel.Text = $"Overall Progress: {value}% ({remainingOperations:N0}/{totalOperations:N0} operations remaining)";
            }
        }

        private void UpdateCurrentProgress(int value)
        {
            if (progressBarCurrent.InvokeRequired)
            {
                progressBarCurrent.Invoke(new Action(() =>
                {
                    progressBarCurrent.Value = value;
                    currentProgressbarLabel.Text = $"Current File Progress: {value}%";
                }));
            }
            else
            {
                progressBarCurrent.Value = value;
                currentProgressbarLabel.Text = $"Current File Progress: {value}%";
            }
        }

        private void LogMessage(string message)
        {
            if (rtbLogs.InvokeRequired)
            {
                rtbLogs.Invoke(new Action(() =>
                {
                    rtbLogs.AppendText(message + Environment.NewLine);
                    if (chkAutoScroll.Checked)
                    {
                        rtbLogs.SelectionStart = rtbLogs.TextLength;
                        rtbLogs.ScrollToCaret();
                    }
                }));
            }
            else
            {
                rtbLogs.AppendText(message + Environment.NewLine);
                if (chkAutoScroll.Checked)
                {
                    rtbLogs.SelectionStart = rtbLogs.TextLength;
                    rtbLogs.ScrollToCaret();
                }
            }
        }
    }
}