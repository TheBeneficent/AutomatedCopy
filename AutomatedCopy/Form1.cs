using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
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
        private RadioButton rbtnCopy;
        private RadioButton rbtnMove;
        private Label overalProgressbarLabel;
        private ProgressBar progressBarOverall;
        private Label currentProgressbarLabel;
        private ProgressBar progressBarCurrent;
        private RichTextBox rtbLogs;
        private bool isMoveOperation = false;

        private List<string> successfulFiles = new List<string>();
        private List<string> failedFiles = new List<string>();

        public Form1()
        {
            this.Text = "Automated Copy";
            this.Name = "Automated Copy";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            InitializeControls();
        }

        private void InitializeControls()
        {
            this.Size = new Size(600, 550);

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

            rbtnCopy = new RadioButton
            {
                Location = new Point(120, 40),
                Size = new Size(80, 23),
                Text = "Copy",
                Checked = true
            };
            rbtnCopy.CheckedChanged += (s, e) => { isMoveOperation = false; };
            this.Controls.Add(rbtnCopy);

            rbtnMove = new RadioButton
            {
                Location = new Point(200, 40),
                Size = new Size(80, 23),
                Text = "Move"
            };
            rbtnMove.CheckedChanged += (s, e) => { isMoveOperation = true; };
            this.Controls.Add(rbtnMove);

            overalProgressbarLabel = new Label
            {
                Location = new Point(10, 80),
                Size = new Size(150, 23),
                Text = "Overall Progress"
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
                Size = new Size(150, 23),
                Text = "Current File Progress"
            };
            this.Controls.Add(currentProgressbarLabel);

            progressBarCurrent = new ProgressBar
            {
                Location = new Point(10, 175),
                Size = new Size(560, 23)
            };
            this.Controls.Add(progressBarCurrent);

            rtbLogs = new RichTextBox
            {
                Location = new Point(10, 210),
                Size = new Size(560, 280),
                ReadOnly = true,
                Font = new Font("Arial", 10, FontStyle.Bold),
                WordWrap = true
            };
            this.Controls.Add(rtbLogs);
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
                }
            }
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
            progressBarOverall.Value = 0;
            progressBarCurrent.Value = 0;
            rtbLogs.Clear();
            successfulFiles.Clear();
            failedFiles.Clear();

            try
            {
                await Task.Run(() => ProcessExcelFile(excelFilePath));
                MessageBox.Show(isMoveOperation ? "Files moved successfully!" : "Files copied successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
            finally
            {
                btnCopyFiles.Enabled = true;
            }
        }

        private void ProcessExcelFile(string excelFilePath)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string sourcePath = worksheet.Cells[row, 1].Text.Trim();
                    string targetPath = worksheet.Cells[row, 2].Text.Trim();

                    if (string.IsNullOrEmpty(sourcePath) || string.IsNullOrEmpty(targetPath))
                    {
                        LogMessage($"Row {row} skipped due to empty source or target path.");
                        continue;
                    }

                    LogMessage($"Processing Row {row}: Source='{sourcePath}', Target='{targetPath}'");

                    CopyOrMoveDirectory(sourcePath, targetPath);

                    UpdateOverallProgress((row - 1) * 100 / (rowCount - 1));
                }
            }

            LogSummary();
            SaveLogToFile(); 
        }

        private void CopyOrMoveDirectory(string sourceDir, string targetDir)
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
                string file = files[i];
                string fileName = Path.GetFileName(file);
                string destFile = Path.Combine(targetDir, fileName);


                try
                {
                    if (!File.Exists(destFile))
                    {
                        if (isMoveOperation)
                        {
                            File.Move(file, destFile);
                            LogMessage($"Moved: {fileName} to {targetDir}");
                        }
                        else
                        {
                            File.Copy(file, destFile);
                            LogMessage($"Copied: {fileName} to {targetDir}");
                        }
                        successfulFiles.Add(destFile);
                    }
                    else
                    {
                        LogMessage($"Skipped: {fileName} (already exists in target)");
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"Failed: {fileName} - Error: {ex.Message}");
                    failedFiles.Add(file);
                }

                UpdateCurrentProgress((i + 1) * 100 / files.Length);
            }

            foreach (string subDir in Directory.GetDirectories(sourceDir))
            {
                string destSubDir = Path.Combine(targetDir, Path.GetFileName(subDir));
                CopyOrMoveDirectory(subDir, destSubDir);
            }

            if (isMoveOperation && Directory.Exists(sourceDir) && Directory.GetFileSystemEntries(sourceDir).Length == 0)
            {
                Directory.Delete(sourceDir);
            }
        }

        private void LogSummary()
        {
            LogMessage("\n=== Summary ===");
            LogMessage($"Total Successful Files: {successfulFiles.Count}");
            LogMessage($"Total Failed Files: {failedFiles.Count}");

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

        private void SaveLogToFile()
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LogFile.txt");
            using (StreamWriter writer = new StreamWriter(logPath, false, Encoding.UTF8))
            {
                writer.Write(rtbLogs.Text);
            }
        }

        private void UpdateOverallProgress(int value)
        {
            if (progressBarOverall.InvokeRequired)
                progressBarOverall.Invoke(new Action(() => progressBarOverall.Value = value));
            else
                progressBarOverall.Value = value;
        }

        private void UpdateCurrentProgress(int value)
        {
            if (progressBarCurrent.InvokeRequired)
                progressBarCurrent.Invoke(new Action(() => progressBarCurrent.Value = value));
            else
                progressBarCurrent.Value = value;
        }

        private void LogMessage(string message)
        {
            if (rtbLogs.InvokeRequired)
                rtbLogs.Invoke(new Action(() => rtbLogs.AppendText(message + Environment.NewLine)));
            else
                rtbLogs.AppendText(message + Environment.NewLine);
        }
    }
}
