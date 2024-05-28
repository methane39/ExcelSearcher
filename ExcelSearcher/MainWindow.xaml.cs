using System.IO;
using System.Windows;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;
using System;
using System.Windows.Controls;
using System.ComponentModel;

namespace ExcelSearcher
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private System.Windows.Controls.ProgressBar progressBar;
        private TextBlock statusText;
        private BackgroundWorker backgroundWorker;

        string filePathIn;
        string filePathOut;
        string name;
        string id;
        public MainWindow()
        {
            InitializeComponent();
        }
        
        private void InClick(object sender, RoutedEventArgs e)
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "请选择输入路径";
                folderBrowserDialog.ShowNewFolderButton = true;

                // 显示对话框并获取结果
                DialogResult result = folderBrowserDialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    // 将选中的文件夹路径赋值给 InTextBox
                    InTextBox.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void OutClick(object sender, RoutedEventArgs e)
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "请选择输出路径";
                folderBrowserDialog.ShowNewFolderButton = true;

                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    OutTextBox.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void SearchClick(object sender, RoutedEventArgs e)
        {
            StartSearch.IsEnabled = false;
            nameBox.IsReadOnly = true;
            idBox.IsReadOnly = true;
            filePathIn = InTextBox.Text;
            filePathOut = OutTextBox.Text;
            name = nameBox.Text;
            id = idBox.Text;
            if (CheckValid(filePathIn, filePathOut, name, id))
            {
                StartSearch.Visibility = Visibility.Collapsed;
                progressBar = new System.Windows.Controls.ProgressBar { 
                    Width = 430,
                    Height = 40,
                    Minimum = 0,
                    Maximum = 100,
                    Value = 0
                };
                statusText = new TextBlock
                {
                    Text = "初始化中",
                    Margin = new Thickness(0, 5, 0, 0),
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Center
                };
                ProgressContainer.Children.Add(statusText);
                ProgressContainer.Children.Add(progressBar);

                backgroundWorker = new BackgroundWorker
                {
                    WorkerReportsProgress = true,
                    WorkerSupportsCancellation = true
                };
                backgroundWorker.DoWork += BackgroundWorker_DoWork;
                backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
                backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
                backgroundWorker.RunWorkerAsync();

                nameBox.IsReadOnly = false;
                idBox.IsReadOnly = false;
                StartSearch.IsEnabled = true;
            }
            else
            {
                nameBox.IsReadOnly = false;
                idBox.IsReadOnly = false;
                StartSearch.IsEnabled = true;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private bool CheckValid(string inPath, string outPath, string name, string id)
        {
            if (inPath == null || outPath == null) { System.Windows.MessageBox.Show("请输入有效的路径"); return false; }
            if (!Directory.Exists(inPath) || !Directory.Exists(outPath)) { System.Windows.MessageBox.Show("请输入有效的路径"); return false; }
            if (name == null || id == null)
            {
                System.Windows.MessageBox.Show("学号、姓名不可为空");
                return false;
            } else if (id.Length != 10)
            {
                System.Windows.MessageBox.Show("学号长度错误");
                return false;
            }
            return true;
        }
            
        private void ResultWindow_Closed(object sender, EventArgs e)
        {
            var resultWindow = sender as ResultWindow;
            resultWindow.Closed -= ResultWindow_Closed;
            resultWindow = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            ExcelService.ProgressReporter reporter = (progress, currentFile) =>
            {
                backgroundWorker.ReportProgress(progress, currentFile);
            };
            string[] results = ExcelService.Search(filePathIn, name, id, reporter);
            if (results.Length > 0)
            {
                Dispatcher.Invoke(() =>
                {
                    ResultWindow resultWindow = new ResultWindow(results);
                    resultWindow.Closed += ResultWindow_Closed;
                    resultWindow.Show();
                });
                
                
                results = null;

            }
            else if (results.Length == 0)
            {
                System.Windows.MessageBox.Show("未能找到结果");
                results = null;

            }
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // 更新进度条和提示文本
            progressBar.Value = e.ProgressPercentage;
            statusText.Text = $"Processing {e.UserState}...";
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                statusText.Text = "Operation Cancelled";
            }
            else if (e.Error != null)
            {
                statusText.Text = $"Error: {e.Error.Message}";
            }
            else
            {
                statusText.Text = "Search Completed";
            }

            // 清理操作
            ProgressContainer.Children.Remove(progressBar);
            StartSearch.Visibility = Visibility.Visible;
        }
    }

}
