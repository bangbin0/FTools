using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Aspose.Words;
using Aspose.Words.Saving;
using Microsoft.SqlServer.Server;
using Sunny.UI;
using Sunny.UI.Win32;
using Tools.Utils;
using static System.Net.Mime.MediaTypeNames;

namespace Tools
{
    public partial class Form1 : UIForm
    {

        List<string> wordFilesList = new List<string>();
        List<string> ReNameFilesList = new List<string>();
        public Form1()
        {
            InitializeComponent();
        }
        #region word2pdf
        private void BtnSeachWJJ2_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)

                {
                    string selectedPath = folderDialog.SelectedPath;
                    this.folderPath2.Text = selectedPath;
                    if (String.IsNullOrEmpty(targetPath.Text))
                    {

                        targetPath.Text = this.folderPath2.Text + "\\target";
                    }
                    AddFile2List();
                }
            }
        }

        private async void StartChangeBtn2_Click(object sender, EventArgs e)
        {
            bool error = false;
            if (string.IsNullOrEmpty(this.folderPath2.Text))
            {
                MessageBox.Show("请选择转换文件夹");
                error = true;
            }
            if (TypeBox.Items.Count == 0)
            {
                MessageBox.Show("请选择转换类型");
                error = true;
            }
            if (!error)
            {

                uiProcessBar1.Value = 0;
                // 确保输出目录存在
                if (!Directory.Exists(this.folderPath2.Text))
                {
                    Directory.CreateDirectory(this.folderPath2.Text);
                }
                object[] objects = FileList.SelectedItems.ToArray();
                List<string> wordFiles = objects.Cast<string>().ToList();

                List<string> failedFiles = new List<string>(); // 记录失败的文件
                int successCount = 0; // 成功转换的文件计数

                // 异步处理文件转换
                List<Task> tasks = new List<Task>();
                // 输出转换结果

                for (int i = 0; i < wordFiles.Count; i++)
                {
                    int index = i; // 必须在循环中保存索引变量
                    string wordFile = wordFiles[i];

                    // 创建一个任务
                    Task task = Task.Run(() =>
                    {
                        string pdfFilePath = "";
                        if (string.IsNullOrEmpty(this.targetPath.Text))
                        {
                            pdfFilePath = Path.Combine(this.folderPath2.Text + "\\target", Path.GetFileNameWithoutExtension(wordFile) + ".pdf");
                        }
                        else
                        {
                            pdfFilePath = Path.Combine(this.targetPath.Text, Path.GetFileNameWithoutExtension(wordFile) + ".pdf");
                        }

                        try
                        {
                            // 进行文件转换
                            Document doc = new Document(wordFile);
                            PdfSaveOptions saveOptions = new PdfSaveOptions();

                            // 设置PDF加密选项
                            PdfEncryptionDetails encryptionDetails = new PdfEncryptionDetails(string.Empty, "password", PdfEncryptionAlgorithm.RC4_128);
                            encryptionDetails.Permissions = PdfPermissions.AllowAll;
                            saveOptions.EncryptionDetails = encryptionDetails;

                            // 保存为 PDF 文件
                            doc.Save(pdfFilePath, saveOptions);

                            // 更新进度条（通过主线程更新 UI）
                            this.Invoke(new Action(() =>
                            {
                                uiProcessBar1.Value = index + 1; // 更新进度条
                            }));
                            // 成功计数
                            this.Invoke(new Action(() =>
                            {
                                successCount++;
                            }));
                        }
                        catch (Exception ex)
                        {
                            // 记录失败的文件
                            this.Invoke(new Action(() =>
                            {
                                failedFiles.Add(wordFile); // 添加到失败列表
                            }));
                            // 异常处理，记录错误
                            Console.WriteLine($"Error converting file {wordFile}: {ex.Message}");
                        }
                    });

                    // 添加任务到任务列表
                    tasks.Add(task);
                }
                await Task.WhenAll(tasks);

                // 弹出提示框，显示成功和失败数量
                MessageBox.Show($"转换完成！\n成功转换 {successCount} 个文件。\n失败 {failedFiles.Count} 个文件。",
                                "转换结果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 输出失败的文件列表到目标文件夹
                if (failedFiles.Count > 0)
                {
                    string failedListFilePath = Path.Combine(this.targetPath.Text, "failed_files.txt");
                    System.IO.File.WriteAllLines(failedListFilePath, failedFiles);
                    MessageBox.Show($"失败的文件已保存到 {failedListFilePath}", "失败文件列表", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void SeachtargetPathBtn_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderDialog.SelectedPath;
                    this.targetPath.Text = selectedPath;
                }
            }
        }

        private void OpenTargetBtn_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(this.targetPath.Text))
            {
                Process.Start("explorer.exe", this.targetPath.Text);
            }
        }


        private void AddFile2List()
        {
            if (!String.IsNullOrEmpty(this.folderPath2.Text))
            {
                FileList.Items.Clear();
                wordFilesList.Clear();

                object[] objects = TypeBox.SelectedItems.ToArray();
                List<string> TypeBoxList = objects.Cast<string>().ToList();
                wordFilesList = FileHelper.FindFiles(folderPath2.Text, TypeBoxList, P1Sun.Checked);
                // 将列表转换为数组
                string[] wordFiles = wordFilesList.ToArray();

                for (int i = 0; i < wordFiles.Length; i++)
                {
                    string wordFile = wordFiles[i];

                    FileList.Items.Add(wordFile);
                }

            }
        }



        private void selectAllBtn_Click(object sender, EventArgs e)
        {
            FileList.SelectAll();
        }
        private void FileList_ValueChanged(object sender, CheckBoxGroupEventArgs e)
        {
            uiProcessBar1.Maximum = FileList.SelectedItems.Count;

        }


        private void NoSeachBtn_Click(object sender, EventArgs e)
        {
            FileList.UnSelectAll();
        }

        private void TypeBox_ValueChanged(object sender, CheckBoxGroupEventArgs e)
        {
            AddFile2List();
        }

        private void P1Sun_CheckedChanged(object sender, EventArgs e)
        {
            AddFile2List();
        }

        #endregion
        #region 批量重名名文件
        private void SeachBtn_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderDialog.SelectedPath;
                    this.ChangeText.Text = selectedPath;
                }
            }
        }

        private async void ReplaceBtn_Click(object sender, EventArgs e)
        {

            if (Replacetype.SelectedItems.Count != 0)
            {
                P2Pro.Value = 0;
                bool error = false;
                if (string.IsNullOrEmpty(this.ChangeText.Text))
                {
                    MessageBox.Show("请选择替换文件所在文件夹");
                    error = true;
                }
                if (string.IsNullOrEmpty(this.FindText.Text))
                {
                    MessageBox.Show("请填写查找内容");
                    error = true;
                }
                ReNameFilesList.Clear();
                if (!error)
                {

                    object[] objects = Replacetype.SelectedItems.ToArray();
                    List<string> ReNameTypeList = objects.Cast<string>().ToList();
                    ReNameFilesList = FileHelper.FindFiles(ChangeText.Text, ReNameTypeList, P2Sun.Checked);
                    List<string> ReNameTempList = new List<string>();
                    //查找文件名包含替换内容的文件
                    for (int i = 0; i < ReNameFilesList.Count; i++)
                    {

                        if (!FindText.Text.IsNullOrEmpty())
                        {
                            if (ZZCHeck.Checked)
                            {
                                // 这里用户可以输入正则表达式，进行匹配
                                string customRegexPattern = FindText.Text;
                                // 判断用户输入的正则表达式是否正确
                                if (!Tools.StringHelper.IsValidRegex(customRegexPattern))
                                {
                                    MessageBox.Show("输入的正则表达式无效，请重新输入！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                                // 根据用户输入的正则表达式来过滤文件
                                ReNameTempList = ReNameFilesList.Where(file =>
                                    Regex.IsMatch(Path.GetFileName(file), customRegexPattern)).ToList();
                            }
                            else
                            {
                                string wordFile = ReNameFilesList[i];
                                bool containsPattern = Path.GetFileName(wordFile).Contains(FindText.Text);


                                // 输出结果
                                if (containsPattern)
                                {
                                    ReNameTempList.Add(wordFile);
                                }
                            }
                        }
                    }
                    ReNameFilesList = ReNameTempList;
                    // 异步处理文件转换
                    List<Task> tasks = new List<Task>();
                    // 输出转换结果
                    P2Pro.Maximum = ReNameFilesList.Count;
                    for (int i = 0; i < ReNameFilesList.Count; i++)
                    {
                        int index = i; // 必须在循环中保存索引变量
                        string wordFile = ReNameFilesList[i];
                        P2Pro.Value = i + 1;
                        // 创建一个任务
                        Task task = Task.Run(() =>
                        {
                            try
                            {
                                // 获取文件目录、文件名（不含扩展名）和扩展名
                                string directory = Path.GetDirectoryName(wordFile);
                                string fileName = Path.GetFileName(wordFile);
                                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
                                string fileExtension = Path.GetExtension(fileName);

                                // 替换文件名中的指定内容，但保留扩展名
                                string newFileName = fileName;

                                // 使用正则表达式进行文件名替换
                                if (ZZCHeck.Checked)
                                {
                                    newFileName = Regex.Replace(fileNameWithoutExtension, FindText.Text, ThText.Text) + fileExtension;
                                }
                                else
                                {
                                    newFileName = fileNameWithoutExtension.Replace(FindText.Text, ThText.Text) + fileExtension;
                                }

                                // 生成新文件的完整路径
                                string newFilePath = Path.Combine(directory, newFileName);

                                // 确保文件名没有重复，避免覆盖已有文件
                                if (System.IO.File.Exists(newFilePath))
                                {
                                    // 如果目标文件已存在，可以选择处理，如跳过或者重命名
                                    newFilePath = Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(newFileName)}_{Guid.NewGuid()}{Path.GetExtension(newFileName)}");
                                }

                                // 执行重命名操作
                                System.IO.File.Move(wordFile, newFilePath);
                            }
                            catch (Exception ex)
                            {
                                // 异常处理
                                P2Logs.Text += $"文件替换失败: {ex.Message}\r\n";
                            }
                        });

                        // 添加任务到任务列表
                        tasks.Add(task);
                    }
                    await Task.WhenAll(tasks);
                }
            }
        }

        private void CollectFilesSeachFindPathBtn_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)

                {
                    string selectedPath = folderDialog.SelectedPath;
                    this.CollectFilesFindfilePath.Text = selectedPath;
                    if (String.IsNullOrEmpty(CollectFilesSavePath.Text))
                    {

                        CollectFilesSavePath.Text = this.CollectFilesFindfilePath.Text + "\\target";
                    }
                }
            }
        }
        #endregion

        #region 收集文件
        private void CollectFilesOpenSaveBtn_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(this.CollectFilesSavePath.Text))
            {
                Process.Start("explorer.exe", this.CollectFilesSavePath.Text);
            }
        }

        private void CollectFilesSavePathBtn_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderDialog.SelectedPath;
                    CollectFilesSavePath.Text = selectedPath;
                }
            }
        }

        private async void CollectFilesStartBtn_Click(object sender, EventArgs e)
        {
            P3Pro.Value = 0;
            // 确保目标目录存在
            if (!Directory.Exists(CollectFilesSavePath.Text))
            {
                Directory.CreateDirectory(CollectFilesSavePath.Text);
            }

            object[] objects = CollectfFilesType.SelectedItems.ToArray();
            List<string> CollectfFilesTypes = objects.Cast<string>().ToList();

            List<string> FindFileList = new List<string>();
            FindFileList = FileHelper.FindFiles(CollectFilesFindfilePath.Text, CollectfFilesTypes, CollectFilesFindSunFIles.Checked);
            List<Task> tasks = new List<Task>();
            P3Pro.Maximum = FindFileList.Count;
            for (int i = 0; i < FindFileList.Count; i++)
            {
                int index = i; // 必须在循环中保存索引变量
                string sFile = FindFileList[i];
                P3Pro.Value = i + 1;
                // 创建一个任务
                Task task = Task.Run(() =>
                {
                    try
                    {
                        // 如果目标文件已经存在，则可以根据需要决定如何处理，比如覆盖或者跳过
                        if (System.IO.File.Exists(sFile))
                        {
                            // 例如：跳过文件拷贝（你可以选择覆盖文件等）
                            CollectFilesLogs.Text += $"文件已存在，跳过拷贝: {sFile}\r\n";
                        }
                        else
                        {
                            // 执行文件拷贝操作
                            System.IO.File.Copy(sFile, CollectFilesSavePath.Text);
                            CollectFilesLogs.Text += $"文件拷贝成功: {sFile}\r\n";
                        }
                    }
                    catch (Exception ex)
                    {
                        CollectFilesLogs.Text += $"文件替换失败: {sFile},{ex.Message}\r\n";
                    }
                });

                // 添加任务到任务列表
                tasks.Add(task);
            }
            await Task.WhenAll(tasks);
        }



        #endregion

        #region pdf文件操作
        private Task _currentTask;
        private void p4_SeachFileBtn_Click(object sender, EventArgs e)
        {
            // 创建 OpenFileDialog 实例
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 设置文件类型过滤器，只允许选择 PDF 文件
            openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf";

            // 设置是否允许多选
            openFileDialog.Multiselect = false;  // 设置为 true 允许多选文件

            // 打开文件对话框并判断用户是否选择了文件
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // 获取选中的文件路径
                string selectedPdfFile = openFileDialog.FileName;

                p4_PdfFilePath.Text = selectedPdfFile;
            }
        }
        private async void p4_StartBtn_Click(object sender, EventArgs e)
        {
            p4_waitbar.Visible = true;
            p4_StartBtn.Enabled = false;

            // 创建CancellationTokenSource

            try
            {
                // 调用异步方法并传递CancellationToken
                _currentTask = pdf2word.ConvertWordToImagesAsync(p4_PdfFilePath.Text);

                // 等待任务完成
                await _currentTask;
            }
            catch (OperationCanceledException)
            {
                // 任务被取消时的处理
                MessageBox.Show("任务被取消");
            }
            finally
            {
                p4_waitbar.Visible = false;
                p4_StartBtn.Enabled = true;
            }

        }
    
  
        #endregion


    }
}
