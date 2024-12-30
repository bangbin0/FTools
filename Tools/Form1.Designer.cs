using System.Windows.Forms;

namespace Tools
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.uiTabControlMenu1 = new Sunny.UI.UITabControlMenu();
            this.p1 = new System.Windows.Forms.TabPage();
            this.targetPath = new Sunny.UI.UITextBox();
            this.folderPath2 = new Sunny.UI.UITextBox();
            this.FileList = new Sunny.UI.UICheckBoxGroup();
            this.selectAllBtn = new Sunny.UI.UIButton();
            this.NoSeachBtn = new Sunny.UI.UIButton();
            this.P1Sun = new Sunny.UI.UICheckBox();
            this.OpenTargetBtn = new Sunny.UI.UIButton();
            this.uiLabel2 = new Sunny.UI.UILabel();
            this.uiLabel1 = new Sunny.UI.UILabel();
            this.SeachtargetPathBtn = new Sunny.UI.UIButton();
            this.TypeBox = new Sunny.UI.UICheckBoxGroup();
            this.StartChangeBtn2 = new Sunny.UI.UIButton();
            this.uiProcessBar1 = new Sunny.UI.UIProcessBar();
            this.BtnSeachWJJ2 = new Sunny.UI.UIButton();
            this.p2 = new System.Windows.Forms.TabPage();
            this.ChangeText = new Sunny.UI.UITextBox();
            this.P2Logs = new Sunny.UI.UIRichTextBox();
            this.P2Sun = new Sunny.UI.UICheckBox();
            this.Replacetype = new Sunny.UI.UICheckBoxGroup();
            this.uiCheckBoxGroup1 = new Sunny.UI.UICheckBoxGroup();
            this.P2Pro = new Sunny.UI.UIProcessBar();
            this.ZZCHeck = new Sunny.UI.UICheckBox();
            this.ReplaceBtn = new Sunny.UI.UIButton();
            this.uiLabel5 = new Sunny.UI.UILabel();
            this.ThText = new Sunny.UI.UITextBox();
            this.uiLabel4 = new Sunny.UI.UILabel();
            this.FindText = new Sunny.UI.UITextBox();
            this.uiLabel3 = new Sunny.UI.UILabel();
            this.SeachBtn = new Sunny.UI.UIButton();
            this.p3 = new System.Windows.Forms.TabPage();
            this.CollectFilesFindfilePath = new Sunny.UI.UITextBox();
            this.CollectFilesSavePath = new Sunny.UI.UITextBox();
            this.CollectFilesLogs = new Sunny.UI.UIRichTextBox();
            this.CollectFilesFindSunFIles = new Sunny.UI.UICheckBox();
            this.P3Pro = new Sunny.UI.UIProcessBar();
            this.CollectFilesOpenSaveBtn = new Sunny.UI.UIButton();
            this.CollectFilesStartBtn = new Sunny.UI.UIButton();
            this.CollectfFilesType = new Sunny.UI.UICheckBoxGroup();
            this.uiLabel6 = new Sunny.UI.UILabel();
            this.uiLabel7 = new Sunny.UI.UILabel();
            this.CollectFilesSavePathBtn = new Sunny.UI.UIButton();
            this.CollectFilesSeachFindPathBtn = new Sunny.UI.UIButton();
            this.p4 = new System.Windows.Forms.TabPage();
            this.p4_waitbar = new Sunny.UI.UIWaitingBar();
            this.p4_PdfFilePath = new Sunny.UI.UITextBox();
            this.uiLabel8 = new Sunny.UI.UILabel();
            this.p4_SeachFileBtn = new Sunny.UI.UIButton();
            this.p4_StartBtn = new Sunny.UI.UIButton();
            this.uiStyleManager1 = new Sunny.UI.UIStyleManager(this.components);
            this.uiTabControl1 = new Sunny.UI.UITabControl();
            this.FileOperations = new System.Windows.Forms.TabPage();
            this.uiTabControlMenu1.SuspendLayout();
            this.p1.SuspendLayout();
            this.FileList.SuspendLayout();
            this.p2.SuspendLayout();
            this.uiCheckBoxGroup1.SuspendLayout();
            this.p3.SuspendLayout();
            this.p4.SuspendLayout();
            this.uiTabControl1.SuspendLayout();
            this.FileOperations.SuspendLayout();
            this.SuspendLayout();
            // 
            // uiTabControlMenu1
            // 
            this.uiTabControlMenu1.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.uiTabControlMenu1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.uiTabControlMenu1.Controls.Add(this.p1);
            this.uiTabControlMenu1.Controls.Add(this.p2);
            this.uiTabControlMenu1.Controls.Add(this.p3);
            this.uiTabControlMenu1.Controls.Add(this.p4);
            this.uiTabControlMenu1.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.uiTabControlMenu1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiTabControlMenu1.Location = new System.Drawing.Point(0, 0);
            this.uiTabControlMenu1.Margin = new System.Windows.Forms.Padding(0);
            this.uiTabControlMenu1.MenuStyle = Sunny.UI.UIMenuStyle.Custom;
            this.uiTabControlMenu1.Multiline = true;
            this.uiTabControlMenu1.Name = "uiTabControlMenu1";
            this.uiTabControlMenu1.Padding = new System.Drawing.Point(0, 0);
            this.uiTabControlMenu1.SelectedIndex = 0;
            this.uiTabControlMenu1.Size = new System.Drawing.Size(1186, 565);
            this.uiTabControlMenu1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.uiTabControlMenu1.TabIndex = 1;
            this.uiTabControlMenu1.TabUnSelectedForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.uiTabControlMenu1.TextAlignment = System.Windows.Forms.HorizontalAlignment.Left;
            this.uiTabControlMenu1.ZoomScaleDisabled = true;
            // 
            // p1
            // 
            this.p1.Controls.Add(this.targetPath);
            this.p1.Controls.Add(this.folderPath2);
            this.p1.Controls.Add(this.FileList);
            this.p1.Controls.Add(this.P1Sun);
            this.p1.Controls.Add(this.OpenTargetBtn);
            this.p1.Controls.Add(this.uiLabel2);
            this.p1.Controls.Add(this.uiLabel1);
            this.p1.Controls.Add(this.SeachtargetPathBtn);
            this.p1.Controls.Add(this.TypeBox);
            this.p1.Controls.Add(this.StartChangeBtn2);
            this.p1.Controls.Add(this.uiProcessBar1);
            this.p1.Controls.Add(this.BtnSeachWJJ2);
            this.p1.Location = new System.Drawing.Point(201, 0);
            this.p1.Name = "p1";
            this.p1.Size = new System.Drawing.Size(985, 565);
            this.p1.TabIndex = 0;
            this.p1.Text = "批量word转pdf";
            this.p1.UseVisualStyleBackColor = true;
            // 
            // targetPath
            // 
            this.targetPath.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.targetPath.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.targetPath.Location = new System.Drawing.Point(128, 55);
            this.targetPath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.targetPath.MinimumSize = new System.Drawing.Size(1, 13);
            this.targetPath.Name = "targetPath";
            this.targetPath.Padding = new System.Windows.Forms.Padding(4);
            this.targetPath.ShowText = false;
            this.targetPath.Size = new System.Drawing.Size(741, 38);
            this.targetPath.TabIndex = 30;
            this.targetPath.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.targetPath.Watermark = "";
            // 
            // folderPath2
            // 
            this.folderPath2.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.folderPath2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.folderPath2.Location = new System.Drawing.Point(128, 9);
            this.folderPath2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.folderPath2.MinimumSize = new System.Drawing.Size(1, 13);
            this.folderPath2.Name = "folderPath2";
            this.folderPath2.Padding = new System.Windows.Forms.Padding(4);
            this.folderPath2.ShowText = false;
            this.folderPath2.Size = new System.Drawing.Size(741, 38);
            this.folderPath2.TabIndex = 27;
            this.folderPath2.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.folderPath2.Watermark = "";
            // 
            // FileList
            // 
            this.FileList.Controls.Add(this.NoSeachBtn);
            this.FileList.Controls.Add(this.selectAllBtn);
            this.FileList.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FileList.HoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(236)))), ((int)(((byte)(255)))));
            this.FileList.Location = new System.Drawing.Point(8, 218);
            this.FileList.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.FileList.MinimumSize = new System.Drawing.Size(1, 1);
            this.FileList.Name = "FileList";
            this.FileList.Padding = new System.Windows.Forms.Padding(0, 32, 0, 0);
            this.FileList.SelectedIndexes = ((System.Collections.Generic.List<int>)(resources.GetObject("FileList.SelectedIndexes")));
            this.FileList.Size = new System.Drawing.Size(966, 274);
            this.FileList.TabIndex = 39;
            this.FileList.Text = "文件列表";
            this.FileList.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.FileList.ValueChanged += new Sunny.UI.UICheckBoxGroup.OnValueChanged(this.FileList_ValueChanged);
            // 
            // selectAllBtn
            // 
            this.selectAllBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.selectAllBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.selectAllBtn.Location = new System.Drawing.Point(88, 3);
            this.selectAllBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.selectAllBtn.Name = "selectAllBtn";
            this.selectAllBtn.Size = new System.Drawing.Size(49, 29);
            this.selectAllBtn.TabIndex = 0;
            this.selectAllBtn.Text = "全选";
            this.selectAllBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.selectAllBtn.Click += new System.EventHandler(this.selectAllBtn_Click);
            // 
            // NoSeachBtn
            // 
            this.NoSeachBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.NoSeachBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.NoSeachBtn.Location = new System.Drawing.Point(143, 2);
            this.NoSeachBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.NoSeachBtn.Name = "NoSeachBtn";
            this.NoSeachBtn.Size = new System.Drawing.Size(87, 31);
            this.NoSeachBtn.TabIndex = 1;
            this.NoSeachBtn.Text = "全不选";
            this.NoSeachBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.NoSeachBtn.Click += new System.EventHandler(this.NoSeachBtn_Click);
            // 
            // P1Sun
            // 
            this.P1Sun.Cursor = System.Windows.Forms.Cursors.Hand;
            this.P1Sun.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.P1Sun.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.P1Sun.Location = new System.Drawing.Point(13, 99);
            this.P1Sun.MinimumSize = new System.Drawing.Size(1, 1);
            this.P1Sun.Name = "P1Sun";
            this.P1Sun.Size = new System.Drawing.Size(150, 29);
            this.P1Sun.TabIndex = 37;
            this.P1Sun.Text = "查找子目录";
            this.P1Sun.CheckedChanged += new System.EventHandler(this.P1Sun_CheckedChanged);
            // 
            // OpenTargetBtn
            // 
            this.OpenTargetBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.OpenTargetBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OpenTargetBtn.Location = new System.Drawing.Point(847, 499);
            this.OpenTargetBtn.Margin = new System.Windows.Forms.Padding(2);
            this.OpenTargetBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.OpenTargetBtn.Name = "OpenTargetBtn";
            this.OpenTargetBtn.Size = new System.Drawing.Size(127, 38);
            this.OpenTargetBtn.TabIndex = 35;
            this.OpenTargetBtn.Text = "打开目标文件夹";
            this.OpenTargetBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OpenTargetBtn.Click += new System.EventHandler(this.OpenTargetBtn_Click);
            // 
            // uiLabel2
            // 
            this.uiLabel2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.uiLabel2.Location = new System.Drawing.Point(13, 55);
            this.uiLabel2.Name = "uiLabel2";
            this.uiLabel2.Size = new System.Drawing.Size(132, 36);
            this.uiLabel2.TabIndex = 34;
            this.uiLabel2.Text = "目标目录：";
            this.uiLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // uiLabel1
            // 
            this.uiLabel1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.uiLabel1.Location = new System.Drawing.Point(13, 9);
            this.uiLabel1.Name = "uiLabel1";
            this.uiLabel1.Size = new System.Drawing.Size(132, 36);
            this.uiLabel1.TabIndex = 33;
            this.uiLabel1.Text = "转换目录：";
            this.uiLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // SeachtargetPathBtn
            // 
            this.SeachtargetPathBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SeachtargetPathBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SeachtargetPathBtn.Location = new System.Drawing.Point(874, 55);
            this.SeachtargetPathBtn.Margin = new System.Windows.Forms.Padding(2);
            this.SeachtargetPathBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.SeachtargetPathBtn.Name = "SeachtargetPathBtn";
            this.SeachtargetPathBtn.Size = new System.Drawing.Size(102, 38);
            this.SeachtargetPathBtn.TabIndex = 28;
            this.SeachtargetPathBtn.Text = "选择文件夹";
            this.SeachtargetPathBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SeachtargetPathBtn.Click += new System.EventHandler(this.SeachtargetPathBtn_Click);
            // 
            // TypeBox
            // 
            this.TypeBox.ColumnCount = 6;
            this.TypeBox.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TypeBox.HoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(236)))), ((int)(((byte)(255)))));
            this.TypeBox.Items.AddRange(new object[] {
            "DOC",
            "DOCX",
            "WPS"});
            this.TypeBox.Location = new System.Drawing.Point(10, 128);
            this.TypeBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.TypeBox.MinimumSize = new System.Drawing.Size(1, 1);
            this.TypeBox.Name = "TypeBox";
            this.TypeBox.Padding = new System.Windows.Forms.Padding(0, 32, 0, 0);
            this.TypeBox.SelectedIndexes = ((System.Collections.Generic.List<int>)(resources.GetObject("TypeBox.SelectedIndexes")));
            this.TypeBox.Size = new System.Drawing.Size(964, 81);
            this.TypeBox.TabIndex = 32;
            this.TypeBox.Text = "原始类型";
            this.TypeBox.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.TypeBox.ValueChanged += new Sunny.UI.UICheckBoxGroup.OnValueChanged(this.TypeBox_ValueChanged);
            // 
            // StartChangeBtn2
            // 
            this.StartChangeBtn2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.StartChangeBtn2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.StartChangeBtn2.Location = new System.Drawing.Point(741, 499);
            this.StartChangeBtn2.Margin = new System.Windows.Forms.Padding(2);
            this.StartChangeBtn2.MinimumSize = new System.Drawing.Size(1, 1);
            this.StartChangeBtn2.Name = "StartChangeBtn2";
            this.StartChangeBtn2.Size = new System.Drawing.Size(102, 38);
            this.StartChangeBtn2.TabIndex = 31;
            this.StartChangeBtn2.Text = "开始转换";
            this.StartChangeBtn2.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.StartChangeBtn2.Click += new System.EventHandler(this.StartChangeBtn2_Click);
            // 
            // uiProcessBar1
            // 
            this.uiProcessBar1.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(235)))), ((int)(((byte)(243)))), ((int)(((byte)(255)))));
            this.uiProcessBar1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiProcessBar1.Location = new System.Drawing.Point(10, 508);
            this.uiProcessBar1.Margin = new System.Windows.Forms.Padding(2);
            this.uiProcessBar1.MinimumSize = new System.Drawing.Size(2, 2);
            this.uiProcessBar1.Name = "uiProcessBar1";
            this.uiProcessBar1.Size = new System.Drawing.Size(727, 23);
            this.uiProcessBar1.TabIndex = 29;
            this.uiProcessBar1.Text = "uiProcessBar1";
            this.uiProcessBar1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // BtnSeachWJJ2
            // 
            this.BtnSeachWJJ2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BtnSeachWJJ2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BtnSeachWJJ2.Location = new System.Drawing.Point(874, 9);
            this.BtnSeachWJJ2.Margin = new System.Windows.Forms.Padding(2);
            this.BtnSeachWJJ2.MinimumSize = new System.Drawing.Size(1, 1);
            this.BtnSeachWJJ2.Name = "BtnSeachWJJ2";
            this.BtnSeachWJJ2.Size = new System.Drawing.Size(102, 38);
            this.BtnSeachWJJ2.TabIndex = 26;
            this.BtnSeachWJJ2.Text = "选择文件夹";
            this.BtnSeachWJJ2.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BtnSeachWJJ2.Click += new System.EventHandler(this.BtnSeachWJJ2_Click);
            // 
            // p2
            // 
            this.p2.Controls.Add(this.ChangeText);
            this.p2.Controls.Add(this.P2Logs);
            this.p2.Controls.Add(this.P2Sun);
            this.p2.Controls.Add(this.Replacetype);
            this.p2.Controls.Add(this.uiCheckBoxGroup1);
            this.p2.Controls.Add(this.uiLabel3);
            this.p2.Controls.Add(this.SeachBtn);
            this.p2.Location = new System.Drawing.Point(201, 0);
            this.p2.Name = "p2";
            this.p2.Size = new System.Drawing.Size(985, 565);
            this.p2.TabIndex = 1;
            this.p2.Text = "批量修改文件名称";
            this.p2.UseVisualStyleBackColor = true;
            // 
            // ChangeText
            // 
            this.ChangeText.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.ChangeText.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ChangeText.Location = new System.Drawing.Point(116, 9);
            this.ChangeText.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ChangeText.MinimumSize = new System.Drawing.Size(1, 13);
            this.ChangeText.Name = "ChangeText";
            this.ChangeText.Padding = new System.Windows.Forms.Padding(4);
            this.ChangeText.ShowText = false;
            this.ChangeText.Size = new System.Drawing.Size(758, 38);
            this.ChangeText.TabIndex = 27;
            this.ChangeText.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.ChangeText.Watermark = "";
            // 
            // P2Logs
            // 
            this.P2Logs.FillColor = System.Drawing.Color.White;
            this.P2Logs.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.P2Logs.Location = new System.Drawing.Point(14, 375);
            this.P2Logs.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.P2Logs.MinimumSize = new System.Drawing.Size(1, 1);
            this.P2Logs.Name = "P2Logs";
            this.P2Logs.Padding = new System.Windows.Forms.Padding(2);
            this.P2Logs.ShowText = false;
            this.P2Logs.Size = new System.Drawing.Size(964, 159);
            this.P2Logs.TabIndex = 32;
            this.P2Logs.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // P2Sun
            // 
            this.P2Sun.Cursor = System.Windows.Forms.Cursors.Hand;
            this.P2Sun.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.P2Sun.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.P2Sun.Location = new System.Drawing.Point(14, 54);
            this.P2Sun.MinimumSize = new System.Drawing.Size(1, 1);
            this.P2Sun.Name = "P2Sun";
            this.P2Sun.Size = new System.Drawing.Size(150, 29);
            this.P2Sun.TabIndex = 31;
            this.P2Sun.Text = "查找子目录";
            // 
            // Replacetype
            // 
            this.Replacetype.ColumnCount = 6;
            this.Replacetype.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Replacetype.HoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(236)))), ((int)(((byte)(255)))));
            this.Replacetype.Items.AddRange(new object[] {
            "DOC",
            "DOCX",
            "PDF",
            "WPS",
            "XLS",
            "XLSX"});
            this.Replacetype.Location = new System.Drawing.Point(14, 89);
            this.Replacetype.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Replacetype.MinimumSize = new System.Drawing.Size(1, 1);
            this.Replacetype.Name = "Replacetype";
            this.Replacetype.Padding = new System.Windows.Forms.Padding(0, 32, 0, 0);
            this.Replacetype.SelectedIndexes = ((System.Collections.Generic.List<int>)(resources.GetObject("Replacetype.SelectedIndexes")));
            this.Replacetype.Size = new System.Drawing.Size(964, 67);
            this.Replacetype.TabIndex = 30;
            this.Replacetype.Text = "查找类型";
            this.Replacetype.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // uiCheckBoxGroup1
            // 
            this.uiCheckBoxGroup1.Controls.Add(this.P2Pro);
            this.uiCheckBoxGroup1.Controls.Add(this.ZZCHeck);
            this.uiCheckBoxGroup1.Controls.Add(this.ReplaceBtn);
            this.uiCheckBoxGroup1.Controls.Add(this.uiLabel5);
            this.uiCheckBoxGroup1.Controls.Add(this.ThText);
            this.uiCheckBoxGroup1.Controls.Add(this.uiLabel4);
            this.uiCheckBoxGroup1.Controls.Add(this.FindText);
            this.uiCheckBoxGroup1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiCheckBoxGroup1.HoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(236)))), ((int)(((byte)(255)))));
            this.uiCheckBoxGroup1.Location = new System.Drawing.Point(14, 212);
            this.uiCheckBoxGroup1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.uiCheckBoxGroup1.MinimumSize = new System.Drawing.Size(1, 1);
            this.uiCheckBoxGroup1.Name = "uiCheckBoxGroup1";
            this.uiCheckBoxGroup1.Padding = new System.Windows.Forms.Padding(0, 32, 0, 0);
            this.uiCheckBoxGroup1.SelectedIndexes = ((System.Collections.Generic.List<int>)(resources.GetObject("uiCheckBoxGroup1.SelectedIndexes")));
            this.uiCheckBoxGroup1.Size = new System.Drawing.Size(964, 153);
            this.uiCheckBoxGroup1.TabIndex = 29;
            this.uiCheckBoxGroup1.Text = "替换";
            this.uiCheckBoxGroup1.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // P2Pro
            // 
            this.P2Pro.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(235)))), ((int)(((byte)(243)))), ((int)(((byte)(255)))));
            this.P2Pro.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.P2Pro.Location = new System.Drawing.Point(177, 113);
            this.P2Pro.MinimumSize = new System.Drawing.Size(3, 3);
            this.P2Pro.Name = "P2Pro";
            this.P2Pro.Size = new System.Drawing.Size(719, 29);
            this.P2Pro.TabIndex = 6;
            // 
            // ZZCHeck
            // 
            this.ZZCHeck.Cursor = System.Windows.Forms.Cursors.Hand;
            this.ZZCHeck.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ZZCHeck.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.ZZCHeck.Location = new System.Drawing.Point(21, 113);
            this.ZZCHeck.MinimumSize = new System.Drawing.Size(1, 1);
            this.ZZCHeck.Name = "ZZCHeck";
            this.ZZCHeck.Size = new System.Drawing.Size(150, 29);
            this.ZZCHeck.TabIndex = 5;
            this.ZZCHeck.Text = "正则表达式";
            // 
            // ReplaceBtn
            // 
            this.ReplaceBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.ReplaceBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ReplaceBtn.Location = new System.Drawing.Point(868, 36);
            this.ReplaceBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.ReplaceBtn.Name = "ReplaceBtn";
            this.ReplaceBtn.Size = new System.Drawing.Size(85, 69);
            this.ReplaceBtn.TabIndex = 4;
            this.ReplaceBtn.Text = "替换";
            this.ReplaceBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ReplaceBtn.Click += new System.EventHandler(this.ReplaceBtn_Click);
            // 
            // uiLabel5
            // 
            this.uiLabel5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.uiLabel5.Location = new System.Drawing.Point(3, 77);
            this.uiLabel5.Name = "uiLabel5";
            this.uiLabel5.Size = new System.Drawing.Size(92, 23);
            this.uiLabel5.TabIndex = 3;
            this.uiLabel5.Text = "替换内容：";
            this.uiLabel5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ThText
            // 
            this.ThText.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.ThText.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ThText.Location = new System.Drawing.Point(102, 76);
            this.ThText.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.ThText.MinimumSize = new System.Drawing.Size(1, 16);
            this.ThText.Name = "ThText";
            this.ThText.Padding = new System.Windows.Forms.Padding(5);
            this.ThText.ShowText = false;
            this.ThText.Size = new System.Drawing.Size(758, 29);
            this.ThText.TabIndex = 2;
            this.ThText.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.ThText.Watermark = "";
            // 
            // uiLabel4
            // 
            this.uiLabel4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.uiLabel4.Location = new System.Drawing.Point(3, 38);
            this.uiLabel4.Name = "uiLabel4";
            this.uiLabel4.Size = new System.Drawing.Size(92, 23);
            this.uiLabel4.TabIndex = 1;
            this.uiLabel4.Text = "查找内容：";
            this.uiLabel4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // FindText
            // 
            this.FindText.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.FindText.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FindText.Location = new System.Drawing.Point(102, 37);
            this.FindText.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.FindText.MinimumSize = new System.Drawing.Size(1, 16);
            this.FindText.Name = "FindText";
            this.FindText.Padding = new System.Windows.Forms.Padding(5);
            this.FindText.ShowText = false;
            this.FindText.Size = new System.Drawing.Size(758, 29);
            this.FindText.TabIndex = 0;
            this.FindText.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.FindText.Watermark = "";
            // 
            // uiLabel3
            // 
            this.uiLabel3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.uiLabel3.Location = new System.Drawing.Point(11, 11);
            this.uiLabel3.Name = "uiLabel3";
            this.uiLabel3.Size = new System.Drawing.Size(110, 36);
            this.uiLabel3.TabIndex = 28;
            this.uiLabel3.Text = "修改目录：";
            this.uiLabel3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // SeachBtn
            // 
            this.SeachBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SeachBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SeachBtn.Location = new System.Drawing.Point(879, 9);
            this.SeachBtn.Margin = new System.Windows.Forms.Padding(2);
            this.SeachBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.SeachBtn.Name = "SeachBtn";
            this.SeachBtn.Size = new System.Drawing.Size(99, 38);
            this.SeachBtn.TabIndex = 26;
            this.SeachBtn.Text = "选择文件夹";
            this.SeachBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SeachBtn.Click += new System.EventHandler(this.SeachBtn_Click);
            // 
            // p3
            // 
            this.p3.Controls.Add(this.CollectFilesFindfilePath);
            this.p3.Controls.Add(this.CollectFilesSavePath);
            this.p3.Controls.Add(this.CollectFilesLogs);
            this.p3.Controls.Add(this.CollectFilesFindSunFIles);
            this.p3.Controls.Add(this.P3Pro);
            this.p3.Controls.Add(this.CollectFilesOpenSaveBtn);
            this.p3.Controls.Add(this.CollectFilesStartBtn);
            this.p3.Controls.Add(this.CollectfFilesType);
            this.p3.Controls.Add(this.uiLabel6);
            this.p3.Controls.Add(this.uiLabel7);
            this.p3.Controls.Add(this.CollectFilesSavePathBtn);
            this.p3.Controls.Add(this.CollectFilesSeachFindPathBtn);
            this.p3.Location = new System.Drawing.Point(601, 0);
            this.p3.Name = "p3";
            this.p3.Size = new System.Drawing.Size(0, 60);
            this.p3.TabIndex = 2;
            this.p3.Text = "收集文件";
            this.p3.UseVisualStyleBackColor = true;
            // 
            // CollectFilesFindfilePath
            // 
            this.CollectFilesFindfilePath.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.CollectFilesFindfilePath.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesFindfilePath.Location = new System.Drawing.Point(117, 4);
            this.CollectFilesFindfilePath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.CollectFilesFindfilePath.MinimumSize = new System.Drawing.Size(1, 13);
            this.CollectFilesFindfilePath.Name = "CollectFilesFindfilePath";
            this.CollectFilesFindfilePath.Padding = new System.Windows.Forms.Padding(4);
            this.CollectFilesFindfilePath.ShowText = false;
            this.CollectFilesFindfilePath.Size = new System.Drawing.Size(751, 38);
            this.CollectFilesFindfilePath.TabIndex = 26;
            this.CollectFilesFindfilePath.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.CollectFilesFindfilePath.Watermark = "";
            // 
            // CollectFilesSavePath
            // 
            this.CollectFilesSavePath.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.CollectFilesSavePath.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesSavePath.Location = new System.Drawing.Point(117, 50);
            this.CollectFilesSavePath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.CollectFilesSavePath.MinimumSize = new System.Drawing.Size(1, 13);
            this.CollectFilesSavePath.Name = "CollectFilesSavePath";
            this.CollectFilesSavePath.Padding = new System.Windows.Forms.Padding(4);
            this.CollectFilesSavePath.ShowText = false;
            this.CollectFilesSavePath.Size = new System.Drawing.Size(751, 38);
            this.CollectFilesSavePath.TabIndex = 28;
            this.CollectFilesSavePath.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.CollectFilesSavePath.Watermark = "";
            // 
            // CollectFilesLogs
            // 
            this.CollectFilesLogs.FillColor = System.Drawing.Color.White;
            this.CollectFilesLogs.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesLogs.Location = new System.Drawing.Point(13, 290);
            this.CollectFilesLogs.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.CollectFilesLogs.MinimumSize = new System.Drawing.Size(1, 1);
            this.CollectFilesLogs.Name = "CollectFilesLogs";
            this.CollectFilesLogs.Padding = new System.Windows.Forms.Padding(2);
            this.CollectFilesLogs.ShowText = false;
            this.CollectFilesLogs.Size = new System.Drawing.Size(961, 246);
            this.CollectFilesLogs.TabIndex = 36;
            this.CollectFilesLogs.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CollectFilesFindSunFIles
            // 
            this.CollectFilesFindSunFIles.Cursor = System.Windows.Forms.Cursors.Hand;
            this.CollectFilesFindSunFIles.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesFindSunFIles.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.CollectFilesFindSunFIles.Location = new System.Drawing.Point(21, 95);
            this.CollectFilesFindSunFIles.MinimumSize = new System.Drawing.Size(1, 1);
            this.CollectFilesFindSunFIles.Name = "CollectFilesFindSunFIles";
            this.CollectFilesFindSunFIles.Size = new System.Drawing.Size(150, 29);
            this.CollectFilesFindSunFIles.TabIndex = 35;
            this.CollectFilesFindSunFIles.Text = "查找子目录";
            // 
            // P3Pro
            // 
            this.P3Pro.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(235)))), ((int)(((byte)(243)))), ((int)(((byte)(255)))));
            this.P3Pro.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.P3Pro.Location = new System.Drawing.Point(13, 257);
            this.P3Pro.MinimumSize = new System.Drawing.Size(3, 3);
            this.P3Pro.Name = "P3Pro";
            this.P3Pro.Size = new System.Drawing.Size(727, 22);
            this.P3Pro.TabIndex = 34;
            this.P3Pro.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // CollectFilesOpenSaveBtn
            // 
            this.CollectFilesOpenSaveBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.CollectFilesOpenSaveBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesOpenSaveBtn.Location = new System.Drawing.Point(847, 247);
            this.CollectFilesOpenSaveBtn.Margin = new System.Windows.Forms.Padding(2);
            this.CollectFilesOpenSaveBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.CollectFilesOpenSaveBtn.Name = "CollectFilesOpenSaveBtn";
            this.CollectFilesOpenSaveBtn.Size = new System.Drawing.Size(127, 38);
            this.CollectFilesOpenSaveBtn.TabIndex = 33;
            this.CollectFilesOpenSaveBtn.Text = "打开存放目录";
            this.CollectFilesOpenSaveBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesOpenSaveBtn.Click += new System.EventHandler(this.CollectFilesOpenSaveBtn_Click);
            // 
            // CollectFilesStartBtn
            // 
            this.CollectFilesStartBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.CollectFilesStartBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesStartBtn.Location = new System.Drawing.Point(745, 247);
            this.CollectFilesStartBtn.Margin = new System.Windows.Forms.Padding(2);
            this.CollectFilesStartBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.CollectFilesStartBtn.Name = "CollectFilesStartBtn";
            this.CollectFilesStartBtn.Size = new System.Drawing.Size(98, 38);
            this.CollectFilesStartBtn.TabIndex = 32;
            this.CollectFilesStartBtn.Text = "开始收集";
            this.CollectFilesStartBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesStartBtn.Click += new System.EventHandler(this.CollectFilesStartBtn_Click);
            // 
            // CollectfFilesType
            // 
            this.CollectfFilesType.ColumnCount = 6;
            this.CollectfFilesType.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectfFilesType.HoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(236)))), ((int)(((byte)(255)))));
            this.CollectfFilesType.Items.AddRange(new object[] {
            "DOC",
            "DOCX",
            "PDF",
            "WPS",
            "XLS",
            "XLSX",
            "TXT",
            "BAT",
            "PNG",
            "JPEG",
            "LRMX"});
            this.CollectfFilesType.Location = new System.Drawing.Point(13, 127);
            this.CollectfFilesType.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.CollectfFilesType.MinimumSize = new System.Drawing.Size(1, 1);
            this.CollectfFilesType.Name = "CollectfFilesType";
            this.CollectfFilesType.Padding = new System.Windows.Forms.Padding(0, 32, 0, 0);
            this.CollectfFilesType.SelectedIndexes = ((System.Collections.Generic.List<int>)(resources.GetObject("CollectfFilesType.SelectedIndexes")));
            this.CollectfFilesType.Size = new System.Drawing.Size(960, 114);
            this.CollectfFilesType.TabIndex = 31;
            this.CollectfFilesType.Text = "查找类型";
            this.CollectfFilesType.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // uiLabel6
            // 
            this.uiLabel6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.uiLabel6.Location = new System.Drawing.Point(11, 50);
            this.uiLabel6.Name = "uiLabel6";
            this.uiLabel6.Size = new System.Drawing.Size(121, 36);
            this.uiLabel6.TabIndex = 30;
            this.uiLabel6.Text = "存放目录：";
            this.uiLabel6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // uiLabel7
            // 
            this.uiLabel7.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.uiLabel7.Location = new System.Drawing.Point(11, 4);
            this.uiLabel7.Name = "uiLabel7";
            this.uiLabel7.Size = new System.Drawing.Size(121, 36);
            this.uiLabel7.TabIndex = 29;
            this.uiLabel7.Text = "查找目录：";
            this.uiLabel7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CollectFilesSavePathBtn
            // 
            this.CollectFilesSavePathBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.CollectFilesSavePathBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesSavePathBtn.Location = new System.Drawing.Point(871, 50);
            this.CollectFilesSavePathBtn.Margin = new System.Windows.Forms.Padding(2);
            this.CollectFilesSavePathBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.CollectFilesSavePathBtn.Name = "CollectFilesSavePathBtn";
            this.CollectFilesSavePathBtn.Size = new System.Drawing.Size(102, 38);
            this.CollectFilesSavePathBtn.TabIndex = 27;
            this.CollectFilesSavePathBtn.Text = "选择文件夹";
            this.CollectFilesSavePathBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesSavePathBtn.Click += new System.EventHandler(this.CollectFilesSavePathBtn_Click);
            // 
            // CollectFilesSeachFindPathBtn
            // 
            this.CollectFilesSeachFindPathBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.CollectFilesSeachFindPathBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesSeachFindPathBtn.Location = new System.Drawing.Point(873, 4);
            this.CollectFilesSeachFindPathBtn.Margin = new System.Windows.Forms.Padding(2);
            this.CollectFilesSeachFindPathBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.CollectFilesSeachFindPathBtn.Name = "CollectFilesSeachFindPathBtn";
            this.CollectFilesSeachFindPathBtn.Size = new System.Drawing.Size(102, 38);
            this.CollectFilesSeachFindPathBtn.TabIndex = 25;
            this.CollectFilesSeachFindPathBtn.Text = "选择文件夹";
            this.CollectFilesSeachFindPathBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CollectFilesSeachFindPathBtn.Click += new System.EventHandler(this.CollectFilesSeachFindPathBtn_Click);
            // 
            // p4
            // 
            this.p4.Controls.Add(this.p4_waitbar);
            this.p4.Controls.Add(this.p4_PdfFilePath);
            this.p4.Controls.Add(this.uiLabel8);
            this.p4.Controls.Add(this.p4_SeachFileBtn);
            this.p4.Controls.Add(this.p4_StartBtn);
            this.p4.Location = new System.Drawing.Point(801, 0);
            this.p4.Name = "p4";
            this.p4.Size = new System.Drawing.Size(0, 60);
            this.p4.TabIndex = 3;
            this.p4.Text = "PDF文件操作";
            this.p4.UseVisualStyleBackColor = true;
            // 
            // p4_waitbar
            // 
            this.p4_waitbar.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(249)))), ((int)(((byte)(255)))));
            this.p4_waitbar.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.p4_waitbar.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.p4_waitbar.Location = new System.Drawing.Point(566, 55);
            this.p4_waitbar.MinimumSize = new System.Drawing.Size(70, 23);
            this.p4_waitbar.Name = "p4_waitbar";
            this.p4_waitbar.Size = new System.Drawing.Size(300, 29);
            this.p4_waitbar.TabIndex = 37;
            this.p4_waitbar.Text = "uiWaitingBar1";
            this.p4_waitbar.Visible = false;
            // 
            // p4_PdfFilePath
            // 
            this.p4_PdfFilePath.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.p4_PdfFilePath.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.p4_PdfFilePath.Location = new System.Drawing.Point(126, 4);
            this.p4_PdfFilePath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.p4_PdfFilePath.MinimumSize = new System.Drawing.Size(1, 13);
            this.p4_PdfFilePath.Name = "p4_PdfFilePath";
            this.p4_PdfFilePath.Padding = new System.Windows.Forms.Padding(4);
            this.p4_PdfFilePath.ShowText = false;
            this.p4_PdfFilePath.Size = new System.Drawing.Size(741, 38);
            this.p4_PdfFilePath.TabIndex = 35;
            this.p4_PdfFilePath.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.p4_PdfFilePath.Watermark = "";
            // 
            // uiLabel8
            // 
            this.uiLabel8.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel8.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.uiLabel8.Location = new System.Drawing.Point(11, 4);
            this.uiLabel8.Name = "uiLabel8";
            this.uiLabel8.Size = new System.Drawing.Size(132, 36);
            this.uiLabel8.TabIndex = 36;
            this.uiLabel8.Text = "选择pdf文件：";
            this.uiLabel8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // p4_SeachFileBtn
            // 
            this.p4_SeachFileBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.p4_SeachFileBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.p4_SeachFileBtn.Location = new System.Drawing.Point(872, 4);
            this.p4_SeachFileBtn.Margin = new System.Windows.Forms.Padding(2);
            this.p4_SeachFileBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.p4_SeachFileBtn.Name = "p4_SeachFileBtn";
            this.p4_SeachFileBtn.Size = new System.Drawing.Size(102, 38);
            this.p4_SeachFileBtn.TabIndex = 34;
            this.p4_SeachFileBtn.Text = "选择文件夹";
            this.p4_SeachFileBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.p4_SeachFileBtn.Click += new System.EventHandler(this.p4_SeachFileBtn_Click);
            // 
            // p4_StartBtn
            // 
            this.p4_StartBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.p4_StartBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.p4_StartBtn.Location = new System.Drawing.Point(872, 55);
            this.p4_StartBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.p4_StartBtn.Name = "p4_StartBtn";
            this.p4_StartBtn.Size = new System.Drawing.Size(100, 35);
            this.p4_StartBtn.TabIndex = 8;
            this.p4_StartBtn.Text = "开始转换";
            this.p4_StartBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.p4_StartBtn.Click += new System.EventHandler(this.p4_StartBtn_Click);
            // 
            // uiStyleManager1
            // 
            this.uiStyleManager1.DPIScale = true;
            // 
            // uiTabControl1
            // 
            this.uiTabControl1.AllowDrop = true;
            this.uiTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.uiTabControl1.Controls.Add(this.FileOperations);
            this.uiTabControl1.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.uiTabControl1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiTabControl1.ItemSize = new System.Drawing.Size(150, 40);
            this.uiTabControl1.Location = new System.Drawing.Point(0, 35);
            this.uiTabControl1.MainPage = "";
            this.uiTabControl1.Name = "uiTabControl1";
            this.uiTabControl1.Padding = new System.Drawing.Point(0, 0);
            this.uiTabControl1.SelectedIndex = 0;
            this.uiTabControl1.Size = new System.Drawing.Size(1186, 605);
            this.uiTabControl1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.uiTabControl1.TabIndex = 2;
            this.uiTabControl1.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            // 
            // FileOperations
            // 
            this.FileOperations.Controls.Add(this.uiTabControlMenu1);
            this.FileOperations.Location = new System.Drawing.Point(0, 40);
            this.FileOperations.Name = "FileOperations";
            this.FileOperations.Size = new System.Drawing.Size(1186, 565);
            this.FileOperations.TabIndex = 0;
            this.FileOperations.Text = "文件操作";
            this.FileOperations.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(1186, 640);
            this.Controls.Add(this.uiTabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.RectColor = System.Drawing.Color.DeepSkyBlue;
            this.Text = "工具";
            this.TitleColor = System.Drawing.Color.Silver;
            this.ZoomScaleRect = new System.Drawing.Rectangle(15, 15, 1186, 544);
            this.uiTabControlMenu1.ResumeLayout(false);
            this.p1.ResumeLayout(false);
            this.FileList.ResumeLayout(false);
            this.p2.ResumeLayout(false);
            this.uiCheckBoxGroup1.ResumeLayout(false);
            this.p3.ResumeLayout(false);
            this.p4.ResumeLayout(false);
            this.uiTabControl1.ResumeLayout(false);
            this.FileOperations.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private Sunny.UI.UITabControlMenu uiTabControlMenu1;
        private TabPage p1;
        private Sunny.UI.UICheckBox P1Sun;
        private Sunny.UI.UIButton NoSeachBtn;
        private Sunny.UI.UIButton selectAllBtn;
        private Sunny.UI.UIButton OpenTargetBtn;
        private Sunny.UI.UILabel uiLabel2;
        private Sunny.UI.UILabel uiLabel1;
        private Sunny.UI.UITextBox targetPath;
        private Sunny.UI.UIButton SeachtargetPathBtn;
        private Sunny.UI.UICheckBoxGroup TypeBox;
        private Sunny.UI.UIButton StartChangeBtn2;
        private Sunny.UI.UIProcessBar uiProcessBar1;
        private Sunny.UI.UITextBox folderPath2;
        private Sunny.UI.UIButton BtnSeachWJJ2;
        private TabPage p2;
        private Sunny.UI.UIRichTextBox P2Logs;
        private Sunny.UI.UICheckBox P2Sun;
        private Sunny.UI.UICheckBoxGroup Replacetype;
        private Sunny.UI.UICheckBoxGroup uiCheckBoxGroup1;
        private Sunny.UI.UIProcessBar P2Pro;
        private Sunny.UI.UICheckBox ZZCHeck;
        private Sunny.UI.UIButton ReplaceBtn;
        private Sunny.UI.UILabel uiLabel5;
        private Sunny.UI.UITextBox ThText;
        private Sunny.UI.UILabel uiLabel4;
        private Sunny.UI.UITextBox FindText;
        private Sunny.UI.UILabel uiLabel3;
        private Sunny.UI.UITextBox ChangeText;
        private Sunny.UI.UIButton SeachBtn;
        private TabPage p3;
        private Sunny.UI.UIRichTextBox CollectFilesLogs;
        private Sunny.UI.UICheckBox CollectFilesFindSunFIles;
        private Sunny.UI.UIProcessBar P3Pro;
        private Sunny.UI.UIButton CollectFilesOpenSaveBtn;
        private Sunny.UI.UIButton CollectFilesStartBtn;
        private Sunny.UI.UICheckBoxGroup CollectfFilesType;
        private Sunny.UI.UILabel uiLabel6;
        private Sunny.UI.UILabel uiLabel7;
        private Sunny.UI.UITextBox CollectFilesSavePath;
        private Sunny.UI.UIButton CollectFilesSavePathBtn;
        private Sunny.UI.UITextBox CollectFilesFindfilePath;
        private Sunny.UI.UIButton CollectFilesSeachFindPathBtn;
        private Sunny.UI.UIStyleManager uiStyleManager1;
        private Sunny.UI.UICheckBoxGroup FileList;
        private TabPage p4;
        private Sunny.UI.UIButton p4_StartBtn;
        private Sunny.UI.UITextBox p4_PdfFilePath;
        private Sunny.UI.UILabel uiLabel8;
        private Sunny.UI.UIButton p4_SeachFileBtn;
        private Sunny.UI.UIWaitingBar p4_waitbar;
        private Sunny.UI.UINavMenu uiNavMenu1;
        private Sunny.UI.UITabControl uiTabControl1;
        private TabPage FileOperations;
    }
}
