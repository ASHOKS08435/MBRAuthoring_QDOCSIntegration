
namespace MBRAuthoring.Presentation
{
    partial class FormFileOpenWizard
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.cmdCancel = new System.Windows.Forms.Button();
            this.cmdBack = new System.Windows.Forms.Button();
            this.cmdNext = new System.Windows.Forms.Button();
            this.cmdFinish = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.lblFileOpen = new System.Windows.Forms.Label();
            this.txtFindString = new System.Windows.Forms.TextBox();
            this.lblFileName = new System.Windows.Forms.Label();
            this.chkFindProcessingSteps = new System.Windows.Forms.CheckBox();
            this.lblProject = new System.Windows.Forms.Label();
            this.cboSelectFile = new System.Windows.Forms.ComboBox();
            this.lblRecordType = new System.Windows.Forms.Label();
            this.cboProject = new System.Windows.Forms.ComboBox();
            this.lblOpUnit = new System.Windows.Forms.Label();
            this.cboDocType = new System.Windows.Forms.ComboBox();
            this.cboOpUnit = new System.Windows.Forms.ComboBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lblUpdated = new System.Windows.Forms.Label();
            this.cmdUpdateDocTitle = new System.Windows.Forms.Button();
            this.txtDocProductName = new System.Windows.Forms.TextBox();
            this.cmdCopyFileName = new System.Windows.Forms.Button();
            this.lblClickTo = new System.Windows.Forms.Label();
            this.txtDocProductID = new System.Windows.Forms.TextBox();
            this.cmdCopyDocTitle = new System.Windows.Forms.Button();
            this.txtDocTitle = new System.Windows.Forms.TextBox();
            this.cmdCopyProductName = new System.Windows.Forms.Button();
            this.txtDocFileName = new System.Windows.Forms.TextBox();
            this.cmdCopyProductID = new System.Windows.Forms.Button();
            this.lblWizardInfo = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label0 = new System.Windows.Forms.Label();
            this.shp1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.shp0 = new System.Windows.Forms.Label();
            this.mapStatusStart = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmdCancel
            // 
            this.cmdCancel.BackColor = System.Drawing.SystemColors.Control;
            this.cmdCancel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdCancel.Location = new System.Drawing.Point(254, 318);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(75, 23);
            this.cmdCancel.TabIndex = 12;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = false;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // cmdBack
            // 
            this.cmdBack.BackColor = System.Drawing.SystemColors.Control;
            this.cmdBack.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdBack.Location = new System.Drawing.Point(348, 318);
            this.cmdBack.Name = "cmdBack";
            this.cmdBack.Size = new System.Drawing.Size(75, 23);
            this.cmdBack.TabIndex = 13;
            this.cmdBack.Text = "< Back";
            this.cmdBack.UseVisualStyleBackColor = false;
            this.cmdBack.Click += new System.EventHandler(this.cmdBack_Click);
            // 
            // cmdNext
            // 
            this.cmdNext.BackColor = System.Drawing.SystemColors.Control;
            this.cmdNext.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdNext.Location = new System.Drawing.Point(429, 318);
            this.cmdNext.Name = "cmdNext";
            this.cmdNext.Size = new System.Drawing.Size(75, 23);
            this.cmdNext.TabIndex = 14;
            this.cmdNext.Text = "Next >";
            this.cmdNext.UseVisualStyleBackColor = false;
            this.cmdNext.Click += new System.EventHandler(this.cmdNext_Click);
            // 
            // cmdFinish
            // 
            this.cmdFinish.BackColor = System.Drawing.SystemColors.Control;
            this.cmdFinish.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdFinish.Location = new System.Drawing.Point(520, 318);
            this.cmdFinish.Name = "cmdFinish";
            this.cmdFinish.Size = new System.Drawing.Size(75, 23);
            this.cmdFinish.TabIndex = 15;
            this.cmdFinish.Text = "Open";
            this.cmdFinish.UseVisualStyleBackColor = false;
            this.cmdFinish.Click += new System.EventHandler(this.cmdFinish_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(139, -5);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(467, 287);
            this.tabControl1.TabIndex = 10;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.lblFileOpen);
            this.tabPage1.Controls.Add(this.txtFindString);
            this.tabPage1.Controls.Add(this.lblFileName);
            this.tabPage1.Controls.Add(this.chkFindProcessingSteps);
            this.tabPage1.Controls.Add(this.lblProject);
            this.tabPage1.Controls.Add(this.cboSelectFile);
            this.tabPage1.Controls.Add(this.lblRecordType);
            this.tabPage1.Controls.Add(this.cboProject);
            this.tabPage1.Controls.Add(this.lblOpUnit);
            this.tabPage1.Controls.Add(this.cboDocType);
            this.tabPage1.Controls.Add(this.cboOpUnit);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(459, 261);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "tabPage1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // lblFileOpen
            // 
            this.lblFileOpen.AutoSize = true;
            this.lblFileOpen.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileOpen.Location = new System.Drawing.Point(18, 15);
            this.lblFileOpen.Name = "lblFileOpen";
            this.lblFileOpen.Size = new System.Drawing.Size(69, 16);
            this.lblFileOpen.TabIndex = 0;
            this.lblFileOpen.Text = "File Open ";
            // 
            // txtFindString
            // 
            this.txtFindString.Location = new System.Drawing.Point(237, 221);
            this.txtFindString.Name = "txtFindString";
            this.txtFindString.Size = new System.Drawing.Size(216, 20);
            this.txtFindString.TabIndex = 10;
            // 
            // lblFileName
            // 
            this.lblFileName.AutoSize = true;
            this.lblFileName.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileName.Location = new System.Drawing.Point(23, 165);
            this.lblFileName.Name = "lblFileName";
            this.lblFileName.Size = new System.Drawing.Size(60, 13);
            this.lblFileName.TabIndex = 4;
            this.lblFileName.Text = "File Name :";
            // 
            // chkFindProcessingSteps
            // 
            this.chkFindProcessingSteps.AutoSize = true;
            this.chkFindProcessingSteps.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkFindProcessingSteps.Location = new System.Drawing.Point(120, 224);
            this.chkFindProcessingSteps.Name = "chkFindProcessingSteps";
            this.chkFindProcessingSteps.Size = new System.Drawing.Size(111, 17);
            this.chkFindProcessingSteps.TabIndex = 9;
            this.chkFindProcessingSteps.Text = "After Open, Find:";
            this.chkFindProcessingSteps.UseVisualStyleBackColor = true;
            // 
            // lblProject
            // 
            this.lblProject.AutoSize = true;
            this.lblProject.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblProject.Location = new System.Drawing.Point(23, 128);
            this.lblProject.Name = "lblProject";
            this.lblProject.Size = new System.Drawing.Size(48, 13);
            this.lblProject.TabIndex = 3;
            this.lblProject.Text = "Project :";
            // 
            // cboSelectFile
            // 
            this.cboSelectFile.DropDownWidth = 520;
            this.cboSelectFile.FormattingEnabled = true;
            this.cboSelectFile.IntegralHeight = false;
            this.cboSelectFile.Location = new System.Drawing.Point(21, 186);
            this.cboSelectFile.Name = "cboSelectFile";
            this.cboSelectFile.Size = new System.Drawing.Size(432, 21);
            this.cboSelectFile.TabIndex = 8;
            // 
            // lblRecordType
            // 
            this.lblRecordType.AutoSize = true;
            this.lblRecordType.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRecordType.Location = new System.Drawing.Point(21, 86);
            this.lblRecordType.Name = "lblRecordType";
            this.lblRecordType.Size = new System.Drawing.Size(89, 13);
            this.lblRecordType.TabIndex = 2;
            this.lblRecordType.Text = "Document Type :";
            // 
            // cboProject
            // 
            this.cboProject.FormattingEnabled = true;
            this.cboProject.Location = new System.Drawing.Point(120, 120);
            this.cboProject.Name = "cboProject";
            this.cboProject.Size = new System.Drawing.Size(182, 21);
            this.cboProject.TabIndex = 6;
            this.cboProject.SelectedIndexChanged += new System.EventHandler(this.cboProject_SelectedIndexChanged);
            // 
            // lblOpUnit
            // 
            this.lblOpUnit.AutoSize = true;
            this.lblOpUnit.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOpUnit.Location = new System.Drawing.Point(20, 44);
            this.lblOpUnit.Name = "lblOpUnit";
            this.lblOpUnit.Size = new System.Drawing.Size(84, 13);
            this.lblOpUnit.TabIndex = 1;
            this.lblOpUnit.Text = "Operating Unit :";
            // 
            // cboDocType
            // 
            this.cboDocType.FormattingEnabled = true;
            this.cboDocType.Location = new System.Drawing.Point(120, 78);
            this.cboDocType.Name = "cboDocType";
            this.cboDocType.Size = new System.Drawing.Size(182, 21);
            this.cboDocType.TabIndex = 7;
            this.cboDocType.SelectedIndexChanged += new System.EventHandler(this.cboDocType_SelectedIndexChanged);
            // 
            // cboOpUnit
            // 
            this.cboOpUnit.FormattingEnabled = true;
            this.cboOpUnit.Location = new System.Drawing.Point(120, 36);
            this.cboOpUnit.Name = "cboOpUnit";
            this.cboOpUnit.Size = new System.Drawing.Size(182, 21);
            this.cboOpUnit.TabIndex = 5;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lblUpdated);
            this.tabPage2.Controls.Add(this.cmdUpdateDocTitle);
            this.tabPage2.Controls.Add(this.txtDocProductName);
            this.tabPage2.Controls.Add(this.cmdCopyFileName);
            this.tabPage2.Controls.Add(this.lblClickTo);
            this.tabPage2.Controls.Add(this.txtDocProductID);
            this.tabPage2.Controls.Add(this.cmdCopyDocTitle);
            this.tabPage2.Controls.Add(this.txtDocTitle);
            this.tabPage2.Controls.Add(this.cmdCopyProductName);
            this.tabPage2.Controls.Add(this.txtDocFileName);
            this.tabPage2.Controls.Add(this.cmdCopyProductID);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(459, 261);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // lblUpdated
            // 
            this.lblUpdated.AutoSize = true;
            this.lblUpdated.Font = new System.Drawing.Font("Tahoma", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUpdated.Location = new System.Drawing.Point(225, 120);
            this.lblUpdated.Name = "lblUpdated";
            this.lblUpdated.Size = new System.Drawing.Size(0, 13);
            this.lblUpdated.TabIndex = 10;
            // 
            // cmdUpdateDocTitle
            // 
            this.cmdUpdateDocTitle.BackColor = System.Drawing.SystemColors.Control;
            this.cmdUpdateDocTitle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdUpdateDocTitle.Location = new System.Drawing.Point(144, 115);
            this.cmdUpdateDocTitle.Name = "cmdUpdateDocTitle";
            this.cmdUpdateDocTitle.Size = new System.Drawing.Size(75, 23);
            this.cmdUpdateDocTitle.TabIndex = 9;
            this.cmdUpdateDocTitle.Text = "&Update Title";
            this.cmdUpdateDocTitle.UseVisualStyleBackColor = false;
            this.cmdUpdateDocTitle.Click += new System.EventHandler(this.cmdUpdateDocTitle_Click);
            // 
            // txtDocProductName
            // 
            this.txtDocProductName.BackColor = System.Drawing.SystemColors.Window;
            this.txtDocProductName.Location = new System.Drawing.Point(144, 199);
            this.txtDocProductName.Multiline = true;
            this.txtDocProductName.Name = "txtDocProductName";
            this.txtDocProductName.ReadOnly = true;
            this.txtDocProductName.Size = new System.Drawing.Size(309, 25);
            this.txtDocProductName.TabIndex = 8;
            // 
            // cmdCopyFileName
            // 
            this.cmdCopyFileName.BackColor = System.Drawing.SystemColors.Control;
            this.cmdCopyFileName.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdCopyFileName.Location = new System.Drawing.Point(12, 33);
            this.cmdCopyFileName.Name = "cmdCopyFileName";
            this.cmdCopyFileName.Size = new System.Drawing.Size(121, 36);
            this.cmdCopyFileName.TabIndex = 0;
            this.cmdCopyFileName.Text = "F&ILE NAME :";
            this.cmdCopyFileName.UseVisualStyleBackColor = false;
            this.cmdCopyFileName.Click += new System.EventHandler(this.cmdCopyFileName_Click);
            // 
            // lblClickTo
            // 
            this.lblClickTo.AutoSize = true;
            this.lblClickTo.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblClickTo.Location = new System.Drawing.Point(17, 10);
            this.lblClickTo.Name = "lblClickTo";
            this.lblClickTo.Size = new System.Drawing.Size(96, 16);
            this.lblClickTo.TabIndex = 4;
            this.lblClickTo.Text = "Click to COPY...";
            // 
            // txtDocProductID
            // 
            this.txtDocProductID.BackColor = System.Drawing.SystemColors.Window;
            this.txtDocProductID.Location = new System.Drawing.Point(144, 147);
            this.txtDocProductID.Multiline = true;
            this.txtDocProductID.Name = "txtDocProductID";
            this.txtDocProductID.ReadOnly = true;
            this.txtDocProductID.Size = new System.Drawing.Size(309, 25);
            this.txtDocProductID.TabIndex = 7;
            // 
            // cmdCopyDocTitle
            // 
            this.cmdCopyDocTitle.BackColor = System.Drawing.SystemColors.Control;
            this.cmdCopyDocTitle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdCopyDocTitle.Location = new System.Drawing.Point(12, 89);
            this.cmdCopyDocTitle.Name = "cmdCopyDocTitle";
            this.cmdCopyDocTitle.Size = new System.Drawing.Size(120, 32);
            this.cmdCopyDocTitle.TabIndex = 1;
            this.cmdCopyDocTitle.Text = "&DOC TITLE :";
            this.cmdCopyDocTitle.UseVisualStyleBackColor = false;
            this.cmdCopyDocTitle.Click += new System.EventHandler(this.cmdCopyDocTitle_Click);
            // 
            // txtDocTitle
            // 
            this.txtDocTitle.BackColor = System.Drawing.SystemColors.Window;
            this.txtDocTitle.Location = new System.Drawing.Point(144, 90);
            this.txtDocTitle.Multiline = true;
            this.txtDocTitle.Name = "txtDocTitle";
            this.txtDocTitle.ReadOnly = true;
            this.txtDocTitle.Size = new System.Drawing.Size(309, 25);
            this.txtDocTitle.TabIndex = 6;
            // 
            // cmdCopyProductName
            // 
            this.cmdCopyProductName.BackColor = System.Drawing.SystemColors.Control;
            this.cmdCopyProductName.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdCopyProductName.Location = new System.Drawing.Point(12, 194);
            this.cmdCopyProductName.Name = "cmdCopyProductName";
            this.cmdCopyProductName.Size = new System.Drawing.Size(121, 35);
            this.cmdCopyProductName.TabIndex = 3;
            this.cmdCopyProductName.Text = "PRODUCT &NAME :";
            this.cmdCopyProductName.UseVisualStyleBackColor = false;
            this.cmdCopyProductName.Click += new System.EventHandler(this.cmdCopyProductName_Click);
            // 
            // txtDocFileName
            // 
            this.txtDocFileName.BackColor = System.Drawing.SystemColors.Window;
            this.txtDocFileName.Location = new System.Drawing.Point(143, 40);
            this.txtDocFileName.Multiline = true;
            this.txtDocFileName.Name = "txtDocFileName";
            this.txtDocFileName.ReadOnly = true;
            this.txtDocFileName.Size = new System.Drawing.Size(309, 25);
            this.txtDocFileName.TabIndex = 5;
            // 
            // cmdCopyProductID
            // 
            this.cmdCopyProductID.BackColor = System.Drawing.SystemColors.Control;
            this.cmdCopyProductID.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdCopyProductID.Location = new System.Drawing.Point(11, 141);
            this.cmdCopyProductID.Name = "cmdCopyProductID";
            this.cmdCopyProductID.Size = new System.Drawing.Size(121, 34);
            this.cmdCopyProductID.TabIndex = 2;
            this.cmdCopyProductID.Text = "&PRODUCT ID :";
            this.cmdCopyProductID.UseVisualStyleBackColor = false;
            this.cmdCopyProductID.Click += new System.EventHandler(this.cmdCopyProductID_Click);
            // 
            // lblWizardInfo
            // 
            this.lblWizardInfo.AutoSize = true;
            this.lblWizardInfo.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.lblWizardInfo.Location = new System.Drawing.Point(14, 297);
            this.lblWizardInfo.Name = "lblWizardInfo";
            this.lblWizardInfo.Size = new System.Drawing.Size(0, 13);
            this.lblWizardInfo.TabIndex = 17;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label0);
            this.panel1.Controls.Add(this.shp1);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.shp0);
            this.panel1.Controls.Add(this.mapStatusStart);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Location = new System.Drawing.Point(6, 20);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(104, 226);
            this.panel1.TabIndex = 30;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Black;
            this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(53, 164);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(33, 13);
            this.label2.TabIndex = 31;
            this.label2.Tag = "1";
            this.label2.Text = "Open";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Black;
            this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(58, 86);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 26);
            this.label1.TabIndex = 30;
            this.label1.Tag = "1";
            this.label1.Text = "File \r\nDetails";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label0
            // 
            this.label0.AutoSize = true;
            this.label0.BackColor = System.Drawing.Color.Black;
            this.label0.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label0.ForeColor = System.Drawing.Color.White;
            this.label0.Location = new System.Drawing.Point(61, 15);
            this.label0.Name = "label0";
            this.label0.Size = new System.Drawing.Size(36, 26);
            this.label0.TabIndex = 29;
            this.label0.Tag = "0";
            this.label0.Text = "Select\r\nFile";
            this.label0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // shp1
            // 
            this.shp1.BackColor = System.Drawing.Color.Silver;
            this.shp1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.shp1.Location = new System.Drawing.Point(40, 86);
            this.shp1.Name = "shp1";
            this.shp1.Size = new System.Drawing.Size(15, 15);
            this.shp1.TabIndex = 28;
            this.shp1.Tag = "1";
            this.shp1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.shp1.Click += new System.EventHandler(this.shp1_Click);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Red;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(3, 167);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(15, 15);
            this.label3.TabIndex = 27;
            this.label3.Tag = "Exclude";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // shp0
            // 
            this.shp0.BackColor = System.Drawing.Color.Lime;
            this.shp0.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.shp0.Location = new System.Drawing.Point(43, 11);
            this.shp0.Name = "shp0";
            this.shp0.Size = new System.Drawing.Size(15, 15);
            this.shp0.TabIndex = 25;
            this.shp0.Tag = "0";
            this.shp0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // mapStatusStart
            // 
            this.mapStatusStart.BackColor = System.Drawing.Color.Red;
            this.mapStatusStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mapStatusStart.Location = new System.Drawing.Point(3, 13);
            this.mapStatusStart.Name = "mapStatusStart";
            this.mapStatusStart.Size = new System.Drawing.Size(15, 15);
            this.mapStatusStart.TabIndex = 24;
            this.mapStatusStart.Tag = "Exclude";
            this.mapStatusStart.Text = "-";
            this.mapStatusStart.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label7.Location = new System.Drawing.Point(11, 22);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(37, 157);
            this.label7.TabIndex = 23;
            this.label7.Tag = "Exclude";
            // 
            // FormFileOpenWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(607, 353);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lblWizardInfo);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.cmdFinish);
            this.Controls.Add(this.cmdNext);
            this.Controls.Add(this.cmdBack);
            this.Controls.Add(this.cmdCancel);
            this.Name = "FormFileOpenWizard";
            this.Text = "File Open Wizard";
            this.Activated += new System.EventHandler(this.FormFileOpenWizard_Activate);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormFileOpenWizard_FormClosing);
            this.Load += new System.EventHandler(this.FormFileOpenWizard_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Button cmdBack;
        private System.Windows.Forms.Button cmdNext;
        private System.Windows.Forms.Button cmdFinish;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label lblFileOpen;
        private System.Windows.Forms.TextBox txtFindString;
        private System.Windows.Forms.Label lblFileName;
        private System.Windows.Forms.CheckBox chkFindProcessingSteps;
        private System.Windows.Forms.Label lblProject;
        private System.Windows.Forms.ComboBox cboSelectFile;
        private System.Windows.Forms.Label lblRecordType;
        private System.Windows.Forms.ComboBox cboProject;
        private System.Windows.Forms.Label lblOpUnit;
        private System.Windows.Forms.ComboBox cboDocType;
        private System.Windows.Forms.ComboBox cboOpUnit;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button cmdUpdateDocTitle;
        private System.Windows.Forms.TextBox txtDocProductName;
        private System.Windows.Forms.Button cmdCopyFileName;
        private System.Windows.Forms.Label lblClickTo;
        private System.Windows.Forms.TextBox txtDocProductID;
        private System.Windows.Forms.Button cmdCopyDocTitle;
        private System.Windows.Forms.TextBox txtDocTitle;
        private System.Windows.Forms.Button cmdCopyProductName;
        private System.Windows.Forms.TextBox txtDocFileName;
        private System.Windows.Forms.Button cmdCopyProductID;
        private System.Windows.Forms.Label lblWizardInfo;
        public System.Windows.Forms.Label lblUpdated;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label0;
        private System.Windows.Forms.Label shp1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label shp0;
        private System.Windows.Forms.Label mapStatusStart;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label2;
    }
}