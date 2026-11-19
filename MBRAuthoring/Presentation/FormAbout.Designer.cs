
namespace MBRAuthoring.Presentation
{
    partial class FormAbout
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormAbout));
            this.cmdDBStats = new System.Windows.Forms.Button();
            this.cmdClose = new System.Windows.Forms.Button();
            this.lblVersionCap = new System.Windows.Forms.Label();
            this.lblDateCap = new System.Windows.Forms.Label();
            this.lblEnvironmentCap = new System.Windows.Forms.Label();
            this.lblDocumentPathCap = new System.Windows.Forms.Label();
            this.lblTemplatePathCap = new System.Windows.Forms.Label();
            this.lblWordVersionCap = new System.Windows.Forms.Label();
            this.lblSysUserType = new System.Windows.Forms.Label();
            this.lblDBServerCap = new System.Windows.Forms.Label();
            this.lblDBVersionCap = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblDBNameCap = new System.Windows.Forms.Label();
            this.lblVersion = new System.Windows.Forms.Label();
            this.lblDate = new System.Windows.Forms.Label();
            this.lblEnvironment = new System.Windows.Forms.Label();
            this.lblDocumentPath = new System.Windows.Forms.Label();
            this.lblTemplatePath = new System.Windows.Forms.Label();
            this.lblWordVersion = new System.Windows.Forms.Label();
            this.lblUserType = new System.Windows.Forms.Label();
            this.lblDBServer = new System.Windows.Forms.Label();
            this.lblDBVersion = new System.Windows.Forms.Label();
            this.lblDBName = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cmdDBStats
            // 
            this.cmdDBStats.Location = new System.Drawing.Point(60, 330);
            this.cmdDBStats.Name = "cmdDBStats";
            this.cmdDBStats.Size = new System.Drawing.Size(154, 30);
            this.cmdDBStats.TabIndex = 0;
            this.cmdDBStats.Text = "Database Statistics";
            this.cmdDBStats.UseVisualStyleBackColor = true;
            this.cmdDBStats.Click += new System.EventHandler(this.cmdDBStats_Click);
            // 
            // cmdClose
            // 
            this.cmdClose.Location = new System.Drawing.Point(364, 330);
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.Size = new System.Drawing.Size(154, 30);
            this.cmdClose.TabIndex = 1;
            this.cmdClose.Text = "OK";
            this.cmdClose.UseVisualStyleBackColor = true;
            this.cmdClose.Click += new System.EventHandler(this.cmdClose_Click);
            // 
            // lblVersionCap
            // 
            this.lblVersionCap.AutoSize = true;
            this.lblVersionCap.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVersionCap.Location = new System.Drawing.Point(131, 65);
            this.lblVersionCap.Name = "lblVersionCap";
            this.lblVersionCap.Size = new System.Drawing.Size(55, 13);
            this.lblVersionCap.TabIndex = 2;
            this.lblVersionCap.Text = "Version :";
            // 
            // lblDateCap
            // 
            this.lblDateCap.AutoSize = true;
            this.lblDateCap.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDateCap.Location = new System.Drawing.Point(146, 88);
            this.lblDateCap.Name = "lblDateCap";
            this.lblDateCap.Size = new System.Drawing.Size(40, 13);
            this.lblDateCap.TabIndex = 3;
            this.lblDateCap.Text = "Date :";
            // 
            // lblEnvironmentCap
            // 
            this.lblEnvironmentCap.AutoSize = true;
            this.lblEnvironmentCap.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEnvironmentCap.Location = new System.Drawing.Point(102, 111);
            this.lblEnvironmentCap.Name = "lblEnvironmentCap";
            this.lblEnvironmentCap.Size = new System.Drawing.Size(85, 13);
            this.lblEnvironmentCap.TabIndex = 4;
            this.lblEnvironmentCap.Text = "Environment :";
            // 
            // lblDocumentPathCap
            // 
            this.lblDocumentPathCap.AutoSize = true;
            this.lblDocumentPathCap.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDocumentPathCap.Location = new System.Drawing.Point(88, 134);
            this.lblDocumentPathCap.Name = "lblDocumentPathCap";
            this.lblDocumentPathCap.Size = new System.Drawing.Size(100, 13);
            this.lblDocumentPathCap.TabIndex = 5;
            this.lblDocumentPathCap.Text = "Document Path :";
            // 
            // lblTemplatePathCap
            // 
            this.lblTemplatePathCap.AutoSize = true;
            this.lblTemplatePathCap.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTemplatePathCap.Location = new System.Drawing.Point(92, 157);
            this.lblTemplatePathCap.Name = "lblTemplatePathCap";
            this.lblTemplatePathCap.Size = new System.Drawing.Size(96, 13);
            this.lblTemplatePathCap.TabIndex = 6;
            this.lblTemplatePathCap.Text = "Template Path :";
            // 
            // lblWordVersionCap
            // 
            this.lblWordVersionCap.AutoSize = true;
            this.lblWordVersionCap.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWordVersionCap.Location = new System.Drawing.Point(101, 180);
            this.lblWordVersionCap.Name = "lblWordVersionCap";
            this.lblWordVersionCap.Size = new System.Drawing.Size(88, 13);
            this.lblWordVersionCap.TabIndex = 7;
            this.lblWordVersionCap.Text = "Word Version :";
            // 
            // lblSysUserType
            // 
            this.lblSysUserType.AutoSize = true;
            this.lblSysUserType.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSysUserType.Location = new System.Drawing.Point(119, 203);
            this.lblSysUserType.Name = "lblSysUserType";
            this.lblSysUserType.Size = new System.Drawing.Size(70, 13);
            this.lblSysUserType.TabIndex = 8;
            this.lblSysUserType.Text = "User Type :";
            // 
            // lblDBServerCap
            // 
            this.lblDBServerCap.AutoSize = true;
            this.lblDBServerCap.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDBServerCap.Location = new System.Drawing.Point(121, 227);
            this.lblDBServerCap.Name = "lblDBServerCap";
            this.lblDBServerCap.Size = new System.Drawing.Size(69, 13);
            this.lblDBServerCap.TabIndex = 9;
            this.lblDBServerCap.Text = "DB Server :";
            // 
            // lblDBVersionCap
            // 
            this.lblDBVersionCap.AutoSize = true;
            this.lblDBVersionCap.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDBVersionCap.Location = new System.Drawing.Point(118, 252);
            this.lblDBVersionCap.Name = "lblDBVersionCap";
            this.lblDBVersionCap.Size = new System.Drawing.Size(73, 13);
            this.lblDBVersionCap.TabIndex = 10;
            this.lblDBVersionCap.Text = "DB Version :";
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(40, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(132, 19);
            this.lblTitle.TabIndex = 11;
            this.lblTitle.Text = "MBR Authoring";
            // 
            // lblDBNameCap
            // 
            this.lblDBNameCap.AutoSize = true;
            this.lblDBNameCap.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDBNameCap.Location = new System.Drawing.Point(128, 278);
            this.lblDBNameCap.Name = "lblDBNameCap";
            this.lblDBNameCap.Size = new System.Drawing.Size(63, 13);
            this.lblDBNameCap.TabIndex = 12;
            this.lblDBNameCap.Text = "DB Name :";
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Location = new System.Drawing.Point(193, 64);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(52, 13);
            this.lblVersion.TabIndex = 13;
            this.lblVersion.Text = "lblVersion";
            // 
            // lblDate
            // 
            this.lblDate.AutoSize = true;
            this.lblDate.Location = new System.Drawing.Point(193, 87);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(40, 13);
            this.lblDate.TabIndex = 14;
            this.lblDate.Text = "lblDate";
            // 
            // lblEnvironment
            // 
            this.lblEnvironment.AutoSize = true;
            this.lblEnvironment.Location = new System.Drawing.Point(193, 110);
            this.lblEnvironment.Name = "lblEnvironment";
            this.lblEnvironment.Size = new System.Drawing.Size(76, 13);
            this.lblEnvironment.TabIndex = 15;
            this.lblEnvironment.Text = "lblEnvironment";
            // 
            // lblDocumentPath
            // 
            this.lblDocumentPath.AutoSize = true;
            this.lblDocumentPath.Location = new System.Drawing.Point(194, 134);
            this.lblDocumentPath.Name = "lblDocumentPath";
            this.lblDocumentPath.Size = new System.Drawing.Size(88, 13);
            this.lblDocumentPath.TabIndex = 16;
            this.lblDocumentPath.Text = "lblDocumentPath";
            // 
            // lblTemplatePath
            // 
            this.lblTemplatePath.AutoSize = true;
            this.lblTemplatePath.Location = new System.Drawing.Point(194, 157);
            this.lblTemplatePath.Name = "lblTemplatePath";
            this.lblTemplatePath.Size = new System.Drawing.Size(83, 13);
            this.lblTemplatePath.TabIndex = 17;
            this.lblTemplatePath.Text = "lblTemplatePath";
            // 
            // lblWordVersion
            // 
            this.lblWordVersion.AutoSize = true;
            this.lblWordVersion.Location = new System.Drawing.Point(194, 180);
            this.lblWordVersion.Name = "lblWordVersion";
            this.lblWordVersion.Size = new System.Drawing.Size(78, 13);
            this.lblWordVersion.TabIndex = 18;
            this.lblWordVersion.Text = "lblWordVersion";
            // 
            // lblUserType
            // 
            this.lblUserType.AutoSize = true;
            this.lblUserType.Location = new System.Drawing.Point(194, 205);
            this.lblUserType.Name = "lblUserType";
            this.lblUserType.Size = new System.Drawing.Size(63, 13);
            this.lblUserType.TabIndex = 19;
            this.lblUserType.Text = "lblUserType";
            // 
            // lblDBServer
            // 
            this.lblDBServer.AutoSize = true;
            this.lblDBServer.Location = new System.Drawing.Point(193, 229);
            this.lblDBServer.Name = "lblDBServer";
            this.lblDBServer.Size = new System.Drawing.Size(63, 13);
            this.lblDBServer.TabIndex = 20;
            this.lblDBServer.Text = "lblDBServer";
            // 
            // lblDBVersion
            // 
            this.lblDBVersion.AutoSize = true;
            this.lblDBVersion.Location = new System.Drawing.Point(193, 254);
            this.lblDBVersion.Name = "lblDBVersion";
            this.lblDBVersion.Size = new System.Drawing.Size(67, 13);
            this.lblDBVersion.TabIndex = 21;
            this.lblDBVersion.Text = "lblDBVersion";
            // 
            // lblDBName
            // 
            this.lblDBName.AutoSize = true;
            this.lblDBName.Location = new System.Drawing.Point(193, 278);
            this.lblDBName.Name = "lblDBName";
            this.lblDBName.Size = new System.Drawing.Size(60, 13);
            this.lblDBName.TabIndex = 22;
            this.lblDBName.Text = "lblDBName";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(1, 301);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(745, 13);
            this.label1.TabIndex = 23;
            this.label1.Text = resources.GetString("label1.Text");
            // 
            // FormAbout
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.ClientSize = new System.Drawing.Size(589, 373);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblDBName);
            this.Controls.Add(this.lblDBVersion);
            this.Controls.Add(this.lblDBServer);
            this.Controls.Add(this.lblUserType);
            this.Controls.Add(this.lblWordVersion);
            this.Controls.Add(this.lblTemplatePath);
            this.Controls.Add(this.lblDocumentPath);
            this.Controls.Add(this.lblEnvironment);
            this.Controls.Add(this.lblDate);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.lblDBNameCap);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.lblDBVersionCap);
            this.Controls.Add(this.lblDBServerCap);
            this.Controls.Add(this.lblSysUserType);
            this.Controls.Add(this.lblWordVersionCap);
            this.Controls.Add(this.lblTemplatePathCap);
            this.Controls.Add(this.lblDocumentPathCap);
            this.Controls.Add(this.lblEnvironmentCap);
            this.Controls.Add(this.lblDateCap);
            this.Controls.Add(this.lblVersionCap);
            this.Controls.Add(this.cmdClose);
            this.Controls.Add(this.cmdDBStats);
            this.Name = "FormAbout";
            this.Text = "About MBR";
            this.Load += new System.EventHandler(this.FormAbout_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button cmdDBStats;
        private System.Windows.Forms.Button cmdClose;
        private System.Windows.Forms.Label lblVersionCap;
        private System.Windows.Forms.Label lblDateCap;
        private System.Windows.Forms.Label lblEnvironmentCap;
        private System.Windows.Forms.Label lblDocumentPathCap;
        private System.Windows.Forms.Label lblTemplatePathCap;
        private System.Windows.Forms.Label lblWordVersionCap;
        private System.Windows.Forms.Label lblSysUserType;
        private System.Windows.Forms.Label lblDBServerCap;
        private System.Windows.Forms.Label lblDBVersionCap;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblDBNameCap;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.Label lblDate;
        private System.Windows.Forms.Label lblEnvironment;
        private System.Windows.Forms.Label lblDocumentPath;
        private System.Windows.Forms.Label lblTemplatePath;
        private System.Windows.Forms.Label lblWordVersion;
        private System.Windows.Forms.Label lblUserType;
        private System.Windows.Forms.Label lblDBServer;
        private System.Windows.Forms.Label lblDBVersion;
        private System.Windows.Forms.Label lblDBName;
        private System.Windows.Forms.Label label1;
    }
}