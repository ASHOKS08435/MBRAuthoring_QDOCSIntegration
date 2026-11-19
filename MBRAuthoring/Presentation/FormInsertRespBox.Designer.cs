
namespace MBRAuthoring.Presentation
{
    partial class FormInsertRespBox
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
            this.tabResponseBox = new System.Windows.Forms.TabControl();
            this.tbResponseBox = new System.Windows.Forms.TabPage();
            this.grpResponseBox = new System.Windows.Forms.GroupBox();
            this.rdoEntryBox = new System.Windows.Forms.RadioButton();
            this.rdoDoneByCheckedBy = new System.Windows.Forms.RadioButton();
            this.rdoDonebyDate = new System.Windows.Forms.RadioButton();
            this.rdoYesNo = new System.Windows.Forms.RadioButton();
            this.btnRspBoxInsert = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnFinish = new System.Windows.Forms.Button();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label0 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.shp0 = new System.Windows.Forms.Label();
            this.mapStatusStart = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnBack = new System.Windows.Forms.Button();
            this.tabResponseBox.SuspendLayout();
            this.tbResponseBox.SuspendLayout();
            this.grpResponseBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabResponseBox
            // 
            this.tabResponseBox.Controls.Add(this.tbResponseBox);
            this.tabResponseBox.Location = new System.Drawing.Point(154, -1);
            this.tabResponseBox.Name = "tabResponseBox";
            this.tabResponseBox.SelectedIndex = 0;
            this.tabResponseBox.Size = new System.Drawing.Size(480, 307);
            this.tabResponseBox.TabIndex = 0;
            // 
            // tbResponseBox
            // 
            this.tbResponseBox.Controls.Add(this.grpResponseBox);
            this.tbResponseBox.Controls.Add(this.btnRspBoxInsert);
            this.tbResponseBox.Controls.Add(this.lblTitle);
            this.tbResponseBox.Location = new System.Drawing.Point(4, 22);
            this.tbResponseBox.Name = "tbResponseBox";
            this.tbResponseBox.Padding = new System.Windows.Forms.Padding(3);
            this.tbResponseBox.Size = new System.Drawing.Size(472, 281);
            this.tbResponseBox.TabIndex = 0;
            this.tbResponseBox.Text = "ResponseBox";
            this.tbResponseBox.UseVisualStyleBackColor = true;
            // 
            // grpResponseBox
            // 
            this.grpResponseBox.Controls.Add(this.rdoEntryBox);
            this.grpResponseBox.Controls.Add(this.rdoDoneByCheckedBy);
            this.grpResponseBox.Controls.Add(this.rdoDonebyDate);
            this.grpResponseBox.Controls.Add(this.rdoYesNo);
            this.grpResponseBox.Location = new System.Drawing.Point(45, 36);
            this.grpResponseBox.Name = "grpResponseBox";
            this.grpResponseBox.Size = new System.Drawing.Size(387, 188);
            this.grpResponseBox.TabIndex = 3;
            this.grpResponseBox.TabStop = false;
            this.grpResponseBox.Text = "Select choice of Response Box to be entered...";
            // 
            // rdoEntryBox
            // 
            this.rdoEntryBox.AutoSize = true;
            this.rdoEntryBox.Location = new System.Drawing.Point(20, 124);
            this.rdoEntryBox.Name = "rdoEntryBox";
            this.rdoEntryBox.Size = new System.Drawing.Size(70, 17);
            this.rdoEntryBox.TabIndex = 14;
            this.rdoEntryBox.TabStop = true;
            this.rdoEntryBox.Text = "Entry Box";
            this.rdoEntryBox.UseVisualStyleBackColor = true;
            this.rdoEntryBox.CheckedChanged += new System.EventHandler(this.rdoEntryBox_CheckedChanged);
            // 
            // rdoDoneByCheckedBy
            // 
            this.rdoDoneByCheckedBy.AutoSize = true;
            this.rdoDoneByCheckedBy.Location = new System.Drawing.Point(19, 95);
            this.rdoDoneByCheckedBy.Name = "rdoDoneByCheckedBy";
            this.rdoDoneByCheckedBy.Size = new System.Drawing.Size(211, 17);
            this.rdoDoneByCheckedBy.TabIndex = 13;
            this.rdoDoneByCheckedBy.TabStop = true;
            this.rdoDoneByCheckedBy.Text = "Done By, Checked By, and Date boxes";
            this.rdoDoneByCheckedBy.UseVisualStyleBackColor = true;
            this.rdoDoneByCheckedBy.CheckedChanged += new System.EventHandler(this.rdoDoneByCheckedBy_CheckedChanged);
            // 
            // rdoDonebyDate
            // 
            this.rdoDonebyDate.AutoSize = true;
            this.rdoDonebyDate.Location = new System.Drawing.Point(20, 66);
            this.rdoDonebyDate.Name = "rdoDonebyDate";
            this.rdoDonebyDate.Size = new System.Drawing.Size(144, 17);
            this.rdoDonebyDate.TabIndex = 12;
            this.rdoDonebyDate.TabStop = true;
            this.rdoDonebyDate.Text = "Done By and Date boxes";
            this.rdoDonebyDate.UseVisualStyleBackColor = true;
            this.rdoDonebyDate.CheckedChanged += new System.EventHandler(this.rdoDonebyDate_CheckedChanged);
            // 
            // rdoYesNo
            // 
            this.rdoYesNo.AutoSize = true;
            this.rdoYesNo.Location = new System.Drawing.Point(19, 37);
            this.rdoYesNo.Name = "rdoYesNo";
            this.rdoYesNo.Size = new System.Drawing.Size(133, 17);
            this.rdoYesNo.TabIndex = 11;
            this.rdoYesNo.TabStop = true;
            this.rdoYesNo.Text = "Yes / No Check boxes";
            this.rdoYesNo.UseVisualStyleBackColor = true;
            this.rdoYesNo.CheckedChanged += new System.EventHandler(this.rdoYesNo_CheckedChanged);
            // 
            // btnRspBoxInsert
            // 
            this.btnRspBoxInsert.Location = new System.Drawing.Point(45, 230);
            this.btnRspBoxInsert.Name = "btnRspBoxInsert";
            this.btnRspBoxInsert.Size = new System.Drawing.Size(387, 39);
            this.btnRspBoxInsert.TabIndex = 11;
            this.btnRspBoxInsert.Text = "Insert Selected Immediately";
            this.btnRspBoxInsert.UseVisualStyleBackColor = true;
            this.btnRspBoxInsert.Click += new System.EventHandler(this.btnRspBoxInsert_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(13, 9);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(193, 16);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "(Title set in program code...)";
            // 
            // btnClose
            // 
            this.btnClose.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.Location = new System.Drawing.Point(282, 329);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(88, 31);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnFinish
            // 
            this.btnFinish.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFinish.Location = new System.Drawing.Point(552, 329);
            this.btnFinish.Name = "btnFinish";
            this.btnFinish.Size = new System.Drawing.Size(81, 31);
            this.btnFinish.TabIndex = 2;
            this.btnFinish.Text = "Finish";
            this.btnFinish.UseVisualStyleBackColor = true;
            this.btnFinish.Click += new System.EventHandler(this.btnFinish_Click);
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label0);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.shp0);
            this.panel1.Controls.Add(this.mapStatusStart);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(136, 226);
            this.panel1.TabIndex = 33;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Black;
            this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(53, 164);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 13);
            this.label2.TabIndex = 31;
            this.label2.Tag = "1";
            this.label2.Text = "Finish";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label0
            // 
            this.label0.AutoSize = true;
            this.label0.BackColor = System.Drawing.Color.Black;
            this.label0.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label0.ForeColor = System.Drawing.Color.White;
            this.label0.Location = new System.Drawing.Point(59, 11);
            this.label0.Name = "label0";
            this.label0.Size = new System.Drawing.Size(77, 39);
            this.label0.TabIndex = 29;
            this.label0.Tag = "0";
            this.label0.Text = "Insert \r\n    Response \r\nBox(es)";
            this.label0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label4.Location = new System.Drawing.Point(11, 22);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(37, 157);
            this.label4.TabIndex = 23;
            this.label4.Tag = "Exclude";
            // 
            // btnNext
            // 
            this.btnNext.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNext.Location = new System.Drawing.Point(458, 329);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(88, 31);
            this.btnNext.TabIndex = 34;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            // 
            // btnBack
            // 
            this.btnBack.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBack.Location = new System.Drawing.Point(377, 329);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(88, 31);
            this.btnBack.TabIndex = 35;
            this.btnBack.Text = "Back";
            this.btnBack.UseVisualStyleBackColor = true;
            // 
            // FormInsertRespBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(668, 372);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnFinish);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.tabResponseBox);
            this.Name = "FormInsertRespBox";
            this.Text = "Select & Insert Response Box Wizard";
            this.Activated += new System.EventHandler(this.FormInsertRespBox_Activate);
            this.Load += new System.EventHandler(this.FormInsertRespBox_Load);
            this.tabResponseBox.ResumeLayout(false);
            this.tbResponseBox.ResumeLayout(false);
            this.tbResponseBox.PerformLayout();
            this.grpResponseBox.ResumeLayout(false);
            this.grpResponseBox.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabResponseBox;
        private System.Windows.Forms.TabPage tbResponseBox;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button btnRspBoxInsert;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnFinish;
        private System.Windows.Forms.GroupBox grpResponseBox;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label0;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label shp0;
        private System.Windows.Forms.Label mapStatusStart;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RadioButton rdoDoneByCheckedBy;
        private System.Windows.Forms.RadioButton rdoDonebyDate;
        private System.Windows.Forms.RadioButton rdoYesNo;
        private System.Windows.Forms.RadioButton rdoEntryBox;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Button btnNext;
    }
}