
namespace MBRAuthoring.Presentation
{
    partial class FormVerifyWizard
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
            this.cmdPrint = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lstStatus = new System.Windows.Forms.DataGridView();
            this.chkMoveToLockedState = new System.Windows.Forms.CheckBox();
            this.cmdRunDocCompare = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lblRelatedDocs = new System.Windows.Forms.Label();
            this.txtStatus = new System.Windows.Forms.TextBox();
            this.lblThisDocVerify = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.lable1 = new System.Windows.Forms.Label();
            this.lable0 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.shp0 = new System.Windows.Forms.Label();
            this.mapStatusStart = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.lstStatus)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmdCancel
            // 
            this.cmdCancel.BackColor = System.Drawing.SystemColors.Control;
            this.cmdCancel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdCancel.Location = new System.Drawing.Point(333, 397);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(75, 23);
            this.cmdCancel.TabIndex = 5;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = false;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // cmdBack
            // 
            this.cmdBack.BackColor = System.Drawing.SystemColors.Control;
            this.cmdBack.Enabled = false;
            this.cmdBack.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdBack.Location = new System.Drawing.Point(431, 397);
            this.cmdBack.Name = "cmdBack";
            this.cmdBack.Size = new System.Drawing.Size(75, 23);
            this.cmdBack.TabIndex = 4;
            this.cmdBack.Text = "< &Back";
            this.cmdBack.UseVisualStyleBackColor = false;
            // 
            // cmdNext
            // 
            this.cmdNext.BackColor = System.Drawing.SystemColors.Control;
            this.cmdNext.Enabled = false;
            this.cmdNext.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdNext.Location = new System.Drawing.Point(512, 397);
            this.cmdNext.Name = "cmdNext";
            this.cmdNext.Size = new System.Drawing.Size(75, 23);
            this.cmdNext.TabIndex = 3;
            this.cmdNext.Text = "&Next >";
            this.cmdNext.UseVisualStyleBackColor = false;
            this.cmdNext.UseWaitCursor = true;
            // 
            // cmdFinish
            // 
            this.cmdFinish.BackColor = System.Drawing.SystemColors.Control;
            this.cmdFinish.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdFinish.Location = new System.Drawing.Point(620, 397);
            this.cmdFinish.Name = "cmdFinish";
            this.cmdFinish.Size = new System.Drawing.Size(75, 23);
            this.cmdFinish.TabIndex = 2;
            this.cmdFinish.Text = "&Close";
            this.cmdFinish.UseVisualStyleBackColor = false;
            this.cmdFinish.Click += new System.EventHandler(this.cmdFinish_Click);
            // 
            // cmdPrint
            // 
            this.cmdPrint.BackColor = System.Drawing.SystemColors.Control;
            this.cmdPrint.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdPrint.Location = new System.Drawing.Point(170, 397);
            this.cmdPrint.Name = "cmdPrint";
            this.cmdPrint.Size = new System.Drawing.Size(75, 23);
            this.cmdPrint.TabIndex = 1;
            this.cmdPrint.Text = "&Print";
            this.cmdPrint.UseVisualStyleBackColor = false;
            this.cmdPrint.Click += new System.EventHandler(this.cmdPrint_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.panel1.Controls.Add(this.lstStatus);
            this.panel1.Controls.Add(this.chkMoveToLockedState);
            this.panel1.Controls.Add(this.cmdRunDocCompare);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.lblRelatedDocs);
            this.panel1.Controls.Add(this.txtStatus);
            this.panel1.Controls.Add(this.lblThisDocVerify);
            this.panel1.Controls.Add(this.lblTitle);
            this.panel1.Location = new System.Drawing.Point(138, -2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(573, 393);
            this.panel1.TabIndex = 0;
            // 
            // lstStatus
            // 
            this.lstStatus.AllowUserToAddRows = false;
            this.lstStatus.AllowUserToResizeColumns = false;
            this.lstStatus.AllowUserToResizeRows = false;
            this.lstStatus.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.lstStatus.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.lstStatus.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.lstStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.lstStatus.ColumnHeadersVisible = false;
            this.lstStatus.GridColor = System.Drawing.SystemColors.Window;
            this.lstStatus.Location = new System.Drawing.Point(16, 176);
            this.lstStatus.Name = "lstStatus";
            this.lstStatus.ReadOnly = true;
            this.lstStatus.RowHeadersVisible = false;
            this.lstStatus.RowTemplate.Height = 15;
            this.lstStatus.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.lstStatus.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.lstStatus.Size = new System.Drawing.Size(541, 154);
            this.lstStatus.TabIndex = 7;
            // 
            // chkMoveToLockedState
            // 
            this.chkMoveToLockedState.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.chkMoveToLockedState.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkMoveToLockedState.Location = new System.Drawing.Point(16, 367);
            this.chkMoveToLockedState.Name = "chkMoveToLockedState";
            this.chkMoveToLockedState.Size = new System.Drawing.Size(581, 18);
            this.chkMoveToLockedState.TabIndex = 1;
            this.chkMoveToLockedState.Text = "Move To Locked State";
            this.chkMoveToLockedState.UseVisualStyleBackColor = false;
            // 
            // cmdRunDocCompare
            // 
            this.cmdRunDocCompare.BackColor = System.Drawing.SystemColors.Control;
            this.cmdRunDocCompare.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdRunDocCompare.Location = new System.Drawing.Point(16, 336);
            this.cmdRunDocCompare.Name = "cmdRunDocCompare";
            this.cmdRunDocCompare.Size = new System.Drawing.Size(541, 25);
            this.cmdRunDocCompare.TabIndex = 6;
            this.cmdRunDocCompare.Text = "Run \"DocCompa&re\" on this Document";
            this.cmdRunDocCompare.UseVisualStyleBackColor = false;
            this.cmdRunDocCompare.Click += new System.EventHandler(this.cmdRunDocCompare_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 160);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(465, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "  Master Document Name                                 Loc.        Module Doc Nam" +
    "e        Loc.        Pass/Fail";
            // 
            // lblRelatedDocs
            // 
            this.lblRelatedDocs.AutoSize = true;
            this.lblRelatedDocs.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.lblRelatedDocs.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRelatedDocs.Location = new System.Drawing.Point(13, 139);
            this.lblRelatedDocs.Name = "lblRelatedDocs";
            this.lblRelatedDocs.Size = new System.Drawing.Size(121, 13);
            this.lblRelatedDocs.TabIndex = 3;
            this.lblRelatedDocs.Text = "Related Documents:";
            // 
            // txtStatus
            // 
            this.txtStatus.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtStatus.Location = new System.Drawing.Point(16, 45);
            this.txtStatus.Multiline = true;
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtStatus.Size = new System.Drawing.Size(541, 90);
            this.txtStatus.TabIndex = 2;
            // 
            // lblThisDocVerify
            // 
            this.lblThisDocVerify.AutoSize = true;
            this.lblThisDocVerify.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.lblThisDocVerify.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblThisDocVerify.Location = new System.Drawing.Point(13, 27);
            this.lblThisDocVerify.Name = "lblThisDocVerify";
            this.lblThisDocVerify.Size = new System.Drawing.Size(94, 13);
            this.lblThisDocVerify.TabIndex = 1;
            this.lblThisDocVerify.Text = "This Document:";
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.lblTitle.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(12, 4);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(193, 16);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "(Title set in program code...)";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.lable1);
            this.panel2.Controls.Add(this.lable0);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.shp0);
            this.panel2.Controls.Add(this.mapStatusStart);
            this.panel2.Controls.Add(this.label7);
            this.panel2.Location = new System.Drawing.Point(4, 12);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(133, 212);
            this.panel2.TabIndex = 31;
            // 
            // lable1
            // 
            this.lable1.BackColor = System.Drawing.Color.Black;
            this.lable1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lable1.ForeColor = System.Drawing.Color.White;
            this.lable1.Location = new System.Drawing.Point(61, 166);
            this.lable1.Name = "lable1";
            this.lable1.Size = new System.Drawing.Size(37, 26);
            this.lable1.TabIndex = 29;
            this.lable1.Tag = "0";
            this.lable1.Text = "Finish";
            this.lable1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lable1.UseMnemonic = false;
            this.lable1.Click += new System.EventHandler(this.lable1_Click);
            // 
            // lable0
            // 
            this.lable0.BackColor = System.Drawing.Color.Black;
            this.lable0.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lable0.ForeColor = System.Drawing.Color.White;
            this.lable0.Location = new System.Drawing.Point(61, 8);
            this.lable0.Name = "lable0";
            this.lable0.Size = new System.Drawing.Size(72, 37);
            this.lable0.TabIndex = 28;
            this.lable0.Tag = "0";
            this.lable0.Text = "Disposition";
            this.lable0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lable0.UseMnemonic = false;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Red;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(1, 177);
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
            this.shp0.Location = new System.Drawing.Point(40, 19);
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
            this.mapStatusStart.Location = new System.Drawing.Point(1, 19);
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
            this.label7.Location = new System.Drawing.Point(8, 27);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(41, 161);
            this.label7.TabIndex = 23;
            this.label7.Tag = "Exclude";
            // 
            // FormVerifyWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(705, 427);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.cmdCancel);
            this.Controls.Add(this.cmdBack);
            this.Controls.Add(this.cmdNext);
            this.Controls.Add(this.cmdFinish);
            this.Controls.Add(this.cmdPrint);
            this.Controls.Add(this.panel1);
            this.Name = "FormVerifyWizard";
            this.Text = "Verify Document Wizard";
            this.Load += new System.EventHandler(this.FormVerifyWizard_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.lstStatus)).EndInit();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblThisDocVerify;
        private System.Windows.Forms.TextBox txtStatus;
        private System.Windows.Forms.Label lblRelatedDocs;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button cmdRunDocCompare;
        private System.Windows.Forms.CheckBox chkMoveToLockedState;
        private System.Windows.Forms.Button cmdPrint;
        private System.Windows.Forms.Button cmdFinish;
        private System.Windows.Forms.Button cmdNext;
        private System.Windows.Forms.Button cmdBack;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.DataGridView lstStatus;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label lable1;
        private System.Windows.Forms.Label lable0;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label shp0;
        private System.Windows.Forms.Label mapStatusStart;
        private System.Windows.Forms.Label label7;
    }
}