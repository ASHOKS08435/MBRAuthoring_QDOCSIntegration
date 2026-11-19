
namespace MBRAuthoring.Presentation
{
    partial class FormSections
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
            this.grdSections = new System.Windows.Forms.DataGridView();
            this.chkHeaderLinkPrevious = new System.Windows.Forms.CheckBox();
            this.chkFooterLinkPrevious = new System.Windows.Forms.CheckBox();
            this.chkRestartPageNumbering = new System.Windows.Forms.CheckBox();
            this.txtRestartPageNumbering = new System.Windows.Forms.TextBox();
            this.cmdApply = new System.Windows.Forms.Button();
            this.cmdClose = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.grdSections)).BeginInit();
            this.SuspendLayout();
            // 
            // grdSections
            // 
            this.grdSections.AllowUserToAddRows = false;
            this.grdSections.AllowUserToDeleteRows = false;
            this.grdSections.AllowUserToResizeColumns = false;
            this.grdSections.AllowUserToResizeRows = false;
            this.grdSections.BackgroundColor = System.Drawing.SystemColors.Window;
            this.grdSections.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.grdSections.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.grdSections.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.grdSections.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdSections.Location = new System.Drawing.Point(12, 12);
            this.grdSections.Name = "grdSections";
            this.grdSections.ReadOnly = true;
            this.grdSections.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.grdSections.RowHeadersVisible = false;
            this.grdSections.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.grdSections.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grdSections.ShowEditingIcon = false;
            this.grdSections.Size = new System.Drawing.Size(583, 165);
            this.grdSections.TabIndex = 0;
            this.grdSections.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.grdSections_CellClick);
            // 
            // chkHeaderLinkPrevious
            // 
            this.chkHeaderLinkPrevious.AutoSize = true;
            this.chkHeaderLinkPrevious.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.chkHeaderLinkPrevious.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkHeaderLinkPrevious.Location = new System.Drawing.Point(12, 195);
            this.chkHeaderLinkPrevious.Name = "chkHeaderLinkPrevious";
            this.chkHeaderLinkPrevious.Size = new System.Drawing.Size(145, 17);
            this.chkHeaderLinkPrevious.TabIndex = 1;
            this.chkHeaderLinkPrevious.Text = "Header: Link To Previous";
            this.chkHeaderLinkPrevious.UseVisualStyleBackColor = false;
            this.chkHeaderLinkPrevious.CheckedChanged += new System.EventHandler(this.chkHeaderLinkPrevious_CheckedChanged);
            // 
            // chkFooterLinkPrevious
            // 
            this.chkFooterLinkPrevious.AutoSize = true;
            this.chkFooterLinkPrevious.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.chkFooterLinkPrevious.Location = new System.Drawing.Point(12, 225);
            this.chkFooterLinkPrevious.Name = "chkFooterLinkPrevious";
            this.chkFooterLinkPrevious.Size = new System.Drawing.Size(142, 17);
            this.chkFooterLinkPrevious.TabIndex = 2;
            this.chkFooterLinkPrevious.Text = "Footer: Link To Previous";
            this.chkFooterLinkPrevious.UseVisualStyleBackColor = false;
            this.chkFooterLinkPrevious.CheckedChanged += new System.EventHandler(this.chkFooterLinkPrevious_CheckedChanged);
            // 
            // chkRestartPageNumbering
            // 
            this.chkRestartPageNumbering.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.chkRestartPageNumbering.Location = new System.Drawing.Point(12, 255);
            this.chkRestartPageNumbering.Name = "chkRestartPageNumbering";
            this.chkRestartPageNumbering.Size = new System.Drawing.Size(350, 17);
            this.chkRestartPageNumbering.TabIndex = 3;
            this.chkRestartPageNumbering.Text = "Footer: Restart Page Numbering, at Page:";
            this.chkRestartPageNumbering.UseVisualStyleBackColor = false;
            this.chkRestartPageNumbering.CheckedChanged += new System.EventHandler(this.chkRestartPageNumbering_CheckedChanged);
            // 
            // txtRestartPageNumbering
            // 
            this.txtRestartPageNumbering.Location = new System.Drawing.Point(247, 253);
            this.txtRestartPageNumbering.Name = "txtRestartPageNumbering";
            this.txtRestartPageNumbering.Size = new System.Drawing.Size(81, 21);
            this.txtRestartPageNumbering.TabIndex = 4;
            // 
            // cmdApply
            // 
            this.cmdApply.BackColor = System.Drawing.SystemColors.Control;
            this.cmdApply.Location = new System.Drawing.Point(12, 305);
            this.cmdApply.Name = "cmdApply";
            this.cmdApply.Size = new System.Drawing.Size(230, 32);
            this.cmdApply.TabIndex = 5;
            this.cmdApply.Text = "Apply and Close";
            this.cmdApply.UseVisualStyleBackColor = false;
            this.cmdApply.Click += new System.EventHandler(this.cmdApply_Click);
            // 
            // cmdClose
            // 
            this.cmdClose.BackColor = System.Drawing.SystemColors.Control;
            this.cmdClose.Location = new System.Drawing.Point(498, 305);
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.Size = new System.Drawing.Size(97, 32);
            this.cmdClose.TabIndex = 6;
            this.cmdClose.Text = "Close";
            this.cmdClose.UseVisualStyleBackColor = false;
            this.cmdClose.Click += new System.EventHandler(this.cmdClose_Click);
            // 
            // FormSections
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(607, 353);
            this.Controls.Add(this.cmdClose);
            this.Controls.Add(this.cmdApply);
            this.Controls.Add(this.txtRestartPageNumbering);
            this.Controls.Add(this.chkRestartPageNumbering);
            this.Controls.Add(this.chkFooterLinkPrevious);
            this.Controls.Add(this.chkHeaderLinkPrevious);
            this.Controls.Add(this.grdSections);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "FormSections";
            this.Text = "MBR Section Management";
            this.Activated += new System.EventHandler(this.FormSections_Activate);
            this.Load += new System.EventHandler(this.FormSections_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grdSections)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView grdSections;
        private System.Windows.Forms.CheckBox chkHeaderLinkPrevious;
        private System.Windows.Forms.CheckBox chkFooterLinkPrevious;
        private System.Windows.Forms.CheckBox chkRestartPageNumbering;
        private System.Windows.Forms.TextBox txtRestartPageNumbering;
        private System.Windows.Forms.Button cmdApply;
        private System.Windows.Forms.Button cmdClose;
    }
}