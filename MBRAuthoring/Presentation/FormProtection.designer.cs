
namespace MBRAuthoring.Presentation
{
    partial class FormProtection
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormProtection));
            this.lblInstructions = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.chkProtectDocument = new System.Windows.Forms.CheckBox();
            this.chkTrackChanges = new System.Windows.Forms.CheckBox();
            this.chkTrackHighlightScreen = new System.Windows.Forms.CheckBox();
            this.chkTrackHighlightPrint = new System.Windows.Forms.CheckBox();
            this.lblChangesDisplay = new System.Windows.Forms.Label();
            this.cboChangesDisplay = new System.Windows.Forms.ComboBox();
            this.cmdPasswordDisplay = new System.Windows.Forms.Button();
            this.cmdClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblInstructions
            // 
            this.lblInstructions.AutoSize = true;
            this.lblInstructions.Location = new System.Drawing.Point(12, 13);
            this.lblInstructions.Name = "lblInstructions";
            this.lblInstructions.Size = new System.Drawing.Size(535, 52);
            this.lblInstructions.TabIndex = 0;
            this.lblInstructions.Text = resources.GetString("lblInstructions.Text");
            // 
            // lblStatus
            // 
            this.lblStatus.BackColor = System.Drawing.Color.White;
            this.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblStatus.Location = new System.Drawing.Point(12, 85);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(583, 100);
            this.lblStatus.TabIndex = 1;
            this.lblStatus.Text = "<program will display document status: protection on or off, track changes on or " +
    "off.";
            // 
            // chkProtectDocument
            // 
            this.chkProtectDocument.Location = new System.Drawing.Point(12, 205);
            this.chkProtectDocument.Name = "chkProtectDocument";
            this.chkProtectDocument.Size = new System.Drawing.Size(120, 18);
            this.chkProtectDocument.TabIndex = 4;
            this.chkProtectDocument.Text = "Pro&tect Document";
            this.chkProtectDocument.UseVisualStyleBackColor = true;
            this.chkProtectDocument.CheckedChanged += new System.EventHandler(this.chkProtectDocument_CheckedChanged);
            // 
            // chkTrackChanges
            // 
            this.chkTrackChanges.Location = new System.Drawing.Point(12, 232);
            this.chkTrackChanges.Name = "chkTrackChanges";
            this.chkTrackChanges.Size = new System.Drawing.Size(170, 18);
            this.chkTrackChanges.TabIndex = 5;
            this.chkTrackChanges.Text = "&Track changes while editing";
            this.chkTrackChanges.UseVisualStyleBackColor = true;
            this.chkTrackChanges.CheckedChanged += new System.EventHandler(this.chkTrackChanges_CheckedChanged);
            // 
            // chkTrackHighlightScreen
            // 
            this.chkTrackHighlightScreen.Location = new System.Drawing.Point(435, 205);
            this.chkTrackHighlightScreen.Name = "chkTrackHighlightScreen";
            this.chkTrackHighlightScreen.Size = new System.Drawing.Size(160, 18);
            this.chkTrackHighlightScreen.TabIndex = 6;
            this.chkTrackHighlightScreen.Text = "Highlight changes on screen";
            this.chkTrackHighlightScreen.UseVisualStyleBackColor = true;
            // 
            // chkTrackHighlightPrint
            // 
            this.chkTrackHighlightPrint.Location = new System.Drawing.Point(435, 232);
            this.chkTrackHighlightPrint.Name = "chkTrackHighlightPrint";
            this.chkTrackHighlightPrint.Size = new System.Drawing.Size(160, 18);
            this.chkTrackHighlightPrint.TabIndex = 7;
            this.chkTrackHighlightPrint.Text = "Highlight changes in printed document";
            this.chkTrackHighlightPrint.UseVisualStyleBackColor = true;
            // 
            // lblChangesDisplay
            // 
            this.lblChangesDisplay.Location = new System.Drawing.Point(12, 267);
            this.lblChangesDisplay.Name = "lblChangesDisplay";
            this.lblChangesDisplay.Size = new System.Drawing.Size(100, 18);
            this.lblChangesDisplay.TabIndex = 8;
            this.lblChangesDisplay.Text = "Di&splay for Review:";
            // 
            // cboChangesDisplay
            // 
            this.cboChangesDisplay.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cboChangesDisplay.FormattingEnabled = true;
            this.cboChangesDisplay.Location = new System.Drawing.Point(130, 264);
            this.cboChangesDisplay.Name = "cboChangesDisplay";
            this.cboChangesDisplay.Size = new System.Drawing.Size(170, 21);
            this.cboChangesDisplay.TabIndex = 9;
            this.cboChangesDisplay.SelectedIndexChanged += new System.EventHandler(this.cboChangesDisplay_SelectedIndexChanged);
            // 
            // cmdPasswordDisplay
            // 
            this.cmdPasswordDisplay.BackColor = System.Drawing.SystemColors.Control;
            this.cmdPasswordDisplay.Location = new System.Drawing.Point(283, 317);
            this.cmdPasswordDisplay.Name = "cmdPasswordDisplay";
            this.cmdPasswordDisplay.Size = new System.Drawing.Size(120, 23);
            this.cmdPasswordDisplay.TabIndex = 10;
            this.cmdPasswordDisplay.Text = "&Display Password";
            this.cmdPasswordDisplay.UseVisualStyleBackColor = false;
            this.cmdPasswordDisplay.Click += new System.EventHandler(this.cmdPasswordDisplay_Click);
            // 
            // cmdClose
            // 
            this.cmdClose.BackColor = System.Drawing.SystemColors.Control;
            this.cmdClose.Location = new System.Drawing.Point(505, 317);
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.Size = new System.Drawing.Size(90, 23);
            this.cmdClose.TabIndex = 11;
            this.cmdClose.Text = "&Close";
            this.cmdClose.UseVisualStyleBackColor = false;
            this.cmdClose.Click += new System.EventHandler(this.cmdClose_Click);
            // 
            // FormProtection
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.ClientSize = new System.Drawing.Size(611, 353);
            this.Controls.Add(this.cmdClose);
            this.Controls.Add(this.cmdPasswordDisplay);
            this.Controls.Add(this.cboChangesDisplay);
            this.Controls.Add(this.lblChangesDisplay);
            this.Controls.Add(this.chkTrackHighlightPrint);
            this.Controls.Add(this.chkTrackHighlightScreen);
            this.Controls.Add(this.chkTrackChanges);
            this.Controls.Add(this.chkProtectDocument);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.lblInstructions);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "FormProtection";
            this.Text = "MBR ADMIN: Document Protection and Track Changes";
            this.Activated += new System.EventHandler(this.FormProtection_Activated);
            this.Load += new System.EventHandler(this.FormProtection_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblInstructions;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.CheckBox chkProtectDocument;
        private System.Windows.Forms.CheckBox chkTrackChanges;
        private System.Windows.Forms.CheckBox chkTrackHighlightScreen;
        private System.Windows.Forms.CheckBox chkTrackHighlightPrint;
        private System.Windows.Forms.Label lblChangesDisplay;
        private System.Windows.Forms.ComboBox cboChangesDisplay;
        private System.Windows.Forms.Button cmdPasswordDisplay;
        private System.Windows.Forms.Button cmdClose;
    }
}