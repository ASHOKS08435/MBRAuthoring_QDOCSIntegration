
namespace MBRAuthoring.Presentation
{
    partial class FormMBRHelp
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMBRHelp));
            this.lblMBRHelp = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblMBRHelp
            // 
            this.lblMBRHelp.AutoSize = true;
            this.lblMBRHelp.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.lblMBRHelp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblMBRHelp.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMBRHelp.Location = new System.Drawing.Point(12, 18);
            this.lblMBRHelp.Name = "lblMBRHelp";
            this.lblMBRHelp.Size = new System.Drawing.Size(497, 258);
            this.lblMBRHelp.TabIndex = 0;
            this.lblMBRHelp.Text = resources.GetString("lblMBRHelp.Text");
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(190, 298);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(102, 28);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // FormMBRHelp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(576, 338);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblMBRHelp);
            this.Name = "FormMBRHelp";
            this.Text = "MBR Help";
            this.Load += new System.EventHandler(this.FormMBRHelp_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblMBRHelp;
        private System.Windows.Forms.Button btnOK;
    }
}