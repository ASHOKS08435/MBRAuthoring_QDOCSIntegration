
namespace MBRAuthoring.Presentation
{
    partial class FormInsertField
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label0 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.shp0 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.tabProperty = new System.Windows.Forms.TabControl();
            this.tbProperty = new System.Windows.Forms.TabPage();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblDocProperty = new System.Windows.Forms.Label();
            this.cmdPropertyInsert = new System.Windows.Forms.Button();
            this.lblPropertyName = new System.Windows.Forms.Label();
            this.dgvProperty = new System.Windows.Forms.DataGridView();
            this.lblPropertyValue = new System.Windows.Forms.Label();
            this.btnFinish = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.tabProperty.SuspendLayout();
            this.tbProperty.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProperty)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label0);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.shp0);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Location = new System.Drawing.Point(24, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(126, 226);
            this.panel1.TabIndex = 32;
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
            this.label0.Location = new System.Drawing.Point(46, 15);
            this.label0.Name = "label0";
            this.label0.Size = new System.Drawing.Size(71, 13);
            this.label0.TabIndex = 29;
            this.label0.Tag = "0";
            this.label0.Text = "Insert Field";
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
            this.shp0.Location = new System.Drawing.Point(28, 12);
            this.shp0.Name = "shp0";
            this.shp0.Size = new System.Drawing.Size(15, 15);
            this.shp0.TabIndex = 25;
            this.shp0.Tag = "0";
            this.shp0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label4.Location = new System.Drawing.Point(11, 22);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(25, 157);
            this.label4.TabIndex = 23;
            this.label4.Tag = "Exclude";
            // 
            // tabProperty
            // 
            this.tabProperty.Controls.Add(this.tbProperty);
            this.tabProperty.Location = new System.Drawing.Point(168, 3);
            this.tabProperty.Name = "tabProperty";
            this.tabProperty.SelectedIndex = 0;
            this.tabProperty.Size = new System.Drawing.Size(606, 393);
            this.tabProperty.TabIndex = 11;
            // 
            // tbProperty
            // 
            this.tbProperty.Controls.Add(this.lblTitle);
            this.tbProperty.Controls.Add(this.lblDocProperty);
            this.tbProperty.Controls.Add(this.cmdPropertyInsert);
            this.tbProperty.Controls.Add(this.lblPropertyName);
            this.tbProperty.Controls.Add(this.dgvProperty);
            this.tbProperty.Controls.Add(this.lblPropertyValue);
            this.tbProperty.Location = new System.Drawing.Point(4, 22);
            this.tbProperty.Name = "tbProperty";
            this.tbProperty.Padding = new System.Windows.Forms.Padding(3);
            this.tbProperty.Size = new System.Drawing.Size(598, 367);
            this.tbProperty.TabIndex = 0;
            this.tbProperty.Text = "tabPoperty";
            this.tbProperty.UseVisualStyleBackColor = true;
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(18, 13);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(193, 16);
            this.lblTitle.TabIndex = 1;
            this.lblTitle.Text = "(Title set in program code...)";
            // 
            // lblDocProperty
            // 
            this.lblDocProperty.AutoSize = true;
            this.lblDocProperty.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDocProperty.Location = new System.Drawing.Point(21, 41);
            this.lblDocProperty.Name = "lblDocProperty";
            this.lblDocProperty.Size = new System.Drawing.Size(114, 13);
            this.lblDocProperty.TabIndex = 2;
            this.lblDocProperty.Text = "Document Properties :";
            // 
            // cmdPropertyInsert
            // 
            this.cmdPropertyInsert.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdPropertyInsert.Location = new System.Drawing.Point(52, 321);
            this.cmdPropertyInsert.Name = "cmdPropertyInsert";
            this.cmdPropertyInsert.Size = new System.Drawing.Size(537, 31);
            this.cmdPropertyInsert.TabIndex = 7;
            this.cmdPropertyInsert.Text = "Insert Field";
            this.cmdPropertyInsert.UseVisualStyleBackColor = true;
            this.cmdPropertyInsert.Click += new System.EventHandler(this.cmdPropertyInsert_Click);
            // 
            // lblPropertyName
            // 
            this.lblPropertyName.AutoSize = true;
            this.lblPropertyName.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPropertyName.Location = new System.Drawing.Point(142, 61);
            this.lblPropertyName.Name = "lblPropertyName";
            this.lblPropertyName.Size = new System.Drawing.Size(92, 13);
            this.lblPropertyName.TabIndex = 3;
            this.lblPropertyName.Text = "Property Name";
            // 
            // dgvProperty
            // 
            this.dgvProperty.AllowUserToAddRows = false;
            this.dgvProperty.AllowUserToDeleteRows = false;
            this.dgvProperty.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dgvProperty.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvProperty.Location = new System.Drawing.Point(52, 79);
            this.dgvProperty.MultiSelect = false;
            this.dgvProperty.Name = "dgvProperty";
            this.dgvProperty.ReadOnly = true;
            this.dgvProperty.Size = new System.Drawing.Size(537, 236);
            this.dgvProperty.TabIndex = 8;
            // 
            // lblPropertyValue
            // 
            this.lblPropertyValue.AutoSize = true;
            this.lblPropertyValue.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPropertyValue.Location = new System.Drawing.Point(376, 60);
            this.lblPropertyValue.Name = "lblPropertyValue";
            this.lblPropertyValue.Size = new System.Drawing.Size(91, 13);
            this.lblPropertyValue.TabIndex = 4;
            this.lblPropertyValue.Text = "Property Value";
            // 
            // btnFinish
            // 
            this.btnFinish.Location = new System.Drawing.Point(626, 411);
            this.btnFinish.Name = "btnFinish";
            this.btnFinish.Size = new System.Drawing.Size(75, 32);
            this.btnFinish.TabIndex = 10;
            this.btnFinish.Text = "Finish";
            this.btnFinish.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(347, 411);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 32);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "Close";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnBack
            // 
            this.btnBack.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBack.Location = new System.Drawing.Point(441, 412);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(88, 31);
            this.btnBack.TabIndex = 37;
            this.btnBack.Text = "Back";
            this.btnBack.UseVisualStyleBackColor = true;
            // 
            // btnNext
            // 
            this.btnNext.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNext.Location = new System.Drawing.Point(530, 412);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(88, 31);
            this.btnNext.TabIndex = 36;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            // 
            // FormInsertField
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(775, 450);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.tabProperty);
            this.Controls.Add(this.btnFinish);
            this.Controls.Add(this.btnCancel);
            this.Name = "FormInsertField";
            this.Text = "Insert Field Wizard";
            this.Activated += new System.EventHandler(this.FormInsertField_Activate);
            this.Load += new System.EventHandler(this.FormInsertField_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabProperty.ResumeLayout(false);
            this.tbProperty.ResumeLayout(false);
            this.tbProperty.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProperty)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblDocProperty;
        private System.Windows.Forms.Label lblPropertyName;
        private System.Windows.Forms.Label lblPropertyValue;
        private System.Windows.Forms.Button cmdPropertyInsert;
        private System.Windows.Forms.DataGridView dgvProperty;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnFinish;
        private System.Windows.Forms.TabPage tbProperty;
        private System.Windows.Forms.TabControl tabProperty;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label0;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label shp0;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Button btnNext;
    }
}