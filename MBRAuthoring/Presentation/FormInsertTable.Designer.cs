
namespace MBRAuthoring.Presentation
{
    partial class FormInsertTable
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
            this.tbInsertTable = new System.Windows.Forms.TabPage();
            this.lblTableList = new System.Windows.Forms.Label();
            this.lblTableListTitle = new System.Windows.Forms.Label();
            this.ddlTableList = new System.Windows.Forms.ComboBox();
            this.tabInsertTable = new System.Windows.Forms.TabControl();
            this.btnFinish = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label0 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.shp0 = new System.Windows.Forms.Label();
            this.mapStatusStart = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.tbInsertTable.SuspendLayout();
            this.tabInsertTable.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tbInsertTable
            // 
            this.tbInsertTable.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tbInsertTable.Controls.Add(this.lblTableList);
            this.tbInsertTable.Controls.Add(this.lblTableListTitle);
            this.tbInsertTable.Controls.Add(this.ddlTableList);
            this.tbInsertTable.Location = new System.Drawing.Point(4, 22);
            this.tbInsertTable.Name = "tbInsertTable";
            this.tbInsertTable.Padding = new System.Windows.Forms.Padding(3);
            this.tbInsertTable.Size = new System.Drawing.Size(542, 300);
            this.tbInsertTable.TabIndex = 0;
            this.tbInsertTable.Text = "InsertTable";
            this.tbInsertTable.UseVisualStyleBackColor = true;
            // 
            // lblTableList
            // 
            this.lblTableList.AutoSize = true;
            this.lblTableList.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTableList.Location = new System.Drawing.Point(14, 63);
            this.lblTableList.Name = "lblTableList";
            this.lblTableList.Size = new System.Drawing.Size(182, 13);
            this.lblTableList.TabIndex = 2;
            this.lblTableList.Text = "Select table to insert into Document:";
            // 
            // lblTableListTitle
            // 
            this.lblTableListTitle.AutoSize = true;
            this.lblTableListTitle.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTableListTitle.Location = new System.Drawing.Point(15, 12);
            this.lblTableListTitle.Name = "lblTableListTitle";
            this.lblTableListTitle.Size = new System.Drawing.Size(196, 16);
            this.lblTableListTitle.TabIndex = 1;
            this.lblTableListTitle.Text = "(Title is set in Program Code)";
            // 
            // ddlTableList
            // 
            this.ddlTableList.FormattingEnabled = true;
            this.ddlTableList.Location = new System.Drawing.Point(9, 97);
            this.ddlTableList.Name = "ddlTableList";
            this.ddlTableList.Size = new System.Drawing.Size(519, 21);
            this.ddlTableList.TabIndex = 0;
            this.ddlTableList.SelectedIndexChanged += new System.EventHandler(this.ddlTableList_SelectedIndexChanged);
            // 
            // tabInsertTable
            // 
            this.tabInsertTable.Controls.Add(this.tbInsertTable);
            this.tabInsertTable.Location = new System.Drawing.Point(182, 2);
            this.tabInsertTable.Name = "tabInsertTable";
            this.tabInsertTable.SelectedIndex = 0;
            this.tabInsertTable.Size = new System.Drawing.Size(550, 326);
            this.tabInsertTable.TabIndex = 0;
            // 
            // btnFinish
            // 
            this.btnFinish.Location = new System.Drawing.Point(607, 337);
            this.btnFinish.Name = "btnFinish";
            this.btnFinish.Size = new System.Drawing.Size(92, 31);
            this.btnFinish.TabIndex = 1;
            this.btnFinish.Text = "Finish";
            this.btnFinish.UseVisualStyleBackColor = true;
            this.btnFinish.Click += new System.EventHandler(this.btnFinish_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(320, 337);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(92, 31);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label0);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.shp0);
            this.panel1.Controls.Add(this.mapStatusStart);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Location = new System.Drawing.Point(28, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(139, 226);
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
            this.label0.Location = new System.Drawing.Point(61, 15);
            this.label0.Name = "label0";
            this.label0.Size = new System.Drawing.Size(76, 13);
            this.label0.TabIndex = 29;
            this.label0.Tag = "0";
            this.label0.Text = "Select Table";
            this.label0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Red;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(3, 169);
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
            this.shp0.Location = new System.Drawing.Point(41, 14);
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
            this.label4.Size = new System.Drawing.Size(36, 157);
            this.label4.TabIndex = 23;
            this.label4.Tag = "Exclude";
            // 
            // btnBack
            // 
            this.btnBack.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBack.Location = new System.Drawing.Point(422, 337);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(88, 31);
            this.btnBack.TabIndex = 37;
            this.btnBack.Text = "Back";
            this.btnBack.UseVisualStyleBackColor = true;
            // 
            // btnNext
            // 
            this.btnNext.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNext.Location = new System.Drawing.Point(508, 337);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(88, 31);
            this.btnNext.TabIndex = 36;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            // 
            // FormInsertTable
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(734, 375);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnFinish);
            this.Controls.Add(this.tabInsertTable);
            this.Name = "FormInsertTable";
            this.Text = "Insert Table";
            this.Activated += new System.EventHandler(this.FormInsertTable_Activate);
            this.Load += new System.EventHandler(this.FormInsertTable_Load);
            this.tbInsertTable.ResumeLayout(false);
            this.tbInsertTable.PerformLayout();
            this.tabInsertTable.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tbInsertTable;
        private System.Windows.Forms.Label lblTableList;
        private System.Windows.Forms.Label lblTableListTitle;
        private System.Windows.Forms.ComboBox ddlTableList;
        private System.Windows.Forms.TabControl tabInsertTable;
        private System.Windows.Forms.Button btnFinish;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label0;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label shp0;
        private System.Windows.Forms.Label mapStatusStart;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Button btnNext;
    }
}