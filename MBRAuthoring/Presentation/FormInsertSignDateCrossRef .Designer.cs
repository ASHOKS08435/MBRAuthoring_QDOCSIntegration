
namespace MBRAuthoring.Presentation
{
    partial class FormInsertSignDateCrossRef
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
            this.tabStepCrossRef = new System.Windows.Forms.TabControl();
            this.tbStepCrossRef = new System.Windows.Forms.TabPage();
            this.lstNumberedItems = new System.Windows.Forms.ListBox();
            this.optList1Thru3 = new System.Windows.Forms.RadioButton();
            this.optList123 = new System.Windows.Forms.RadioButton();
            this.btnResetSelection = new System.Windows.Forms.Button();
            this.lblNumberedItemSelected = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnFinish = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label0 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.shp0 = new System.Windows.Forms.Label();
            this.mapStatusStart = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.tabStepCrossRef.SuspendLayout();
            this.tbStepCrossRef.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabStepCrossRef
            // 
            this.tabStepCrossRef.Controls.Add(this.tbStepCrossRef);
            this.tabStepCrossRef.Location = new System.Drawing.Point(144, 0);
            this.tabStepCrossRef.Name = "tabStepCrossRef";
            this.tabStepCrossRef.SelectedIndex = 0;
            this.tabStepCrossRef.Size = new System.Drawing.Size(637, 360);
            this.tabStepCrossRef.TabIndex = 0;
            // 
            // tbStepCrossRef
            // 
            this.tbStepCrossRef.Controls.Add(this.lstNumberedItems);
            this.tbStepCrossRef.Controls.Add(this.optList1Thru3);
            this.tbStepCrossRef.Controls.Add(this.optList123);
            this.tbStepCrossRef.Controls.Add(this.btnResetSelection);
            this.tbStepCrossRef.Controls.Add(this.lblNumberedItemSelected);
            this.tbStepCrossRef.Controls.Add(this.lblTitle);
            this.tbStepCrossRef.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbStepCrossRef.Location = new System.Drawing.Point(4, 22);
            this.tbStepCrossRef.Name = "tbStepCrossRef";
            this.tbStepCrossRef.Padding = new System.Windows.Forms.Padding(3);
            this.tbStepCrossRef.Size = new System.Drawing.Size(629, 334);
            this.tbStepCrossRef.TabIndex = 0;
            this.tbStepCrossRef.Text = "StepCrossRef";
            this.tbStepCrossRef.UseVisualStyleBackColor = true;
            // 
            // lstNumberedItems
            // 
            this.lstNumberedItems.FormattingEnabled = true;
            this.lstNumberedItems.Location = new System.Drawing.Point(22, 70);
            this.lstNumberedItems.Name = "lstNumberedItems";
            this.lstNumberedItems.Size = new System.Drawing.Size(412, 199);
            this.lstNumberedItems.TabIndex = 6;
            this.lstNumberedItems.Click += new System.EventHandler(this.lstNumberedItems_Click);
            this.lstNumberedItems.DoubleClick += new System.EventHandler(this.lstNumberedItems_DoubleClick);
            this.lstNumberedItems.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lstNumberedItems_KeyDown);
            this.lstNumberedItems.KeyUp += new System.Windows.Forms.KeyEventHandler(this.lstNumberedItems_KeyUp);
            // 
            // optList1Thru3
            // 
            this.optList1Thru3.AutoSize = true;
            this.optList1Thru3.Location = new System.Drawing.Point(347, 312);
            this.optList1Thru3.Name = "optList1Thru3";
            this.optList1Thru3.Size = new System.Drawing.Size(182, 17);
            this.optList1Thru3.TabIndex = 5;
            this.optList1Thru3.TabStop = true;
            this.optList1Thru3.Text = "Display as Step 1 through Step 3";
            this.optList1Thru3.UseVisualStyleBackColor = true;
            // 
            // optList123
            // 
            this.optList123.AutoSize = true;
            this.optList123.Location = new System.Drawing.Point(17, 312);
            this.optList123.Name = "optList123";
            this.optList123.Size = new System.Drawing.Size(183, 17);
            this.optList123.TabIndex = 4;
            this.optList123.TabStop = true;
            this.optList123.Text = "Display as Step 1, Step 2, Step 3";
            this.optList123.UseVisualStyleBackColor = true;
            // 
            // btnResetSelection
            // 
            this.btnResetSelection.Location = new System.Drawing.Point(168, 277);
            this.btnResetSelection.Name = "btnResetSelection";
            this.btnResetSelection.Size = new System.Drawing.Size(171, 30);
            this.btnResetSelection.TabIndex = 3;
            this.btnResetSelection.Text = "Reset Selection(s)";
            this.btnResetSelection.UseVisualStyleBackColor = true;
            this.btnResetSelection.Click += new System.EventHandler(this.btnResetSelection_Click);
            // 
            // lblNumberedItemSelected
            // 
            this.lblNumberedItemSelected.AutoSize = true;
            this.lblNumberedItemSelected.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNumberedItemSelected.Location = new System.Drawing.Point(22, 51);
            this.lblNumberedItemSelected.Name = "lblNumberedItemSelected";
            this.lblNumberedItemSelected.Size = new System.Drawing.Size(190, 13);
            this.lblNumberedItemSelected.TabIndex = 1;
            this.lblNumberedItemSelected.Text = "Numbered Items for Cross Reference:";
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(22, 20);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(193, 16);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "(Title set in program code...)";
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(332, 381);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(85, 29);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Close";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnFinish
            // 
            this.btnFinish.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFinish.Location = new System.Drawing.Point(600, 380);
            this.btnFinish.Name = "btnFinish";
            this.btnFinish.Size = new System.Drawing.Size(79, 29);
            this.btnFinish.TabIndex = 2;
            this.btnFinish.Text = "Finish";
            this.btnFinish.UseVisualStyleBackColor = true;
            this.btnFinish.Click += new System.EventHandler(this.btnFinish_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label0);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.shp0);
            this.panel1.Controls.Add(this.mapStatusStart);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Location = new System.Drawing.Point(12, 22);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(126, 226);
            this.panel1.TabIndex = 34;
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
            this.label0.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label0.ForeColor = System.Drawing.Color.White;
            this.label0.Location = new System.Drawing.Point(61, 15);
            this.label0.Name = "label0";
            this.label0.Size = new System.Drawing.Size(61, 13);
            this.label0.TabIndex = 29;
            this.label0.Tag = "0";
            this.label0.Text = "Select Step";
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
            // btnBack
            // 
            this.btnBack.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBack.Location = new System.Drawing.Point(421, 380);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(88, 30);
            this.btnBack.TabIndex = 37;
            this.btnBack.Text = "Back";
            this.btnBack.UseVisualStyleBackColor = true;
            // 
            // btnNext
            // 
            this.btnNext.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNext.Location = new System.Drawing.Point(510, 380);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(88, 29);
            this.btnNext.TabIndex = 36;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            // 
            // FormInsertSignDateCrossRef
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(681, 417);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnFinish);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.tabStepCrossRef);
            this.Name = "FormInsertSignDateCrossRef";
            this.Text = "Select & Insert Cross Reference Step";
            this.Activated += new System.EventHandler(this.FormInsertSignDateCrossRef_Activate);
            this.Load += new System.EventHandler(this.FormInsertSignDateCrossRef_Load);
            this.tabStepCrossRef.ResumeLayout(false);
            this.tbStepCrossRef.ResumeLayout(false);
            this.tbStepCrossRef.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabStepCrossRef;
        private System.Windows.Forms.TabPage tbStepCrossRef;
        private System.Windows.Forms.RadioButton optList1Thru3;
        private System.Windows.Forms.RadioButton optList123;
        private System.Windows.Forms.Button btnResetSelection;
        private System.Windows.Forms.Label lblNumberedItemSelected;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnFinish;
        private System.Windows.Forms.ListBox lstNumberedItems;
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