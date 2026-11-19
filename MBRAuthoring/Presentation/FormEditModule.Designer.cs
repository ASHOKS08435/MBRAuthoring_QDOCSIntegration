
namespace MBRAuthoring.Presentation
{
    partial class FormEditModule
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
            this.tabEditModule = new System.Windows.Forms.TabControl();
            this.tabModule = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.btnMoveDown = new System.Windows.Forms.Button();
            this.btnMoveUp = new System.Windows.Forms.Button();
            this.btnModuleAdd = new System.Windows.Forms.Button();
            this.btnModuleUpdate = new System.Windows.Forms.Button();
            this.btnModuleDelete = new System.Windows.Forms.Button();
            this.rdoInfoOnly = new System.Windows.Forms.RadioButton();
            this.rdoBatchRecord = new System.Windows.Forms.RadioButton();
            this.txtRepetitions = new System.Windows.Forms.TextBox();
            this.txtModuleTitle = new System.Windows.Forms.TextBox();
            this.ddlPacket = new System.Windows.Forms.ComboBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.ddlModuleAll = new System.Windows.Forms.ComboBox();
            this.dgvModuleSelected = new System.Windows.Forms.DataGridView();
            this.lblModuleSelected = new System.Windows.Forms.Label();
            this.lblModuleTableRows = new System.Windows.Forms.Label();
            this.lblPackets = new System.Windows.Forms.Label();
            this.lblUsage = new System.Windows.Forms.Label();
            this.lblReps = new System.Windows.Forms.Label();
            this.lblModuleTitle = new System.Windows.Forms.Label();
            this.lblModuleAll = new System.Windows.Forms.Label();
            this.SpnUpDown = new System.Windows.Forms.NumericUpDown();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnFinish = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblMap5 = new System.Windows.Forms.Label();
            this.lblMap0 = new System.Windows.Forms.Label();
            this.lblMap6 = new System.Windows.Forms.Label();
            this.shp0 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.cmdBack = new System.Windows.Forms.Button();
            this.cmdNext = new System.Windows.Forms.Button();
            this.chkTableRows = new System.Windows.Forms.CheckBox();
            this.tabEditModule.SuspendLayout();
            this.tabModule.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvModuleSelected)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.SpnUpDown)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEditModule
            // 
            this.tabEditModule.Controls.Add(this.tabModule);
            this.tabEditModule.Location = new System.Drawing.Point(156, 3);
            this.tabEditModule.Name = "tabEditModule";
            this.tabEditModule.SelectedIndex = 0;
            this.tabEditModule.Size = new System.Drawing.Size(551, 495);
            this.tabEditModule.TabIndex = 0;
            // 
            // tabModule
            // 
            this.tabModule.Controls.Add(this.chkTableRows);
            this.tabModule.Controls.Add(this.label1);
            this.tabModule.Controls.Add(this.btnMoveDown);
            this.tabModule.Controls.Add(this.btnMoveUp);
            this.tabModule.Controls.Add(this.btnModuleAdd);
            this.tabModule.Controls.Add(this.btnModuleUpdate);
            this.tabModule.Controls.Add(this.btnModuleDelete);
            this.tabModule.Controls.Add(this.rdoInfoOnly);
            this.tabModule.Controls.Add(this.rdoBatchRecord);
            this.tabModule.Controls.Add(this.txtRepetitions);
            this.tabModule.Controls.Add(this.txtModuleTitle);
            this.tabModule.Controls.Add(this.ddlPacket);
            this.tabModule.Controls.Add(this.lblTitle);
            this.tabModule.Controls.Add(this.ddlModuleAll);
            this.tabModule.Controls.Add(this.dgvModuleSelected);
            this.tabModule.Controls.Add(this.lblModuleSelected);
            this.tabModule.Controls.Add(this.lblModuleTableRows);
            this.tabModule.Controls.Add(this.lblPackets);
            this.tabModule.Controls.Add(this.lblUsage);
            this.tabModule.Controls.Add(this.lblReps);
            this.tabModule.Controls.Add(this.lblModuleTitle);
            this.tabModule.Controls.Add(this.lblModuleAll);
            this.tabModule.Controls.Add(this.SpnUpDown);
            this.tabModule.Location = new System.Drawing.Point(4, 22);
            this.tabModule.Name = "tabModule";
            this.tabModule.Padding = new System.Windows.Forms.Padding(3);
            this.tabModule.Size = new System.Drawing.Size(543, 469);
            this.tabModule.TabIndex = 0;
            this.tabModule.Text = "Module";
            this.tabModule.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(17, 257);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(435, 13);
            this.label1.TabIndex = 24;
            this.label1.Text = " ModuleName        ModuleTitle       Reps       Usage    Packets             Tabl" +
    "eRows";
            // 
            // btnMoveDown
            // 
            this.btnMoveDown.Location = new System.Drawing.Point(252, 430);
            this.btnMoveDown.Name = "btnMoveDown";
            this.btnMoveDown.Size = new System.Drawing.Size(75, 23);
            this.btnMoveDown.TabIndex = 22;
            this.btnMoveDown.Text = "Move Down";
            this.btnMoveDown.UseVisualStyleBackColor = true;
            this.btnMoveDown.Click += new System.EventHandler(this.btnMoveDown_Click);
            // 
            // btnMoveUp
            // 
            this.btnMoveUp.Location = new System.Drawing.Point(168, 431);
            this.btnMoveUp.Name = "btnMoveUp";
            this.btnMoveUp.Size = new System.Drawing.Size(75, 23);
            this.btnMoveUp.TabIndex = 21;
            this.btnMoveUp.Text = "Move Up";
            this.btnMoveUp.UseVisualStyleBackColor = true;
            this.btnMoveUp.Click += new System.EventHandler(this.btnMoveUp_Click);
            // 
            // btnModuleAdd
            // 
            this.btnModuleAdd.Location = new System.Drawing.Point(162, 214);
            this.btnModuleAdd.Name = "btnModuleAdd";
            this.btnModuleAdd.Size = new System.Drawing.Size(75, 23);
            this.btnModuleAdd.TabIndex = 19;
            this.btnModuleAdd.Text = "Add";
            this.btnModuleAdd.UseVisualStyleBackColor = true;
            this.btnModuleAdd.Click += new System.EventHandler(this.btnModuleAdd_Click);
            // 
            // btnModuleUpdate
            // 
            this.btnModuleUpdate.Location = new System.Drawing.Point(243, 214);
            this.btnModuleUpdate.Name = "btnModuleUpdate";
            this.btnModuleUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnModuleUpdate.TabIndex = 18;
            this.btnModuleUpdate.Text = "Update";
            this.btnModuleUpdate.UseVisualStyleBackColor = true;
            this.btnModuleUpdate.Click += new System.EventHandler(this.btnModuleUpdate_Click);
            // 
            // btnModuleDelete
            // 
            this.btnModuleDelete.Location = new System.Drawing.Point(325, 214);
            this.btnModuleDelete.Name = "btnModuleDelete";
            this.btnModuleDelete.Size = new System.Drawing.Size(75, 23);
            this.btnModuleDelete.TabIndex = 17;
            this.btnModuleDelete.Text = "Delete";
            this.btnModuleDelete.UseVisualStyleBackColor = true;
            this.btnModuleDelete.Click += new System.EventHandler(this.btnModuleDelete_Click);
            // 
            // rdoInfoOnly
            // 
            this.rdoInfoOnly.AutoSize = true;
            this.rdoInfoOnly.Location = new System.Drawing.Point(223, 132);
            this.rdoInfoOnly.Name = "rdoInfoOnly";
            this.rdoInfoOnly.Size = new System.Drawing.Size(119, 17);
            this.rdoInfoOnly.TabIndex = 15;
            this.rdoInfoOnly.TabStop = true;
            this.rdoInfoOnly.Text = "For Information Only";
            this.rdoInfoOnly.UseVisualStyleBackColor = true;
            // 
            // rdoBatchRecord
            // 
            this.rdoBatchRecord.AutoSize = true;
            this.rdoBatchRecord.Location = new System.Drawing.Point(109, 131);
            this.rdoBatchRecord.Name = "rdoBatchRecord";
            this.rdoBatchRecord.Size = new System.Drawing.Size(109, 17);
            this.rdoBatchRecord.TabIndex = 14;
            this.rdoBatchRecord.TabStop = true;
            this.rdoBatchRecord.Text = "For Batch Record";
            this.rdoBatchRecord.UseVisualStyleBackColor = true;
            // 
            // txtRepetitions
            // 
            this.txtRepetitions.Location = new System.Drawing.Point(106, 101);
            this.txtRepetitions.Multiline = true;
            this.txtRepetitions.Name = "txtRepetitions";
            this.txtRepetitions.Size = new System.Drawing.Size(57, 24);
            this.txtRepetitions.TabIndex = 13;
            this.txtRepetitions.Text = "1";
            // 
            // txtModuleTitle
            // 
            this.txtModuleTitle.Location = new System.Drawing.Point(106, 69);
            this.txtModuleTitle.Multiline = true;
            this.txtModuleTitle.Name = "txtModuleTitle";
            this.txtModuleTitle.Size = new System.Drawing.Size(350, 24);
            this.txtModuleTitle.TabIndex = 12;
            // 
            // ddlPacket
            // 
            this.ddlPacket.FormattingEnabled = true;
            this.ddlPacket.Location = new System.Drawing.Point(106, 157);
            this.ddlPacket.Name = "ddlPacket";
            this.ddlPacket.Size = new System.Drawing.Size(262, 21);
            this.ddlPacket.TabIndex = 10;
            this.ddlPacket.SelectedIndexChanged += new System.EventHandler(this.ddlPacket_SelectedIndexChanged);
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(23, 15);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(43, 14);
            this.lblTitle.TabIndex = 9;
            this.lblTitle.Text = "label1";
            // 
            // ddlModuleAll
            // 
            this.ddlModuleAll.FormattingEnabled = true;
            this.ddlModuleAll.Location = new System.Drawing.Point(105, 38);
            this.ddlModuleAll.Name = "ddlModuleAll";
            this.ddlModuleAll.Size = new System.Drawing.Size(350, 21);
            this.ddlModuleAll.TabIndex = 8;
            this.ddlModuleAll.Text = "ddlModuleAll";
            this.ddlModuleAll.SelectedIndexChanged += new System.EventHandler(this.ddlModuleAll_SelectedIndexChanged);
            // 
            // dgvModuleSelected
            // 
            this.dgvModuleSelected.AllowUserToAddRows = false;
            this.dgvModuleSelected.BackgroundColor = System.Drawing.Color.White;
            this.dgvModuleSelected.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgvModuleSelected.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.dgvModuleSelected.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvModuleSelected.ColumnHeadersVisible = false;
            this.dgvModuleSelected.Location = new System.Drawing.Point(20, 286);
            this.dgvModuleSelected.MultiSelect = false;
            this.dgvModuleSelected.Name = "dgvModuleSelected";
            this.dgvModuleSelected.ReadOnly = true;
            this.dgvModuleSelected.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.dgvModuleSelected.RowHeadersVisible = false;
            this.dgvModuleSelected.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvModuleSelected.Size = new System.Drawing.Size(517, 139);
            this.dgvModuleSelected.TabIndex = 7;
            this.dgvModuleSelected.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvModuleSelected_CellDoubleClick);
            // 
            // lblModuleSelected
            // 
            this.lblModuleSelected.AutoSize = true;
            this.lblModuleSelected.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblModuleSelected.Location = new System.Drawing.Point(17, 241);
            this.lblModuleSelected.Name = "lblModuleSelected";
            this.lblModuleSelected.Size = new System.Drawing.Size(133, 13);
            this.lblModuleSelected.TabIndex = 6;
            this.lblModuleSelected.Text = "Modules selected for : ";
            // 
            // lblModuleTableRows
            // 
            this.lblModuleTableRows.AutoSize = true;
            this.lblModuleTableRows.Location = new System.Drawing.Point(6, 189);
            this.lblModuleTableRows.Name = "lblModuleTableRows";
            this.lblModuleTableRows.Size = new System.Drawing.Size(67, 13);
            this.lblModuleTableRows.TabIndex = 5;
            this.lblModuleTableRows.Text = "Table Rows \n";
            // 
            // lblPackets
            // 
            this.lblPackets.AutoSize = true;
            this.lblPackets.Location = new System.Drawing.Point(51, 164);
            this.lblPackets.Name = "lblPackets";
            this.lblPackets.Size = new System.Drawing.Size(46, 13);
            this.lblPackets.TabIndex = 4;
            this.lblPackets.Text = "Packets";
            // 
            // lblUsage
            // 
            this.lblUsage.AutoSize = true;
            this.lblUsage.Location = new System.Drawing.Point(59, 134);
            this.lblUsage.Name = "lblUsage";
            this.lblUsage.Size = new System.Drawing.Size(38, 13);
            this.lblUsage.TabIndex = 3;
            this.lblUsage.Text = "Usage";
            // 
            // lblReps
            // 
            this.lblReps.AutoSize = true;
            this.lblReps.Location = new System.Drawing.Point(37, 107);
            this.lblReps.Name = "lblReps";
            this.lblReps.Size = new System.Drawing.Size(60, 13);
            this.lblReps.TabIndex = 2;
            this.lblReps.Text = "Repetitions";
            // 
            // lblModuleTitle
            // 
            this.lblModuleTitle.AutoSize = true;
            this.lblModuleTitle.Location = new System.Drawing.Point(32, 77);
            this.lblModuleTitle.Name = "lblModuleTitle";
            this.lblModuleTitle.Size = new System.Drawing.Size(65, 13);
            this.lblModuleTitle.TabIndex = 1;
            this.lblModuleTitle.Text = "Module Title";
            // 
            // lblModuleAll
            // 
            this.lblModuleAll.AutoSize = true;
            this.lblModuleAll.Location = new System.Drawing.Point(24, 45);
            this.lblModuleAll.Name = "lblModuleAll";
            this.lblModuleAll.Size = new System.Drawing.Size(73, 13);
            this.lblModuleAll.TabIndex = 0;
            this.lblModuleAll.Text = "Module Name";
            // 
            // SpnUpDown
            // 
            this.SpnUpDown.Location = new System.Drawing.Point(151, 187);
            this.SpnUpDown.Name = "SpnUpDown";
            this.SpnUpDown.Size = new System.Drawing.Size(67, 20);
            this.SpnUpDown.TabIndex = 20;
            this.SpnUpDown.ValueChanged += new System.EventHandler(this.SpnUpDown_ValueChanged);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(238, 504);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnFinish
            // 
            this.btnFinish.Location = new System.Drawing.Point(493, 504);
            this.btnFinish.Name = "btnFinish";
            this.btnFinish.Size = new System.Drawing.Size(75, 23);
            this.btnFinish.TabIndex = 2;
            this.btnFinish.Text = "Finish";
            this.btnFinish.UseVisualStyleBackColor = true;
            this.btnFinish.Click += new System.EventHandler(this.btnFinish_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.lblMap5);
            this.panel1.Controls.Add(this.lblMap0);
            this.panel1.Controls.Add(this.lblMap6);
            this.panel1.Controls.Add(this.shp0);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Location = new System.Drawing.Point(4, 29);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(150, 228);
            this.panel1.TabIndex = 34;
            // 
            // lblMap5
            // 
            this.lblMap5.BackColor = System.Drawing.Color.Black;
            this.lblMap5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMap5.ForeColor = System.Drawing.Color.White;
            this.lblMap5.Location = new System.Drawing.Point(59, 190);
            this.lblMap5.Name = "lblMap5";
            this.lblMap5.Size = new System.Drawing.Size(37, 26);
            this.lblMap5.TabIndex = 29;
            this.lblMap5.Tag = "0";
            this.lblMap5.Text = "Finish";
            this.lblMap5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblMap5.UseMnemonic = false;
            this.lblMap5.Click += new System.EventHandler(this.lblMap5_Click);
            // 
            // lblMap0
            // 
            this.lblMap0.BackColor = System.Drawing.Color.Black;
            this.lblMap0.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMap0.ForeColor = System.Drawing.Color.White;
            this.lblMap0.Location = new System.Drawing.Point(59, 0);
            this.lblMap0.Name = "lblMap0";
            this.lblMap0.Size = new System.Drawing.Size(72, 37);
            this.lblMap0.TabIndex = 28;
            this.lblMap0.Tag = "0";
            this.lblMap0.Text = "Edit Modules";
            this.lblMap0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblMap0.UseMnemonic = false;
            // 
            // lblMap6
            // 
            this.lblMap6.BackColor = System.Drawing.Color.Red;
            this.lblMap6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMap6.Location = new System.Drawing.Point(8, 201);
            this.lblMap6.Name = "lblMap6";
            this.lblMap6.Size = new System.Drawing.Size(15, 15);
            this.lblMap6.TabIndex = 27;
            this.lblMap6.Tag = "Exclude";
            this.lblMap6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // shp0
            // 
            this.shp0.BackColor = System.Drawing.Color.Lime;
            this.shp0.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.shp0.Location = new System.Drawing.Point(38, 10);
            this.shp0.Name = "shp0";
            this.shp0.Size = new System.Drawing.Size(15, 15);
            this.shp0.TabIndex = 25;
            this.shp0.Tag = "0";
            this.shp0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label7.Location = new System.Drawing.Point(18, 19);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(27, 191);
            this.label7.TabIndex = 23;
            this.label7.Tag = "Exclude";
            // 
            // cmdBack
            // 
            this.cmdBack.BackColor = System.Drawing.SystemColors.Control;
            this.cmdBack.Enabled = false;
            this.cmdBack.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdBack.Location = new System.Drawing.Point(329, 504);
            this.cmdBack.Name = "cmdBack";
            this.cmdBack.Size = new System.Drawing.Size(75, 23);
            this.cmdBack.TabIndex = 36;
            this.cmdBack.Text = "< &Back";
            this.cmdBack.UseVisualStyleBackColor = false;
            // 
            // cmdNext
            // 
            this.cmdNext.BackColor = System.Drawing.SystemColors.Control;
            this.cmdNext.Enabled = false;
            this.cmdNext.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdNext.Location = new System.Drawing.Point(412, 504);
            this.cmdNext.Name = "cmdNext";
            this.cmdNext.Size = new System.Drawing.Size(75, 23);
            this.cmdNext.TabIndex = 35;
            this.cmdNext.Text = "&Next >";
            this.cmdNext.UseVisualStyleBackColor = false;
            // 
            // chkTableRows
            // 
            this.chkTableRows.AutoSize = true;
            this.chkTableRows.Location = new System.Drawing.Point(243, 187);
            this.chkTableRows.Name = "chkTableRows";
            this.chkTableRows.Size = new System.Drawing.Size(90, 17);
            this.chkTableRows.TabIndex = 25;
            this.chkTableRows.Text = "P = Pallet Qty";
            this.chkTableRows.UseVisualStyleBackColor = true;
            // 
            // FormEditModule
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(714, 539);
            this.Controls.Add(this.cmdBack);
            this.Controls.Add(this.cmdNext);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnFinish);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.tabEditModule);
            this.Name = "FormEditModule";
            this.Text = "Edit Modules Wizard";
            this.Load += new System.EventHandler(this.FormEditModule_Load);
            this.Shown += new System.EventHandler(this.FormEditModule_Shown);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormEditModule_FormClosed);
            this.tabEditModule.ResumeLayout(false);
            this.tabModule.ResumeLayout(false);
            this.tabModule.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvModuleSelected)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.SpnUpDown)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabEditModule;
        private System.Windows.Forms.TabPage tabModule;
        private System.Windows.Forms.Label lblPackets;
        private System.Windows.Forms.Label lblUsage;
        private System.Windows.Forms.Label lblReps;
        private System.Windows.Forms.Label lblModuleTitle;
        private System.Windows.Forms.Label lblModuleAll;
        private System.Windows.Forms.Label lblModuleSelected;
        private System.Windows.Forms.Label lblModuleTableRows;
        private System.Windows.Forms.DataGridView dgvModuleSelected;
        private System.Windows.Forms.ComboBox ddlModuleAll;
        private System.Windows.Forms.RadioButton rdoInfoOnly;
        private System.Windows.Forms.RadioButton rdoBatchRecord;
        private System.Windows.Forms.TextBox txtRepetitions;
        private System.Windows.Forms.TextBox txtModuleTitle;
        private System.Windows.Forms.ComboBox ddlPacket;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button btnModuleAdd;
        private System.Windows.Forms.Button btnModuleUpdate;
        private System.Windows.Forms.Button btnModuleDelete;
        private System.Windows.Forms.NumericUpDown SpnUpDown;
        private System.Windows.Forms.Button btnMoveDown;
        private System.Windows.Forms.Button btnMoveUp;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnFinish;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblMap5;
        private System.Windows.Forms.Label lblMap0;
        private System.Windows.Forms.Label lblMap6;
        private System.Windows.Forms.Label shp0;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button cmdBack;
        private System.Windows.Forms.Button cmdNext;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkTableRows;
    }
}