using System;
using System.Windows.Forms;

namespace MBRAuthoring.Presentation
{
    partial class FormEditProperty
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
            this.tabEditProperty = new System.Windows.Forms.TabControl();
            this.tabEdit = new System.Windows.Forms.TabPage();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblUpdateMessage = new System.Windows.Forms.Label();
            this.lblReqCount = new System.Windows.Forms.Label();
            this.lblSpaceCount = new System.Windows.Forms.Label();
            this.lblSpace = new System.Windows.Forms.Label();
            this.lblPromptDoubleClick = new System.Windows.Forms.Label();
            this.btnPropertyDelete = new System.Windows.Forms.Button();
            this.lblDocProperty = new System.Windows.Forms.Label();
            this.chkShowModules = new System.Windows.Forms.CheckBox();
            this.chkShowAllProperties = new System.Windows.Forms.CheckBox();
            this.dgvDocProperty = new System.Windows.Forms.DataGridView();
            this.btnDescribe = new System.Windows.Forms.Button();
            this.btnPropertyUpdate = new System.Windows.Forms.Button();
            this.btnPropertyAdd = new System.Windows.Forms.Button();
            this.cmdPropertyValueSize = new System.Windows.Forms.Button();
            this.cmdPropertyValueBracket = new System.Windows.Forms.Button();
            this.ddlPropertyUOM = new System.Windows.Forms.ComboBox();
            this.ddlPropertyValue = new System.Windows.Forms.ComboBox();
            this.txtPropertyValue = new System.Windows.Forms.TextBox();
            this.ddlPropertyAll = new System.Windows.Forms.ComboBox();
            this.ddlPropertyCategory = new System.Windows.Forms.ComboBox();
            this.spnPropertyIndex = new System.Windows.Forms.NumericUpDown();
            this.lblPropertyUOM = new System.Windows.Forms.Label();
            this.lblPropertyValue = new System.Windows.Forms.Label();
            this.lblPropertyIndex = new System.Windows.Forms.Label();
            this.lblPropertyAll = new System.Windows.Forms.Label();
            this.lblPropertyCategory = new System.Windows.Forms.Label();
            this.tabSource = new System.Windows.Forms.TabPage();
            this.lblSourceHeader = new System.Windows.Forms.Label();
            this.lblSourcepty = new System.Windows.Forms.Label();
            this.btnSyncSysProperties = new System.Windows.Forms.Button();
            this.dgvPropertySource = new System.Windows.Forms.DataGridView();
            this.tabUnused = new System.Windows.Forms.TabPage();
            this.lblPropertyHeader = new System.Windows.Forms.Label();
            this.lblUnusedpty = new System.Windows.Forms.Label();
            this.btnUnusedPropertyDeleteAll = new System.Windows.Forms.Button();
            this.btnUnusedPropertyDelete = new System.Windows.Forms.Button();
            this.dgvDocPropertyUnused = new System.Windows.Forms.DataGridView();
            this.tabMessages = new System.Windows.Forms.TabPage();
            this.lblMessages = new System.Windows.Forms.Label();
            this.lstMessages = new System.Windows.Forms.ListBox();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnFinish = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnBack = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.shp2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.shp3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label0 = new System.Windows.Forms.Label();
            this.shp1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.shp0 = new System.Windows.Forms.Label();
            this.mapStatusStart = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.tabEditProperty.SuspendLayout();
            this.tabEdit.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDocProperty)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.spnPropertyIndex)).BeginInit();
            this.tabSource.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPropertySource)).BeginInit();
            this.tabUnused.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDocPropertyUnused)).BeginInit();
            this.tabMessages.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEditProperty
            // 
            this.tabEditProperty.Controls.Add(this.tabEdit);
            this.tabEditProperty.Controls.Add(this.tabSource);
            this.tabEditProperty.Controls.Add(this.tabUnused);
            this.tabEditProperty.Controls.Add(this.tabMessages);
            this.tabEditProperty.Location = new System.Drawing.Point(175, -2);
            this.tabEditProperty.Name = "tabEditProperty";
            this.tabEditProperty.SelectedIndex = 0;
            this.tabEditProperty.Size = new System.Drawing.Size(623, 466);
            this.tabEditProperty.TabIndex = 0;
            // 
            // tabEdit
            // 
            this.tabEdit.Controls.Add(this.lblTitle);
            this.tabEdit.Controls.Add(this.lblUpdateMessage);
            this.tabEdit.Controls.Add(this.lblReqCount);
            this.tabEdit.Controls.Add(this.lblSpaceCount);
            this.tabEdit.Controls.Add(this.lblSpace);
            this.tabEdit.Controls.Add(this.lblPromptDoubleClick);
            this.tabEdit.Controls.Add(this.btnPropertyDelete);
            this.tabEdit.Controls.Add(this.lblDocProperty);
            this.tabEdit.Controls.Add(this.chkShowModules);
            this.tabEdit.Controls.Add(this.chkShowAllProperties);
            this.tabEdit.Controls.Add(this.dgvDocProperty);
            this.tabEdit.Controls.Add(this.btnDescribe);
            this.tabEdit.Controls.Add(this.btnPropertyUpdate);
            this.tabEdit.Controls.Add(this.btnPropertyAdd);
            this.tabEdit.Controls.Add(this.cmdPropertyValueSize);
            this.tabEdit.Controls.Add(this.cmdPropertyValueBracket);
            this.tabEdit.Controls.Add(this.ddlPropertyUOM);
            this.tabEdit.Controls.Add(this.ddlPropertyValue);
            this.tabEdit.Controls.Add(this.txtPropertyValue);
            this.tabEdit.Controls.Add(this.ddlPropertyAll);
            this.tabEdit.Controls.Add(this.ddlPropertyCategory);
            this.tabEdit.Controls.Add(this.spnPropertyIndex);
            this.tabEdit.Controls.Add(this.lblPropertyUOM);
            this.tabEdit.Controls.Add(this.lblPropertyValue);
            this.tabEdit.Controls.Add(this.lblPropertyIndex);
            this.tabEdit.Controls.Add(this.lblPropertyAll);
            this.tabEdit.Controls.Add(this.lblPropertyCategory);
            this.tabEdit.Location = new System.Drawing.Point(4, 22);
            this.tabEdit.Name = "tabEdit";
            this.tabEdit.Padding = new System.Windows.Forms.Padding(3);
            this.tabEdit.Size = new System.Drawing.Size(615, 440);
            this.tabEdit.TabIndex = 0;
            this.tabEdit.Text = "Edit";
            this.tabEdit.UseVisualStyleBackColor = true;
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(20, 12);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(66, 14);
            this.lblTitle.TabIndex = 27;
            this.lblTitle.Text = "To be Set";
            // 
            // lblUpdateMessage
            // 
            this.lblUpdateMessage.AutoSize = true;
            this.lblUpdateMessage.Location = new System.Drawing.Point(20, 171);
            this.lblUpdateMessage.Name = "lblUpdateMessage";
            this.lblUpdateMessage.Size = new System.Drawing.Size(91, 13);
            this.lblUpdateMessage.TabIndex = 26;
            this.lblUpdateMessage.Text = "(update message)";
            // 
            // lblReqCount
            // 
            this.lblReqCount.AutoSize = true;
            this.lblReqCount.Location = new System.Drawing.Point(62, 187);
            this.lblReqCount.Name = "lblReqCount";
            this.lblReqCount.Size = new System.Drawing.Size(13, 13);
            this.lblReqCount.TabIndex = 25;
            this.lblReqCount.Text = "0";
            this.lblReqCount.Click += new System.EventHandler(this.lblReqCount_Click);
            // 
            // lblSpaceCount
            // 
            this.lblSpaceCount.AutoSize = true;
            this.lblSpaceCount.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSpaceCount.Location = new System.Drawing.Point(320, 94);
            this.lblSpaceCount.Name = "lblSpaceCount";
            this.lblSpaceCount.Size = new System.Drawing.Size(28, 13);
            this.lblSpaceCount.TabIndex = 24;
            this.lblSpaceCount.Text = "255";
            // 
            // lblSpace
            // 
            this.lblSpace.AutoSize = true;
            this.lblSpace.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSpace.Location = new System.Drawing.Point(207, 94);
            this.lblSpace.Name = "lblSpace";
            this.lblSpace.Size = new System.Drawing.Size(107, 13);
            this.lblSpace.TabIndex = 23;
            this.lblSpace.Text = "Remaining Space:";
            // 
            // lblPromptDoubleClick
            // 
            this.lblPromptDoubleClick.AutoSize = true;
            this.lblPromptDoubleClick.Location = new System.Drawing.Point(17, 392);
            this.lblPromptDoubleClick.Name = "lblPromptDoubleClick";
            this.lblPromptDoubleClick.Size = new System.Drawing.Size(112, 13);
            this.lblPromptDoubleClick.TabIndex = 22;
            this.lblPromptDoubleClick.Text = "<set in program code>";
            // 
            // btnPropertyDelete
            // 
            this.btnPropertyDelete.Location = new System.Drawing.Point(465, 405);
            this.btnPropertyDelete.Name = "btnPropertyDelete";
            this.btnPropertyDelete.Size = new System.Drawing.Size(75, 23);
            this.btnPropertyDelete.TabIndex = 21;
            this.btnPropertyDelete.Text = "Delete";
            this.btnPropertyDelete.UseVisualStyleBackColor = true;
            this.btnPropertyDelete.Click += new System.EventHandler(this.btnPropertyDelete_Click);
            // 
            // lblDocProperty
            // 
            this.lblDocProperty.AutoSize = true;
            this.lblDocProperty.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDocProperty.Location = new System.Drawing.Point(20, 200);
            this.lblDocProperty.Name = "lblDocProperty";
            this.lblDocProperty.Size = new System.Drawing.Size(361, 13);
            this.lblDocProperty.TabIndex = 20;
            this.lblDocProperty.Text = "Req\'d.                      Document Properties                                  " +
    "Value";
            // 
            // chkShowModules
            // 
            this.chkShowModules.AutoSize = true;
            this.chkShowModules.Location = new System.Drawing.Point(243, 411);
            this.chkShowModules.Name = "chkShowModules";
            this.chkShowModules.Size = new System.Drawing.Size(96, 17);
            this.chkShowModules.TabIndex = 18;
            this.chkShowModules.Text = "Show Modules";
            this.chkShowModules.UseVisualStyleBackColor = true;
            // 
            // chkShowAllProperties
            // 
            this.chkShowAllProperties.AutoSize = true;
            this.chkShowAllProperties.Location = new System.Drawing.Point(21, 411);
            this.chkShowAllProperties.Name = "chkShowAllProperties";
            this.chkShowAllProperties.Size = new System.Drawing.Size(117, 17);
            this.chkShowAllProperties.TabIndex = 17;
            this.chkShowAllProperties.Text = "Show All Properties";
            this.chkShowAllProperties.UseVisualStyleBackColor = true;
            // 
            // dgvDocProperty
            // 
            this.dgvDocProperty.AllowUserToAddRows = false;
            this.dgvDocProperty.BackgroundColor = System.Drawing.Color.White;
            this.dgvDocProperty.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgvDocProperty.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.dgvDocProperty.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDocProperty.ColumnHeadersVisible = false;
            this.dgvDocProperty.Location = new System.Drawing.Point(21, 228);
            this.dgvDocProperty.MultiSelect = false;
            this.dgvDocProperty.Name = "dgvDocProperty";
            this.dgvDocProperty.ReadOnly = true;
            this.dgvDocProperty.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.dgvDocProperty.RowHeadersVisible = false;
            this.dgvDocProperty.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvDocProperty.Size = new System.Drawing.Size(569, 157);
            this.dgvDocProperty.TabIndex = 16;
            this.dgvDocProperty.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvDocProperty_CellDoubleClick);
            // 
            // btnDescribe
            // 
            this.btnDescribe.Location = new System.Drawing.Point(332, 167);
            this.btnDescribe.Name = "btnDescribe";
            this.btnDescribe.Size = new System.Drawing.Size(75, 23);
            this.btnDescribe.TabIndex = 15;
            this.btnDescribe.Text = "Describe";
            this.btnDescribe.UseVisualStyleBackColor = true;
            this.btnDescribe.Click += new System.EventHandler(this.btnDescribe_Click);
            // 
            // btnPropertyUpdate
            // 
            this.btnPropertyUpdate.Location = new System.Drawing.Point(238, 167);
            this.btnPropertyUpdate.Name = "btnPropertyUpdate";
            this.btnPropertyUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnPropertyUpdate.TabIndex = 14;
            this.btnPropertyUpdate.Text = "Update";
            this.btnPropertyUpdate.UseVisualStyleBackColor = true;
            this.btnPropertyUpdate.Click += new System.EventHandler(this.btnPropertyUpdate_Click);
            // 
            // btnPropertyAdd
            // 
            this.btnPropertyAdd.Location = new System.Drawing.Point(145, 167);
            this.btnPropertyAdd.Name = "btnPropertyAdd";
            this.btnPropertyAdd.Size = new System.Drawing.Size(75, 23);
            this.btnPropertyAdd.TabIndex = 13;
            this.btnPropertyAdd.Text = "Add";
            this.btnPropertyAdd.UseVisualStyleBackColor = true;
            this.btnPropertyAdd.Click += new System.EventHandler(this.btnPropertyAdd_Click);
            // 
            // cmdPropertyValueSize
            // 
            this.cmdPropertyValueSize.Location = new System.Drawing.Point(563, 114);
            this.cmdPropertyValueSize.Name = "cmdPropertyValueSize";
            this.cmdPropertyValueSize.Size = new System.Drawing.Size(32, 23);
            this.cmdPropertyValueSize.TabIndex = 12;
            this.cmdPropertyValueSize.Text = "...";
            this.cmdPropertyValueSize.UseVisualStyleBackColor = true;
            this.cmdPropertyValueSize.Click += new System.EventHandler(this.cmdPropertyValueSize_Click);
            // 
            // cmdPropertyValueBracket
            // 
            this.cmdPropertyValueBracket.Location = new System.Drawing.Point(532, 114);
            this.cmdPropertyValueBracket.Name = "cmdPropertyValueBracket";
            this.cmdPropertyValueBracket.Size = new System.Drawing.Size(32, 23);
            this.cmdPropertyValueBracket.TabIndex = 11;
            this.cmdPropertyValueBracket.Text = "[ ]";
            this.cmdPropertyValueBracket.UseVisualStyleBackColor = true;
            this.cmdPropertyValueBracket.Click += new System.EventHandler(this.cmdPropertyValueBracket_Click);
            // 
            // ddlPropertyUOM
            // 
            this.ddlPropertyUOM.FormattingEnabled = true;
            this.ddlPropertyUOM.Location = new System.Drawing.Point(114, 143);
            this.ddlPropertyUOM.Name = "ddlPropertyUOM";
            this.ddlPropertyUOM.Size = new System.Drawing.Size(476, 21);
            this.ddlPropertyUOM.TabIndex = 10;
            // 
            // ddlPropertyValue
            // 
            this.ddlPropertyValue.FormattingEnabled = true;
            this.ddlPropertyValue.Location = new System.Drawing.Point(114, 114);
            this.ddlPropertyValue.Name = "ddlPropertyValue";
            this.ddlPropertyValue.Size = new System.Drawing.Size(476, 21);
            this.ddlPropertyValue.TabIndex = 9;
            this.ddlPropertyValue.SelectedIndexChanged += new System.EventHandler(this.ddlPropertyValue_SelectedIndexChanged);
            // 
            // txtPropertyValue
            // 
            this.txtPropertyValue.Location = new System.Drawing.Point(114, 116);
            this.txtPropertyValue.Name = "txtPropertyValue";
            this.txtPropertyValue.Size = new System.Drawing.Size(141, 20);
            this.txtPropertyValue.TabIndex = 8;
            this.txtPropertyValue.TextChanged += new System.EventHandler(this.txtPropertyValue_TextChanged);
            // 
            // ddlPropertyAll
            // 
            this.ddlPropertyAll.FormattingEnabled = true;
            this.ddlPropertyAll.Location = new System.Drawing.Point(114, 64);
            this.ddlPropertyAll.Name = "ddlPropertyAll";
            this.ddlPropertyAll.Size = new System.Drawing.Size(476, 21);
            this.ddlPropertyAll.TabIndex = 7;
            this.ddlPropertyAll.SelectedIndexChanged += new System.EventHandler(this.ddlPropertyAll_SelectedIndexChanged);
            // 
            // ddlPropertyCategory
            // 
            this.ddlPropertyCategory.FormattingEnabled = true;
            this.ddlPropertyCategory.Location = new System.Drawing.Point(114, 37);
            this.ddlPropertyCategory.Name = "ddlPropertyCategory";
            this.ddlPropertyCategory.Size = new System.Drawing.Size(476, 21);
            this.ddlPropertyCategory.TabIndex = 6;
            this.ddlPropertyCategory.SelectedIndexChanged += new System.EventHandler(this.ddlPropertyCategory_SelectedIndexChanged);
            // 
            // spnPropertyIndex
            // 
            this.spnPropertyIndex.Location = new System.Drawing.Point(114, 91);
            this.spnPropertyIndex.Name = "spnPropertyIndex";
            this.spnPropertyIndex.Size = new System.Drawing.Size(66, 20);
            this.spnPropertyIndex.TabIndex = 5;
            this.spnPropertyIndex.ValueChanged += new System.EventHandler(this.spnPropertyIndex_ValueChanged);
            // 
            // lblPropertyUOM
            // 
            this.lblPropertyUOM.AutoSize = true;
            this.lblPropertyUOM.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPropertyUOM.Location = new System.Drawing.Point(18, 151);
            this.lblPropertyUOM.Name = "lblPropertyUOM";
            this.lblPropertyUOM.Size = new System.Drawing.Size(83, 13);
            this.lblPropertyUOM.TabIndex = 4;
            this.lblPropertyUOM.Text = "Unit of Measure";
            // 
            // lblPropertyValue
            // 
            this.lblPropertyValue.AutoSize = true;
            this.lblPropertyValue.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPropertyValue.Location = new System.Drawing.Point(17, 122);
            this.lblPropertyValue.Name = "lblPropertyValue";
            this.lblPropertyValue.Size = new System.Drawing.Size(78, 13);
            this.lblPropertyValue.TabIndex = 3;
            this.lblPropertyValue.Text = "Property Value";
            // 
            // lblPropertyIndex
            // 
            this.lblPropertyIndex.AutoSize = true;
            this.lblPropertyIndex.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPropertyIndex.Location = new System.Drawing.Point(18, 97);
            this.lblPropertyIndex.Name = "lblPropertyIndex";
            this.lblPropertyIndex.Size = new System.Drawing.Size(80, 13);
            this.lblPropertyIndex.TabIndex = 2;
            this.lblPropertyIndex.Text = "Property Index";
            // 
            // lblPropertyAll
            // 
            this.lblPropertyAll.AutoSize = true;
            this.lblPropertyAll.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPropertyAll.Location = new System.Drawing.Point(17, 70);
            this.lblPropertyAll.Name = "lblPropertyAll";
            this.lblPropertyAll.Size = new System.Drawing.Size(79, 13);
            this.lblPropertyAll.TabIndex = 1;
            this.lblPropertyAll.Text = "Property Name";
            // 
            // lblPropertyCategory
            // 
            this.lblPropertyCategory.AutoSize = true;
            this.lblPropertyCategory.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPropertyCategory.Location = new System.Drawing.Point(17, 42);
            this.lblPropertyCategory.Name = "lblPropertyCategory";
            this.lblPropertyCategory.Size = new System.Drawing.Size(52, 13);
            this.lblPropertyCategory.TabIndex = 0;
            this.lblPropertyCategory.Text = "Category";
            // 
            // tabSource
            // 
            this.tabSource.Controls.Add(this.lblSourceHeader);
            this.tabSource.Controls.Add(this.lblSourcepty);
            this.tabSource.Controls.Add(this.btnSyncSysProperties);
            this.tabSource.Controls.Add(this.dgvPropertySource);
            this.tabSource.Location = new System.Drawing.Point(4, 22);
            this.tabSource.Name = "tabSource";
            this.tabSource.Padding = new System.Windows.Forms.Padding(3);
            this.tabSource.Size = new System.Drawing.Size(615, 440);
            this.tabSource.TabIndex = 1;
            this.tabSource.Text = "Source";
            this.tabSource.UseVisualStyleBackColor = true;
            // 
            // lblSourceHeader
            // 
            this.lblSourceHeader.AutoSize = true;
            this.lblSourceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSourceHeader.Location = new System.Drawing.Point(21, 27);
            this.lblSourceHeader.Name = "lblSourceHeader";
            this.lblSourceHeader.Size = new System.Drawing.Size(473, 16);
            this.lblSourceHeader.TabIndex = 6;
            this.lblSourceHeader.Text = "   Property Name           PropertyValue       Source Module        Location  ";
            // 
            // lblSourcepty
            // 
            this.lblSourcepty.AutoSize = true;
            this.lblSourcepty.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSourcepty.Location = new System.Drawing.Point(21, 3);
            this.lblSourcepty.Name = "lblSourcepty";
            this.lblSourcepty.Size = new System.Drawing.Size(115, 16);
            this.lblSourcepty.TabIndex = 5;
            this.lblSourcepty.Text = "Property Source";
            // 
            // btnSyncSysProperties
            // 
            this.btnSyncSysProperties.Location = new System.Drawing.Point(186, 341);
            this.btnSyncSysProperties.Name = "btnSyncSysProperties";
            this.btnSyncSysProperties.Size = new System.Drawing.Size(279, 32);
            this.btnSyncSysProperties.TabIndex = 1;
            this.btnSyncSysProperties.Text = "Synchronize System Properties";
            this.btnSyncSysProperties.UseVisualStyleBackColor = true;
            this.btnSyncSysProperties.Click += new System.EventHandler(this.btnSyncSysProperties_Click);
            // 
            // dgvPropertySource
            // 
            this.dgvPropertySource.AllowUserToAddRows = false;
            this.dgvPropertySource.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvPropertySource.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPropertySource.ColumnHeadersVisible = false;
            this.dgvPropertySource.Location = new System.Drawing.Point(6, 52);
            this.dgvPropertySource.Name = "dgvPropertySource";
            this.dgvPropertySource.Size = new System.Drawing.Size(585, 283);
            this.dgvPropertySource.TabIndex = 0;
            this.dgvPropertySource.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPropertySource_CellDoubleClick);
            // 
            // tabUnused
            // 
            this.tabUnused.Controls.Add(this.lblPropertyHeader);
            this.tabUnused.Controls.Add(this.lblUnusedpty);
            this.tabUnused.Controls.Add(this.btnUnusedPropertyDeleteAll);
            this.tabUnused.Controls.Add(this.btnUnusedPropertyDelete);
            this.tabUnused.Controls.Add(this.dgvDocPropertyUnused);
            this.tabUnused.Location = new System.Drawing.Point(4, 22);
            this.tabUnused.Name = "tabUnused";
            this.tabUnused.Padding = new System.Windows.Forms.Padding(3);
            this.tabUnused.Size = new System.Drawing.Size(615, 440);
            this.tabUnused.TabIndex = 2;
            this.tabUnused.Text = "Unused";
            this.tabUnused.UseVisualStyleBackColor = true;
            // 
            // lblPropertyHeader
            // 
            this.lblPropertyHeader.AutoSize = true;
            this.lblPropertyHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPropertyHeader.Location = new System.Drawing.Point(105, 34);
            this.lblPropertyHeader.Name = "lblPropertyHeader";
            this.lblPropertyHeader.Size = new System.Drawing.Size(339, 16);
            this.lblPropertyHeader.TabIndex = 4;
            this.lblPropertyHeader.Text = "     Property Name                              PropertyValue";
            // 
            // lblUnusedpty
            // 
            this.lblUnusedpty.AutoSize = true;
            this.lblUnusedpty.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUnusedpty.Location = new System.Drawing.Point(100, 15);
            this.lblUnusedpty.Name = "lblUnusedpty";
            this.lblUnusedpty.Size = new System.Drawing.Size(272, 16);
            this.lblUnusedpty.TabIndex = 3;
            this.lblUnusedpty.Text = "Properties Not Associated with a Module";
            // 
            // btnUnusedPropertyDeleteAll
            // 
            this.btnUnusedPropertyDeleteAll.Location = new System.Drawing.Point(337, 336);
            this.btnUnusedPropertyDeleteAll.Name = "btnUnusedPropertyDeleteAll";
            this.btnUnusedPropertyDeleteAll.Size = new System.Drawing.Size(107, 34);
            this.btnUnusedPropertyDeleteAll.TabIndex = 2;
            this.btnUnusedPropertyDeleteAll.Text = "Delete All";
            this.btnUnusedPropertyDeleteAll.UseVisualStyleBackColor = true;
            this.btnUnusedPropertyDeleteAll.Click += new System.EventHandler(this.btnUnusedPropertyDeleteAll_Click);
            // 
            // btnUnusedPropertyDelete
            // 
            this.btnUnusedPropertyDelete.Location = new System.Drawing.Point(186, 336);
            this.btnUnusedPropertyDelete.Name = "btnUnusedPropertyDelete";
            this.btnUnusedPropertyDelete.Size = new System.Drawing.Size(111, 34);
            this.btnUnusedPropertyDelete.TabIndex = 1;
            this.btnUnusedPropertyDelete.Text = "Delete Selected";
            this.btnUnusedPropertyDelete.UseVisualStyleBackColor = true;
            this.btnUnusedPropertyDelete.Click += new System.EventHandler(this.btnUnusedPropertyDelete_Click);
            // 
            // dgvDocPropertyUnused
            // 
            this.dgvDocPropertyUnused.AllowUserToAddRows = false;
            this.dgvDocPropertyUnused.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDocPropertyUnused.ColumnHeadersVisible = false;
            this.dgvDocPropertyUnused.Location = new System.Drawing.Point(100, 67);
            this.dgvDocPropertyUnused.Name = "dgvDocPropertyUnused";
            this.dgvDocPropertyUnused.Size = new System.Drawing.Size(457, 260);
            this.dgvDocPropertyUnused.TabIndex = 0;
            // 
            // tabMessages
            // 
            this.tabMessages.Controls.Add(this.lblMessages);
            this.tabMessages.Controls.Add(this.lstMessages);
            this.tabMessages.Location = new System.Drawing.Point(4, 22);
            this.tabMessages.Name = "tabMessages";
            this.tabMessages.Padding = new System.Windows.Forms.Padding(3);
            this.tabMessages.Size = new System.Drawing.Size(615, 440);
            this.tabMessages.TabIndex = 3;
            this.tabMessages.Text = "Messages";
            this.tabMessages.UseVisualStyleBackColor = true;
            // 
            // lblMessages
            // 
            this.lblMessages.AutoSize = true;
            this.lblMessages.Location = new System.Drawing.Point(7, 11);
            this.lblMessages.Name = "lblMessages";
            this.lblMessages.Size = new System.Drawing.Size(65, 13);
            this.lblMessages.TabIndex = 1;
            this.lblMessages.Text = "lblMessages";
            // 
            // lstMessages
            // 
            this.lstMessages.FormattingEnabled = true;
            this.lstMessages.Location = new System.Drawing.Point(3, 32);
            this.lstMessages.Name = "lstMessages";
            this.lstMessages.Size = new System.Drawing.Size(588, 355);
            this.lstMessages.TabIndex = 0;
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(464, 470);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(96, 27);
            this.btnNext.TabIndex = 1;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnFinish
            // 
            this.btnFinish.Location = new System.Drawing.Point(566, 470);
            this.btnFinish.Name = "btnFinish";
            this.btnFinish.Size = new System.Drawing.Size(96, 27);
            this.btnFinish.TabIndex = 2;
            this.btnFinish.Text = "Finish";
            this.btnFinish.UseVisualStyleBackColor = true;
            this.btnFinish.Click += new System.EventHandler(this.btnFinish_Click);
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(255, 470);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(97, 27);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnBack
            // 
            this.btnBack.Location = new System.Drawing.Point(358, 470);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(96, 27);
            this.btnBack.TabIndex = 4;
            this.btnBack.Text = "Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.shp2);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.shp3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label0);
            this.panel1.Controls.Add(this.shp1);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.shp0);
            this.panel1.Controls.Add(this.mapStatusStart);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Location = new System.Drawing.Point(12, 32);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(157, 226);
            this.panel1.TabIndex = 31;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Black;
            this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(54, 194);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(34, 13);
            this.label5.TabIndex = 35;
            this.label5.Tag = "1";
            this.label5.Text = "Finish";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // shp2
            // 
            this.shp2.BackColor = System.Drawing.Color.Silver;
            this.shp2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.shp2.Location = new System.Drawing.Point(40, 109);
            this.shp2.Name = "shp2";
            this.shp2.Size = new System.Drawing.Size(16, 15);
            this.shp2.TabIndex = 34;
            this.shp2.Tag = "2";
            this.shp2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Black;
            this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(61, 165);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(68, 13);
            this.label4.TabIndex = 33;
            this.label4.Tag = "3";
            this.label4.Text = "Messages(n)";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // shp3
            // 
            this.shp3.BackColor = System.Drawing.Color.Silver;
            this.shp3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.shp3.Location = new System.Drawing.Point(40, 163);
            this.shp3.Name = "shp3";
            this.shp3.Size = new System.Drawing.Size(15, 15);
            this.shp3.TabIndex = 32;
            this.shp3.Tag = "3";
            this.shp3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Black;
            this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(61, 109);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 26);
            this.label2.TabIndex = 31;
            this.label2.Tag = "2";
            this.label2.Text = "Un-Associated\r\nProperties";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Black;
            this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(58, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 13);
            this.label1.TabIndex = 30;
            this.label1.Tag = "1";
            this.label1.Text = "Property Source";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label0
            // 
            this.label0.AutoSize = true;
            this.label0.BackColor = System.Drawing.Color.Black;
            this.label0.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label0.ForeColor = System.Drawing.Color.White;
            this.label0.Location = new System.Drawing.Point(61, 15);
            this.label0.Name = "label0";
            this.label0.Size = new System.Drawing.Size(77, 13);
            this.label0.TabIndex = 29;
            this.label0.Tag = "0";
            this.label0.Text = "Edit Properties";
            this.label0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // shp1
            // 
            this.shp1.BackColor = System.Drawing.Color.Silver;
            this.shp1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.shp1.Location = new System.Drawing.Point(40, 52);
            this.shp1.Name = "shp1";
            this.shp1.Size = new System.Drawing.Size(15, 15);
            this.shp1.TabIndex = 28;
            this.shp1.Tag = "1";
            this.shp1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Red;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(3, 192);
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
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label7.Location = new System.Drawing.Point(11, 22);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(37, 185);
            this.label7.TabIndex = 23;
            this.label7.Tag = "Exclude";
            // 
            // FormEditProperty
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(800, 520);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.tabEditProperty);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnFinish);
            this.Controls.Add(this.btnNext);
            this.Name = "FormEditProperty";
            this.Text = "Edit Properties Wizard";
            this.Load += new System.EventHandler(this.FormEditProperty_Load);
            this.Shown += new System.EventHandler(this.FormEditProperty_Shown);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormEditProperty_FormClosed);
            this.tabEditProperty.ResumeLayout(false);
            this.tabEdit.ResumeLayout(false);
            this.tabEdit.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDocProperty)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.spnPropertyIndex)).EndInit();
            this.tabSource.ResumeLayout(false);
            this.tabSource.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPropertySource)).EndInit();
            this.tabUnused.ResumeLayout(false);
            this.tabUnused.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDocPropertyUnused)).EndInit();
            this.tabMessages.ResumeLayout(false);
            this.tabMessages.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }


        #endregion

        private System.Windows.Forms.TabControl tabEditProperty;
        private System.Windows.Forms.TabPage tabEdit;
        private System.Windows.Forms.TabPage tabSource;
        private System.Windows.Forms.TabPage tabUnused;
        private System.Windows.Forms.TabPage tabMessages;
        private System.Windows.Forms.Label lblPropertyUOM;
        private System.Windows.Forms.Label lblPropertyValue;
        private System.Windows.Forms.Label lblPropertyIndex;
        private System.Windows.Forms.Label lblPropertyAll;
        private System.Windows.Forms.Label lblPropertyCategory;
        private System.Windows.Forms.CheckBox chkShowModules;
        private System.Windows.Forms.CheckBox chkShowAllProperties;
        private System.Windows.Forms.DataGridView dgvDocProperty;
        private System.Windows.Forms.Button btnDescribe;
        private System.Windows.Forms.Button btnPropertyUpdate;
        private System.Windows.Forms.Button btnPropertyAdd;
        private System.Windows.Forms.Button cmdPropertyValueSize;
        private System.Windows.Forms.Button cmdPropertyValueBracket;
        private System.Windows.Forms.ComboBox ddlPropertyUOM;
        private System.Windows.Forms.ComboBox ddlPropertyValue;
        private System.Windows.Forms.TextBox txtPropertyValue;
        private System.Windows.Forms.ComboBox ddlPropertyAll;
        private System.Windows.Forms.ComboBox ddlPropertyCategory;
        private System.Windows.Forms.NumericUpDown spnPropertyIndex;
        private System.Windows.Forms.Label lblDocProperty;
        private System.Windows.Forms.Button btnPropertyDelete;
        private System.Windows.Forms.DataGridView dgvPropertySource;
        private System.Windows.Forms.Button btnSyncSysProperties;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Label lblPropertyHeader;
        private System.Windows.Forms.Label lblUnusedpty;
        private System.Windows.Forms.Button btnUnusedPropertyDeleteAll;
        private System.Windows.Forms.Button btnUnusedPropertyDelete;
        private System.Windows.Forms.DataGridView dgvDocPropertyUnused;
        private System.Windows.Forms.Label lblPromptDoubleClick;
        private System.Windows.Forms.ListBox lstMessages;
        private System.Windows.Forms.Label lblMessages;
        private System.Windows.Forms.Button btnFinish;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Label lblSpaceCount;
        private System.Windows.Forms.Label lblSpace;
        private System.Windows.Forms.Label lblUpdateMessage;
        private System.Windows.Forms.Label lblReqCount;
        private System.Windows.Forms.Label lblSourceHeader;
        private System.Windows.Forms.Label lblSourcepty;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label0;
        private System.Windows.Forms.Label shp1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label shp0;
        private System.Windows.Forms.Label mapStatusStart;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label shp2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label shp3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblTitle;
        private BindingSource bindingSourceData = new BindingSource();//100925
    }
}