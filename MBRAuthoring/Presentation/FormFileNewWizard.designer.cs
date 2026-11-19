
namespace MBRAuthoring.Presentation
{
    partial class FormFileNewWizard
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
            this.grpFileOptions = new System.Windows.Forms.GroupBox();
            this.optFileCopyDiff = new System.Windows.Forms.RadioButton();
            this.optFileCopyVers = new System.Windows.Forms.RadioButton();
            this.optFileCreate = new System.Windows.Forms.RadioButton();
            this.pnlCmdButtons = new System.Windows.Forms.Panel();
            this.cmdNext = new System.Windows.Forms.Button();
            this.cmdBack = new System.Windows.Forms.Button();
            this.cmdFinish = new System.Windows.Forms.Button();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.mpgWizardPage = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.cboDocumentSupercedes = new System.Windows.Forms.ComboBox();
            this.lblDocumentSupercedes = new System.Windows.Forms.Label();
            this.cboDocBaseNameCreate = new System.Windows.Forms.ComboBox();
            this.lblFileNameMessageCreate = new System.Windows.Forms.Label();
            this.lblFileNameCreate = new System.Windows.Forms.Label();
            this.txtFreeTextCreate = new System.Windows.Forms.TextBox();
            this.txtStepCode = new System.Windows.Forms.TextBox();
            this.optProductOrderByTitleCreate = new System.Windows.Forms.RadioButton();
            this.optProductOrderByNumberCreate = new System.Windows.Forms.RadioButton();
            this.cboDocTypeCreate = new System.Windows.Forms.ComboBox();
            this.cboWorkCenter = new System.Windows.Forms.ComboBox();
            this.cboProductCreate = new System.Windows.Forms.ComboBox();
            this.cboOpUnitCreate = new System.Windows.Forms.ComboBox();
            this.lblFileNameNewMessageLabel = new System.Windows.Forms.Label();
            this.lblRecordTypeNew = new System.Windows.Forms.Label();
            this.lblWCEquip = new System.Windows.Forms.Label();
            this.lblStepCode = new System.Windows.Forms.Label();
            this.lblOrderByCreate = new System.Windows.Forms.Label();
            this.lblProduct = new System.Windows.Forms.Label();
            this.lblFreeTextCreate = new System.Windows.Forms.Label();
            this.lblFileNameLabel = new System.Windows.Forms.Label();
            this.lblOpUnitNew = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.cboProductCopy = new System.Windows.Forms.ComboBox();
            this.cboDocBaseNameCopy = new System.Windows.Forms.ComboBox();
            this.lblEDMSPassword = new System.Windows.Forms.Label();
            this.txtEDMSUserID = new System.Windows.Forms.TextBox();
            this.txtEDMSPassword = new System.Windows.Forms.TextBox();
            this.lblEDMSUserID = new System.Windows.Forms.Label();
            this.lblFileNameMessageCopy = new System.Windows.Forms.Label();
            this.lblFileNameMessageCopyLabel = new System.Windows.Forms.Label();
            this.lblFileNameCopy = new System.Windows.Forms.Label();
            this.lblFileNameCopyLabel = new System.Windows.Forms.Label();
            this.txtVariantTextCopy = new System.Windows.Forms.TextBox();
            this.lblVariantText = new System.Windows.Forms.Label();
            this.optProductOrderByTitleCopy = new System.Windows.Forms.RadioButton();
            this.optProductOrderByNumberCopy = new System.Windows.Forms.RadioButton();
            this.lblOrderByCopy = new System.Windows.Forms.Label();
            this.cboDocNameCopy = new System.Windows.Forms.ComboBox();
            this.lblDocName = new System.Windows.Forms.Label();
            this.cboDocLocation = new System.Windows.Forms.ComboBox();
            this.lblDocLocation = new System.Windows.Forms.Label();
            this.cboDocTypeCopy = new System.Windows.Forms.ComboBox();
            this.lblRecordTypeCopy = new System.Windows.Forms.Label();
            this.cboOpUnitCopy = new System.Windows.Forms.ComboBox();
            this.lblOpUnitCopy = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.mapStatusStart = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.grpFileOptions.SuspendLayout();
            this.pnlCmdButtons.SuspendLayout();
            this.mpgWizardPage.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpFileOptions
            // 
            this.grpFileOptions.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.grpFileOptions.Controls.Add(this.optFileCopyDiff);
            this.grpFileOptions.Controls.Add(this.optFileCopyVers);
            this.grpFileOptions.Controls.Add(this.optFileCreate);
            this.grpFileOptions.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpFileOptions.Location = new System.Drawing.Point(145, -14);
            this.grpFileOptions.Name = "grpFileOptions";
            this.grpFileOptions.Size = new System.Drawing.Size(574, 70);
            this.grpFileOptions.TabIndex = 0;
            this.grpFileOptions.TabStop = false;
            // 
            // optFileCopyDiff
            // 
            this.optFileCopyDiff.Dock = System.Windows.Forms.DockStyle.Left;
            this.optFileCopyDiff.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.optFileCopyDiff.Location = new System.Drawing.Point(326, 17);
            this.optFileCopyDiff.Name = "optFileCopyDiff";
            this.optFileCopyDiff.Size = new System.Drawing.Size(169, 50);
            this.optFileCopyDiff.TabIndex = 2;
            this.optFileCopyDiff.Text = "Copy Existing Document to a Different Name";
            this.optFileCopyDiff.UseVisualStyleBackColor = true;
            this.optFileCopyDiff.Click += new System.EventHandler(this.optFileOptions_Click);
            // 
            // optFileCopyVers
            // 
            this.optFileCopyVers.Dock = System.Windows.Forms.DockStyle.Left;
            this.optFileCopyVers.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.optFileCopyVers.Location = new System.Drawing.Point(156, 17);
            this.optFileCopyVers.Name = "optFileCopyVers";
            this.optFileCopyVers.Size = new System.Drawing.Size(170, 50);
            this.optFileCopyVers.TabIndex = 1;
            this.optFileCopyVers.Text = "Copy Existing Document to a New Version";
            this.optFileCopyVers.UseVisualStyleBackColor = true;
            this.optFileCopyVers.Click += new System.EventHandler(this.optFileOptions_Click);
            // 
            // optFileCreate
            // 
            this.optFileCreate.Dock = System.Windows.Forms.DockStyle.Left;
            this.optFileCreate.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.optFileCreate.Location = new System.Drawing.Point(3, 17);
            this.optFileCreate.Name = "optFileCreate";
            this.optFileCreate.Size = new System.Drawing.Size(153, 50);
            this.optFileCreate.TabIndex = 0;
            this.optFileCreate.Text = "Create a New Document";
            this.optFileCreate.UseVisualStyleBackColor = true;
            this.optFileCreate.Click += new System.EventHandler(this.optFileOptions_Click);
            // 
            // pnlCmdButtons
            // 
            this.pnlCmdButtons.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.pnlCmdButtons.BackColor = System.Drawing.Color.Black;
            this.pnlCmdButtons.Controls.Add(this.cmdNext);
            this.pnlCmdButtons.Controls.Add(this.cmdBack);
            this.pnlCmdButtons.Controls.Add(this.cmdFinish);
            this.pnlCmdButtons.Controls.Add(this.cmdCancel);
            this.pnlCmdButtons.Location = new System.Drawing.Point(145, 405);
            this.pnlCmdButtons.Name = "pnlCmdButtons";
            this.pnlCmdButtons.Size = new System.Drawing.Size(576, 60);
            this.pnlCmdButtons.TabIndex = 9;
            // 
            // cmdNext
            // 
            this.cmdNext.BackColor = System.Drawing.SystemColors.Control;
            this.cmdNext.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdNext.Location = new System.Drawing.Point(334, 14);
            this.cmdNext.Name = "cmdNext";
            this.cmdNext.Size = new System.Drawing.Size(103, 34);
            this.cmdNext.TabIndex = 5;
            this.cmdNext.Text = "Next";
            this.cmdNext.UseVisualStyleBackColor = false;
            // 
            // cmdBack
            // 
            this.cmdBack.BackColor = System.Drawing.SystemColors.Control;
            this.cmdBack.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdBack.Location = new System.Drawing.Point(225, 14);
            this.cmdBack.Name = "cmdBack";
            this.cmdBack.Size = new System.Drawing.Size(103, 34);
            this.cmdBack.TabIndex = 4;
            this.cmdBack.Text = "Back";
            this.cmdBack.UseVisualStyleBackColor = false;
            // 
            // cmdFinish
            // 
            this.cmdFinish.BackColor = System.Drawing.SystemColors.Control;
            this.cmdFinish.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdFinish.Location = new System.Drawing.Point(460, 14);
            this.cmdFinish.Name = "cmdFinish";
            this.cmdFinish.Size = new System.Drawing.Size(103, 34);
            this.cmdFinish.TabIndex = 3;
            this.cmdFinish.Text = "Finish";
            this.cmdFinish.UseVisualStyleBackColor = false;
            this.cmdFinish.Click += new System.EventHandler(this.cmdFinish_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.BackColor = System.Drawing.SystemColors.Control;
            this.cmdCancel.Location = new System.Drawing.Point(103, 14);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(103, 34);
            this.cmdCancel.TabIndex = 1;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = false;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // mpgWizardPage
            // 
            this.mpgWizardPage.Controls.Add(this.tabPage1);
            this.mpgWizardPage.Controls.Add(this.tabPage2);
            this.mpgWizardPage.Location = new System.Drawing.Point(145, 62);
            this.mpgWizardPage.Name = "mpgWizardPage";
            this.mpgWizardPage.SelectedIndex = 0;
            this.mpgWizardPage.Size = new System.Drawing.Size(574, 326);
            this.mpgWizardPage.TabIndex = 2;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.tabPage1.Controls.Add(this.cboDocumentSupercedes);
            this.tabPage1.Controls.Add(this.lblDocumentSupercedes);
            this.tabPage1.Controls.Add(this.cboDocBaseNameCreate);
            this.tabPage1.Controls.Add(this.lblFileNameMessageCreate);
            this.tabPage1.Controls.Add(this.lblFileNameCreate);
            this.tabPage1.Controls.Add(this.txtFreeTextCreate);
            this.tabPage1.Controls.Add(this.txtStepCode);
            this.tabPage1.Controls.Add(this.optProductOrderByTitleCreate);
            this.tabPage1.Controls.Add(this.optProductOrderByNumberCreate);
            this.tabPage1.Controls.Add(this.cboDocTypeCreate);
            this.tabPage1.Controls.Add(this.cboWorkCenter);
            this.tabPage1.Controls.Add(this.cboProductCreate);
            this.tabPage1.Controls.Add(this.cboOpUnitCreate);
            this.tabPage1.Controls.Add(this.lblFileNameNewMessageLabel);
            this.tabPage1.Controls.Add(this.lblRecordTypeNew);
            this.tabPage1.Controls.Add(this.lblWCEquip);
            this.tabPage1.Controls.Add(this.lblStepCode);
            this.tabPage1.Controls.Add(this.lblOrderByCreate);
            this.tabPage1.Controls.Add(this.lblProduct);
            this.tabPage1.Controls.Add(this.lblFreeTextCreate);
            this.tabPage1.Controls.Add(this.lblFileNameLabel);
            this.tabPage1.Controls.Add(this.lblOpUnitNew);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(566, 300);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "0=Create";
            // 
            // cboDocumentSupercedes
            // 
            this.cboDocumentSupercedes.BackColor = System.Drawing.Color.White;
            this.cboDocumentSupercedes.FormattingEnabled = true;
            this.cboDocumentSupercedes.Location = new System.Drawing.Point(161, 275);
            this.cboDocumentSupercedes.Name = "cboDocumentSupercedes";
            this.cboDocumentSupercedes.Size = new System.Drawing.Size(376, 21);
            this.cboDocumentSupercedes.TabIndex = 21;
            // 
            // lblDocumentSupercedes
            // 
            this.lblDocumentSupercedes.Location = new System.Drawing.Point(30, 272);
            this.lblDocumentSupercedes.Name = "lblDocumentSupercedes";
            this.lblDocumentSupercedes.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblDocumentSupercedes.Size = new System.Drawing.Size(114, 25);
            this.lblDocumentSupercedes.TabIndex = 20;
            this.lblDocumentSupercedes.Text = "Supercedes:";
            this.lblDocumentSupercedes.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cboDocBaseNameCreate
            // 
            this.cboDocBaseNameCreate.BackColor = System.Drawing.Color.White;
            this.cboDocBaseNameCreate.FormattingEnabled = true;
            this.cboDocBaseNameCreate.Location = new System.Drawing.Point(161, 168);
            this.cboDocBaseNameCreate.Name = "cboDocBaseNameCreate";
            this.cboDocBaseNameCreate.Size = new System.Drawing.Size(376, 21);
            this.cboDocBaseNameCreate.TabIndex = 19;
            this.cboDocBaseNameCreate.SelectedIndexChanged += new System.EventHandler(this.cboDocBaseNameCreate_SelectedIndexChanged);
            this.cboDocBaseNameCreate.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cboDocBaseNameCreate_KeyDown);
            // 
            // lblFileNameMessageCreate
            // 
            this.lblFileNameMessageCreate.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileNameMessageCreate.ForeColor = System.Drawing.Color.MediumBlue;
            this.lblFileNameMessageCreate.Location = new System.Drawing.Point(158, 220);
            this.lblFileNameMessageCreate.Name = "lblFileNameMessageCreate";
            this.lblFileNameMessageCreate.Size = new System.Drawing.Size(402, 34);
            this.lblFileNameMessageCreate.TabIndex = 18;
            // 
            // lblFileNameCreate
            // 
            this.lblFileNameCreate.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileNameCreate.ForeColor = System.Drawing.Color.MediumBlue;
            this.lblFileNameCreate.Location = new System.Drawing.Point(158, 195);
            this.lblFileNameCreate.Name = "lblFileNameCreate";
            this.lblFileNameCreate.Size = new System.Drawing.Size(402, 19);
            this.lblFileNameCreate.TabIndex = 17;
            // 
            // txtFreeTextCreate
            // 
            this.txtFreeTextCreate.Location = new System.Drawing.Point(161, 167);
            this.txtFreeTextCreate.Name = "txtFreeTextCreate";
            this.txtFreeTextCreate.Size = new System.Drawing.Size(376, 21);
            this.txtFreeTextCreate.TabIndex = 16;
            this.txtFreeTextCreate.Visible = false;
            this.txtFreeTextCreate.TextChanged += new System.EventHandler(this.txtFreeTextCreate_TextChanged);
            // 
            // txtStepCode
            // 
            this.txtStepCode.BackColor = System.Drawing.Color.White;
            this.txtStepCode.Location = new System.Drawing.Point(161, 142);
            this.txtStepCode.Name = "txtStepCode";
            this.txtStepCode.Size = new System.Drawing.Size(121, 21);
            this.txtStepCode.TabIndex = 15;
            // 
            // optProductOrderByTitleCreate
            // 
            this.optProductOrderByTitleCreate.AutoSize = true;
            this.optProductOrderByTitleCreate.Location = new System.Drawing.Point(268, 91);
            this.optProductOrderByTitleCreate.Name = "optProductOrderByTitleCreate";
            this.optProductOrderByTitleCreate.Size = new System.Drawing.Size(45, 17);
            this.optProductOrderByTitleCreate.TabIndex = 14;
            this.optProductOrderByTitleCreate.Text = "Title";
            this.optProductOrderByTitleCreate.UseVisualStyleBackColor = true;
            this.optProductOrderByTitleCreate.Click += new System.EventHandler(this.optProductOrderByCreate_Click);
            // 
            // optProductOrderByNumberCreate
            // 
            this.optProductOrderByNumberCreate.AutoSize = true;
            this.optProductOrderByNumberCreate.Location = new System.Drawing.Point(161, 91);
            this.optProductOrderByNumberCreate.Name = "optProductOrderByNumberCreate";
            this.optProductOrderByNumberCreate.Size = new System.Drawing.Size(62, 17);
            this.optProductOrderByNumberCreate.TabIndex = 13;
            this.optProductOrderByNumberCreate.Text = "Number";
            this.optProductOrderByNumberCreate.UseVisualStyleBackColor = true;
            this.optProductOrderByNumberCreate.Click += new System.EventHandler(this.optProductOrderByCreate_Click);
            // 
            // cboDocTypeCreate
            // 
            this.cboDocTypeCreate.FormattingEnabled = true;
            this.cboDocTypeCreate.Location = new System.Drawing.Point(161, 42);
            this.cboDocTypeCreate.Name = "cboDocTypeCreate";
            this.cboDocTypeCreate.Size = new System.Drawing.Size(240, 21);
            this.cboDocTypeCreate.TabIndex = 12;
            this.cboDocTypeCreate.SelectedIndexChanged += new System.EventHandler(this.cboDocTypeCreate_SelectedIndexChanged);
            // 
            // cboWorkCenter
            // 
            this.cboWorkCenter.FormattingEnabled = true;
            this.cboWorkCenter.Location = new System.Drawing.Point(161, 68);
            this.cboWorkCenter.Name = "cboWorkCenter";
            this.cboWorkCenter.Size = new System.Drawing.Size(240, 21);
            this.cboWorkCenter.TabIndex = 11;
            // 
            // cboProductCreate
            // 
            this.cboProductCreate.BackColor = System.Drawing.Color.White;
            this.cboProductCreate.DropDownHeight = 70;
            this.cboProductCreate.FormattingEnabled = true;
            this.cboProductCreate.IntegralHeight = false;
            this.cboProductCreate.Location = new System.Drawing.Point(161, 114);
            this.cboProductCreate.Name = "cboProductCreate";
            this.cboProductCreate.Size = new System.Drawing.Size(370, 21);
            this.cboProductCreate.TabIndex = 10;
            this.cboProductCreate.SelectedIndexChanged += new System.EventHandler(this.cboProductCreate_SelectedIndexChanged);
            // 
            // cboOpUnitCreate
            // 
            this.cboOpUnitCreate.FormattingEnabled = true;
            this.cboOpUnitCreate.Location = new System.Drawing.Point(161, 14);
            this.cboOpUnitCreate.Name = "cboOpUnitCreate";
            this.cboOpUnitCreate.Size = new System.Drawing.Size(121, 21);
            this.cboOpUnitCreate.TabIndex = 9;
            this.cboOpUnitCreate.SelectedIndexChanged += new System.EventHandler(this.cboOpUnitCreate_SelectedIndexChanged);
            // 
            // lblFileNameNewMessageLabel
            // 
            this.lblFileNameNewMessageLabel.Location = new System.Drawing.Point(30, 214);
            this.lblFileNameNewMessageLabel.Name = "lblFileNameNewMessageLabel";
            this.lblFileNameNewMessageLabel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblFileNameNewMessageLabel.Size = new System.Drawing.Size(114, 25);
            this.lblFileNameNewMessageLabel.TabIndex = 8;
            this.lblFileNameNewMessageLabel.Text = "File Name Status:";
            this.lblFileNameNewMessageLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblRecordTypeNew
            // 
            this.lblRecordTypeNew.Location = new System.Drawing.Point(30, 39);
            this.lblRecordTypeNew.Name = "lblRecordTypeNew";
            this.lblRecordTypeNew.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblRecordTypeNew.Size = new System.Drawing.Size(114, 25);
            this.lblRecordTypeNew.TabIndex = 7;
            this.lblRecordTypeNew.Text = "Document Type:";
            this.lblRecordTypeNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblWCEquip
            // 
            this.lblWCEquip.Location = new System.Drawing.Point(30, 64);
            this.lblWCEquip.Name = "lblWCEquip";
            this.lblWCEquip.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblWCEquip.Size = new System.Drawing.Size(114, 25);
            this.lblWCEquip.TabIndex = 6;
            this.lblWCEquip.Text = "Primary Work Center:";
            this.lblWCEquip.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblStepCode
            // 
            this.lblStepCode.Location = new System.Drawing.Point(30, 139);
            this.lblStepCode.Name = "lblStepCode";
            this.lblStepCode.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblStepCode.Size = new System.Drawing.Size(114, 25);
            this.lblStepCode.TabIndex = 5;
            this.lblStepCode.Text = "Step Code:";
            this.lblStepCode.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblOrderByCreate
            // 
            this.lblOrderByCreate.Location = new System.Drawing.Point(30, 89);
            this.lblOrderByCreate.Name = "lblOrderByCreate";
            this.lblOrderByCreate.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblOrderByCreate.Size = new System.Drawing.Size(114, 25);
            this.lblOrderByCreate.TabIndex = 4;
            this.lblOrderByCreate.Text = "Order by:";
            this.lblOrderByCreate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblProduct
            // 
            this.lblProduct.Location = new System.Drawing.Point(30, 114);
            this.lblProduct.Name = "lblProduct";
            this.lblProduct.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblProduct.Size = new System.Drawing.Size(114, 25);
            this.lblProduct.TabIndex = 3;
            this.lblProduct.Text = "Material:";
            this.lblProduct.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblFreeTextCreate
            // 
            this.lblFreeTextCreate.Location = new System.Drawing.Point(30, 164);
            this.lblFreeTextCreate.Name = "lblFreeTextCreate";
            this.lblFreeTextCreate.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblFreeTextCreate.Size = new System.Drawing.Size(114, 25);
            this.lblFreeTextCreate.TabIndex = 2;
            this.lblFreeTextCreate.Text = "Free text:";
            this.lblFreeTextCreate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblFileNameLabel
            // 
            this.lblFileNameLabel.Location = new System.Drawing.Point(30, 189);
            this.lblFileNameLabel.Name = "lblFileNameLabel";
            this.lblFileNameLabel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblFileNameLabel.Size = new System.Drawing.Size(114, 25);
            this.lblFileNameLabel.TabIndex = 1;
            this.lblFileNameLabel.Text = "File Name:";
            this.lblFileNameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblOpUnitNew
            // 
            this.lblOpUnitNew.Location = new System.Drawing.Point(30, 14);
            this.lblOpUnitNew.Name = "lblOpUnitNew";
            this.lblOpUnitNew.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblOpUnitNew.Size = new System.Drawing.Size(114, 25);
            this.lblOpUnitNew.TabIndex = 0;
            this.lblOpUnitNew.Text = "Operating Unit:";
            this.lblOpUnitNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.tabPage2.Controls.Add(this.cboProductCopy);
            this.tabPage2.Controls.Add(this.cboDocBaseNameCopy);
            this.tabPage2.Controls.Add(this.lblEDMSPassword);
            this.tabPage2.Controls.Add(this.txtEDMSUserID);
            this.tabPage2.Controls.Add(this.txtEDMSPassword);
            this.tabPage2.Controls.Add(this.lblEDMSUserID);
            this.tabPage2.Controls.Add(this.lblFileNameMessageCopy);
            this.tabPage2.Controls.Add(this.lblFileNameMessageCopyLabel);
            this.tabPage2.Controls.Add(this.lblFileNameCopy);
            this.tabPage2.Controls.Add(this.lblFileNameCopyLabel);
            this.tabPage2.Controls.Add(this.txtVariantTextCopy);
            this.tabPage2.Controls.Add(this.lblVariantText);
            this.tabPage2.Controls.Add(this.optProductOrderByTitleCopy);
            this.tabPage2.Controls.Add(this.optProductOrderByNumberCopy);
            this.tabPage2.Controls.Add(this.lblOrderByCopy);
            this.tabPage2.Controls.Add(this.cboDocNameCopy);
            this.tabPage2.Controls.Add(this.lblDocName);
            this.tabPage2.Controls.Add(this.cboDocLocation);
            this.tabPage2.Controls.Add(this.lblDocLocation);
            this.tabPage2.Controls.Add(this.cboDocTypeCopy);
            this.tabPage2.Controls.Add(this.lblRecordTypeCopy);
            this.tabPage2.Controls.Add(this.cboOpUnitCopy);
            this.tabPage2.Controls.Add(this.lblOpUnitCopy);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(566, 300);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "1=Copy";
            // 
            // cboProductCopy
            // 
            this.cboProductCopy.DropDownHeight = 70;
            this.cboProductCopy.FormattingEnabled = true;
            this.cboProductCopy.IntegralHeight = false;
            this.cboProductCopy.Location = new System.Drawing.Point(161, 141);
            this.cboProductCopy.Name = "cboProductCopy";
            this.cboProductCopy.Size = new System.Drawing.Size(376, 21);
            this.cboProductCopy.TabIndex = 32;
            this.cboProductCopy.SelectedIndexChanged += new System.EventHandler(this.cboProductCopy_SelectedIndexChanged);
            // 
            // cboDocBaseNameCopy
            // 
            this.cboDocBaseNameCopy.FormattingEnabled = true;
            this.cboDocBaseNameCopy.Location = new System.Drawing.Point(161, 141);
            this.cboDocBaseNameCopy.Name = "cboDocBaseNameCopy";
            this.cboDocBaseNameCopy.Size = new System.Drawing.Size(376, 21);
            this.cboDocBaseNameCopy.TabIndex = 31;
            this.cboDocBaseNameCopy.SelectedIndexChanged += new System.EventHandler(this.cboDocBaseNameCopy_SelectedIndexChanged);
            // 
            // lblEDMSPassword
            // 
            this.lblEDMSPassword.Location = new System.Drawing.Point(30, 258);
            this.lblEDMSPassword.Name = "lblEDMSPassword";
            this.lblEDMSPassword.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblEDMSPassword.Size = new System.Drawing.Size(114, 25);
            this.lblEDMSPassword.TabIndex = 30;
            this.lblEDMSPassword.Text = "Password:";
            this.lblEDMSPassword.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtEDMSUserID
            // 
            this.txtEDMSUserID.BackColor = System.Drawing.Color.White;
            this.txtEDMSUserID.Enabled = false;
            this.txtEDMSUserID.Location = new System.Drawing.Point(161, 232);
            this.txtEDMSUserID.Name = "txtEDMSUserID";
            this.txtEDMSUserID.Size = new System.Drawing.Size(240, 21);
            this.txtEDMSUserID.TabIndex = 29;
            // 
            // txtEDMSPassword
            // 
            this.txtEDMSPassword.BackColor = System.Drawing.Color.White;
            this.txtEDMSPassword.Location = new System.Drawing.Point(161, 256);
            this.txtEDMSPassword.Name = "txtEDMSPassword";
            this.txtEDMSPassword.PasswordChar = '*';
            this.txtEDMSPassword.Size = new System.Drawing.Size(240, 21);
            this.txtEDMSPassword.TabIndex = 28;
            // 
            // lblEDMSUserID
            // 
            this.lblEDMSUserID.Location = new System.Drawing.Point(30, 233);
            this.lblEDMSUserID.Name = "lblEDMSUserID";
            this.lblEDMSUserID.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblEDMSUserID.Size = new System.Drawing.Size(114, 25);
            this.lblEDMSUserID.TabIndex = 27;
            this.lblEDMSUserID.Text = "UserID:";
            this.lblEDMSUserID.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblFileNameMessageCopy
            // 
            this.lblFileNameMessageCopy.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileNameMessageCopy.ForeColor = System.Drawing.Color.MediumBlue;
            this.lblFileNameMessageCopy.Location = new System.Drawing.Point(158, 195);
            this.lblFileNameMessageCopy.Name = "lblFileNameMessageCopy";
            this.lblFileNameMessageCopy.Size = new System.Drawing.Size(393, 34);
            this.lblFileNameMessageCopy.TabIndex = 26;
            // 
            // lblFileNameMessageCopyLabel
            // 
            this.lblFileNameMessageCopyLabel.Location = new System.Drawing.Point(30, 189);
            this.lblFileNameMessageCopyLabel.Name = "lblFileNameMessageCopyLabel";
            this.lblFileNameMessageCopyLabel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblFileNameMessageCopyLabel.Size = new System.Drawing.Size(114, 25);
            this.lblFileNameMessageCopyLabel.TabIndex = 25;
            this.lblFileNameMessageCopyLabel.Text = "File Name Status:";
            this.lblFileNameMessageCopyLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblFileNameCopy
            // 
            this.lblFileNameCopy.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileNameCopy.ForeColor = System.Drawing.Color.MediumBlue;
            this.lblFileNameCopy.Location = new System.Drawing.Point(158, 170);
            this.lblFileNameCopy.Name = "lblFileNameCopy";
            this.lblFileNameCopy.Size = new System.Drawing.Size(406, 19);
            this.lblFileNameCopy.TabIndex = 24;
            // 
            // lblFileNameCopyLabel
            // 
            this.lblFileNameCopyLabel.Location = new System.Drawing.Point(30, 164);
            this.lblFileNameCopyLabel.Name = "lblFileNameCopyLabel";
            this.lblFileNameCopyLabel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblFileNameCopyLabel.Size = new System.Drawing.Size(114, 25);
            this.lblFileNameCopyLabel.TabIndex = 23;
            this.lblFileNameCopyLabel.Text = "New Doc Name:";
            this.lblFileNameCopyLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtVariantTextCopy
            // 
            this.txtVariantTextCopy.BackColor = System.Drawing.Color.White;
            this.txtVariantTextCopy.Location = new System.Drawing.Point(161, 142);
            this.txtVariantTextCopy.Name = "txtVariantTextCopy";
            this.txtVariantTextCopy.Size = new System.Drawing.Size(240, 21);
            this.txtVariantTextCopy.TabIndex = 22;
            this.txtVariantTextCopy.TextChanged += new System.EventHandler(this.txtVariantTextCopy_TextChanged);
            // 
            // lblVariantText
            // 
            this.lblVariantText.Location = new System.Drawing.Point(30, 139);
            this.lblVariantText.Name = "lblVariantText";
            this.lblVariantText.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblVariantText.Size = new System.Drawing.Size(114, 25);
            this.lblVariantText.TabIndex = 21;
            this.lblVariantText.Text = "Varnt/Base/Matl:";
            this.lblVariantText.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // optProductOrderByTitleCopy
            // 
            this.optProductOrderByTitleCopy.AutoSize = true;
            this.optProductOrderByTitleCopy.Location = new System.Drawing.Point(268, 118);
            this.optProductOrderByTitleCopy.Name = "optProductOrderByTitleCopy";
            this.optProductOrderByTitleCopy.Size = new System.Drawing.Size(45, 17);
            this.optProductOrderByTitleCopy.TabIndex = 20;
            this.optProductOrderByTitleCopy.TabStop = true;
            this.optProductOrderByTitleCopy.Text = "Title";
            this.optProductOrderByTitleCopy.UseVisualStyleBackColor = true;
            this.optProductOrderByTitleCopy.Click += new System.EventHandler(this.optProductOrderByCopy_Click);
            // 
            // optProductOrderByNumberCopy
            // 
            this.optProductOrderByNumberCopy.AutoSize = true;
            this.optProductOrderByNumberCopy.Location = new System.Drawing.Point(161, 118);
            this.optProductOrderByNumberCopy.Name = "optProductOrderByNumberCopy";
            this.optProductOrderByNumberCopy.Size = new System.Drawing.Size(62, 17);
            this.optProductOrderByNumberCopy.TabIndex = 19;
            this.optProductOrderByNumberCopy.TabStop = true;
            this.optProductOrderByNumberCopy.Text = "Number";
            this.optProductOrderByNumberCopy.UseVisualStyleBackColor = true;
            this.optProductOrderByNumberCopy.Click += new System.EventHandler(this.optProductOrderByCopy_Click);
            // 
            // lblOrderByCopy
            // 
            this.lblOrderByCopy.Location = new System.Drawing.Point(30, 114);
            this.lblOrderByCopy.Name = "lblOrderByCopy";
            this.lblOrderByCopy.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblOrderByCopy.Size = new System.Drawing.Size(114, 25);
            this.lblOrderByCopy.TabIndex = 18;
            this.lblOrderByCopy.Text = "Order by:";
            this.lblOrderByCopy.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cboDocNameCopy
            // 
            this.cboDocNameCopy.DropDownHeight = 90;
            this.cboDocNameCopy.DropDownWidth = 390;
            this.cboDocNameCopy.FormattingEnabled = true;
            this.cboDocNameCopy.IntegralHeight = false;
            this.cboDocNameCopy.Location = new System.Drawing.Point(161, 92);
            this.cboDocNameCopy.Name = "cboDocNameCopy";
            this.cboDocNameCopy.Size = new System.Drawing.Size(240, 21);
            this.cboDocNameCopy.TabIndex = 17;
            this.cboDocNameCopy.SelectedIndexChanged += new System.EventHandler(this.cboDocNameCopy_SelectedIndexChanged);
            // 
            // lblDocName
            // 
            this.lblDocName.Location = new System.Drawing.Point(30, 89);
            this.lblDocName.Name = "lblDocName";
            this.lblDocName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblDocName.Size = new System.Drawing.Size(114, 25);
            this.lblDocName.TabIndex = 16;
            this.lblDocName.Text = "Document Name:";
            this.lblDocName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cboDocLocation
            // 
            this.cboDocLocation.FormattingEnabled = true;
            this.cboDocLocation.Location = new System.Drawing.Point(161, 67);
            this.cboDocLocation.Name = "cboDocLocation";
            this.cboDocLocation.Size = new System.Drawing.Size(240, 21);
            this.cboDocLocation.TabIndex = 15;
            this.cboDocLocation.SelectedIndexChanged += new System.EventHandler(this.cboDocLocation_SelectedIndexChanged);
            // 
            // lblDocLocation
            // 
            this.lblDocLocation.Location = new System.Drawing.Point(30, 64);
            this.lblDocLocation.Name = "lblDocLocation";
            this.lblDocLocation.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblDocLocation.Size = new System.Drawing.Size(114, 25);
            this.lblDocLocation.TabIndex = 14;
            this.lblDocLocation.Text = "Document Location:";
            this.lblDocLocation.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cboDocTypeCopy
            // 
            this.cboDocTypeCopy.FormattingEnabled = true;
            this.cboDocTypeCopy.Location = new System.Drawing.Point(161, 42);
            this.cboDocTypeCopy.Name = "cboDocTypeCopy";
            this.cboDocTypeCopy.Size = new System.Drawing.Size(240, 21);
            this.cboDocTypeCopy.TabIndex = 13;
            this.cboDocTypeCopy.SelectedIndexChanged += new System.EventHandler(this.cboDocTypeCopy_Change);
            // 
            // lblRecordTypeCopy
            // 
            this.lblRecordTypeCopy.Location = new System.Drawing.Point(30, 39);
            this.lblRecordTypeCopy.Name = "lblRecordTypeCopy";
            this.lblRecordTypeCopy.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblRecordTypeCopy.Size = new System.Drawing.Size(114, 25);
            this.lblRecordTypeCopy.TabIndex = 11;
            this.lblRecordTypeCopy.Text = "Document Type:";
            this.lblRecordTypeCopy.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cboOpUnitCopy
            // 
            this.cboOpUnitCopy.FormattingEnabled = true;
            this.cboOpUnitCopy.Location = new System.Drawing.Point(161, 17);
            this.cboOpUnitCopy.Name = "cboOpUnitCopy";
            this.cboOpUnitCopy.Size = new System.Drawing.Size(121, 21);
            this.cboOpUnitCopy.TabIndex = 10;
            this.cboOpUnitCopy.SelectedIndexChanged += new System.EventHandler(this.cboOpUnitCopy_SelectedIndexChanged);
            // 
            // lblOpUnitCopy
            // 
            this.lblOpUnitCopy.Location = new System.Drawing.Point(30, 14);
            this.lblOpUnitCopy.Name = "lblOpUnitCopy";
            this.lblOpUnitCopy.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblOpUnitCopy.Size = new System.Drawing.Size(114, 25);
            this.lblOpUnitCopy.TabIndex = 1;
            this.lblOpUnitCopy.Text = "Operating Unit:";
            this.lblOpUnitCopy.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.mapStatusStart);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(125, 266);
            this.panel1.TabIndex = 30;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.Black;
            this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(73, 239);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(37, 26);
            this.label4.TabIndex = 29;
            this.label4.Tag = "0";
            this.label4.Text = "Finish";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label4.UseMnemonic = false;
            this.label4.Click += new System.EventHandler(this.cmdFinish_Click);
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.Black;
            this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(73, 7);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 37);
            this.label2.TabIndex = 28;
            this.label2.Tag = "0";
            this.label2.Text = "File New";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label2.UseMnemonic = false;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Red;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(15, 244);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(15, 15);
            this.label3.TabIndex = 27;
            this.label3.Tag = "Exclude";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Lime;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(58, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(15, 15);
            this.label1.TabIndex = 25;
            this.label1.Tag = "0";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // mapStatusStart
            // 
            this.mapStatusStart.BackColor = System.Drawing.Color.Red;
            this.mapStatusStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mapStatusStart.Location = new System.Drawing.Point(15, 18);
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
            this.label7.Location = new System.Drawing.Point(18, 27);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(50, 228);
            this.label7.TabIndex = 23;
            this.label7.Tag = "Exclude";
            // 
            // FormFileNewWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(714, 460);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.mpgWizardPage);
            this.Controls.Add(this.pnlCmdButtons);
            this.Controls.Add(this.grpFileOptions);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "FormFileNewWizard";
            this.Text = "File New Wizard";
            this.Activated += new System.EventHandler(this.FormFileNewWizard_Activated);
            this.Load += new System.EventHandler(this.FormFileNewWizard_Load);
            this.grpFileOptions.ResumeLayout(false);
            this.pnlCmdButtons.ResumeLayout(false);
            this.mpgWizardPage.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpFileOptions;
        private System.Windows.Forms.RadioButton optFileCreate;
        private System.Windows.Forms.RadioButton optFileCopyDiff;
        private System.Windows.Forms.RadioButton optFileCopyVers;
        private System.Windows.Forms.Panel pnlCmdButtons;
        private System.Windows.Forms.Button cmdFinish;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.TabControl mpgWizardPage;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label lblFileNameNewMessageLabel;
        private System.Windows.Forms.Label lblRecordTypeNew;
        private System.Windows.Forms.Label lblWCEquip;
        private System.Windows.Forms.Label lblStepCode;
        private System.Windows.Forms.Label lblOrderByCreate;
        private System.Windows.Forms.Label lblProduct;
        private System.Windows.Forms.Label lblFreeTextCreate;
        private System.Windows.Forms.Label lblFileNameLabel;
        private System.Windows.Forms.Label lblOpUnitNew;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.RadioButton optProductOrderByTitleCreate;
        private System.Windows.Forms.RadioButton optProductOrderByNumberCreate;
        private System.Windows.Forms.ComboBox cboDocTypeCreate;
        private System.Windows.Forms.ComboBox cboWorkCenter;
        private System.Windows.Forms.ComboBox cboProductCreate;
        private System.Windows.Forms.ComboBox cboOpUnitCreate;
        private System.Windows.Forms.TextBox txtFreeTextCreate;
        private System.Windows.Forms.TextBox txtStepCode;
        private System.Windows.Forms.Label lblFileNameCreate;
        private System.Windows.Forms.Label lblFileNameMessageCreate;
        private System.Windows.Forms.ComboBox cboDocTypeCopy;
        private System.Windows.Forms.Label lblRecordTypeCopy;
        private System.Windows.Forms.ComboBox cboOpUnitCopy;
        private System.Windows.Forms.Label lblOpUnitCopy;
        private System.Windows.Forms.Label lblEDMSPassword;
        private System.Windows.Forms.TextBox txtEDMSUserID;
        private System.Windows.Forms.TextBox txtEDMSPassword;
        private System.Windows.Forms.Label lblEDMSUserID;
        private System.Windows.Forms.Label lblFileNameMessageCopy;
        private System.Windows.Forms.Label lblFileNameMessageCopyLabel;
        private System.Windows.Forms.Label lblFileNameCopy;
        private System.Windows.Forms.Label lblFileNameCopyLabel;
        private System.Windows.Forms.TextBox txtVariantTextCopy;
        private System.Windows.Forms.Label lblVariantText;
        private System.Windows.Forms.RadioButton optProductOrderByTitleCopy;
        private System.Windows.Forms.RadioButton optProductOrderByNumberCopy;
        private System.Windows.Forms.Label lblOrderByCopy;
        private System.Windows.Forms.ComboBox cboDocNameCopy;
        private System.Windows.Forms.Label lblDocName;
        private System.Windows.Forms.ComboBox cboDocLocation;
        private System.Windows.Forms.Label lblDocLocation;
        private System.Windows.Forms.ComboBox cboDocBaseNameCreate;
        private System.Windows.Forms.ComboBox cboDocBaseNameCopy;
        private System.Windows.Forms.ComboBox cboProductCopy;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label mapStatusStart;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button cmdNext;
        private System.Windows.Forms.Button cmdBack;
        private System.Windows.Forms.Label lblDocumentSupercedes;
        private System.Windows.Forms.ComboBox cboDocumentSupercedes;
    }
}