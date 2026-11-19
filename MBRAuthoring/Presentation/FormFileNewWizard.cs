using MBRAuthoring.Business;
using MBRAuthoring.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices.Protocols;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using QDOCS.Services;
using MBRAuthoring.Utilities;

namespace MBRAuthoring.Presentation
{
    public partial class FormFileNewWizard : FormBase
    {
        const string MyCompName = "frmFileNewWizard";
        private const string LOCKED = "Locked";
        private const string AUTHORING = "Authoring";
        private const string APPROVED = "Approved";
        private const string BLOCKED = "Blocked";
        private const string FREETEXT = "Free Text:";
        private const string MODULEID = "Module ID:";
        private List<TblOpUnit> opUnits = null;
        private TblOpUnit _tblOpUnit;
        private OpUnitDocType _opUnitDocType;
        protected override Panel progressBarPanel => panel1;
        private IDocumentManager _documentManager = null;// ServiceLocator.GetDocumentManager();
                                                         // public ICommandBarManager _commandBarManager = null;// ServiceLocator.GetCommandBarManager();
        string mstrOpUnitCreate;
        string mstrOpUnitCopy;
        int mintOpUnitIDCreate;
        int mintOpUnitIDCopy;
        string mstrOpUnitTypeCreate;
        string mstrOpUnitTypeCopy;
        string mstrDocTypeCreate;
        string mstrDocTypeCopy;
        string mstrDocTypeAbbCreate;
        string mstrDocTypeAbbCopy;
        int mintDocTypeIDCreate;
        int mintDocTypeIDCopy;
        string mstrWorkCenter;
        long mlngMaterialIDCreate;
        long mlngMaterialIDCopy;
        bool mblnInitializing;
        bool mblnCancel;
        bool mblnCalledWizard;
        int mintParmsValidated;
        const int cintParmsValidated = 31007;
        long mlngWorkCenterID;
        long mlngSuperDocID;
        bool mblnUseWorkCenterCreate;
        bool mblnUseWorkCenterCopy;
        bool mblnUseProductCreate;
        bool mblnUseProductCopy;
        bool mblnUseFreeTextCreate;
        bool mblnUseFreeTextCopy;
        bool mblnUseDocBaseNameCreate;
        bool mblnUseDocBaseNameCopy;
        int mlngMyTop;
        int mlngMyLeft;
        Document mdocNewDocument;
        qpsEditPropertyParms mqpsEditPropertyParms;
        private static readonly Regex _regex = new Regex(@"^[A-Za-z0-9\-_\s.]+$", RegexOptions.Compiled);
        List<Material> materials = null;

        public static string apiVersion = ConfigurationManager.AppSettings["API_Version"];
        public static string vaultDns = ConfigurationManager.AppSettings["QDOCS_vaultDns"];
        public static string username = GenericUtility.DecryptText(ConfigurationManager.AppSettings["QDOCS_Username"]);
        public static string password = GenericUtility.DecryptText(ConfigurationManager.AppSettings["QDOCS_Password"]);
        static string mailid;

        public FormFileNewWizard()
        {
            InitializeComponent();
            cboDocBaseNameCreate.DropDownHeight = 100;
            Opacity = 0;
            _documentManager = new DocumentManager();
        }

        private void FormFileNewWizard_Load(object sender, EventArgs e)
        {
            string MyProcName = "UserForm_Initialize";
            try
            {
                InitializeForm();
                //ShowCurrentTab(tabPage1);
                mblnInitializing = true;
                Cursor = Cursors.WaitCursor;
                //Call Form_Accelerators(Me);
                //Call Form_TabIndexing(Me);
                //Call Page_JumpTo(Me, 0);
                SetToolTip(cmdCancel, "Terminate Function with NO Output");  // NewCode
                SetToolTip(cmdFinish, "Verify Input, Generate Result, then Terminate Function");  // NewCode
                opUnits = _formManager.Fill_OpUnit();
                cboOpUnitCreate.DisplayMember = "OpUnit";
                cboOpUnitCreate.ValueMember = "OpUnitID";
                cboOpUnitCreate.DataSource = opUnits;
                SetToolTip(cboOpUnitCreate, "cboOpUnitCreate: OpUnits");
                cboOpUnitCopy.DisplayMember = "OpUnit";
                cboOpUnitCopy.ValueMember = "OpUnitID";
                cboOpUnitCopy.DataSource = opUnits;
                SetToolTip(cboOpUnitCopy, "cboOpUnitCopy: OpUnits");
                if (cboOpUnitCreate.Items.Count < 1)
                {
                    MessageBox.Show("NO OPUNIT LIST FOR THIS USER" + Environment.NewLine + Environment.NewLine + "USERNAME=" + CommonService.gstrUserName + "   USERID=" + CommonService.gintUserID + Environment.NewLine + Environment.NewLine + "FUNCTION TERMINATING", "FileNew - No OpUnit List", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                cboDocumentSupercedes.Visible = false;
                lblDocumentSupercedes.Visible = false;
                txtFreeTextCreate.Text = "-";
                this.Text = "New File Wizard";
                optProductOrderByNumberCreate.Checked = false;
                optProductOrderByNumberCreate.Checked = true;
                optProductOrderByCreate_Click(optProductOrderByNumberCreate, EventArgs.Empty);
                optProductOrderByNumberCopy.Checked = false;
                optProductOrderByNumberCopy.Checked = true;
                optProductOrderByCopy_Click(optProductOrderByNumberCopy, EventArgs.Empty);
                mblnInitializing = false;
                lblFileNameCreate__LOAD();
                cmdNext.Enabled = false;
                cmdBack.Enabled = false;
                grpFileOptions.Visible = true;
                //grpFileOptions.Top = 0; // Todo
                //grpFileOptions.Left = 114; //113.65
                //grpFileOptions.Height = 30;
                //grpFileOptions.Width = 378;
                optFileCreate.Checked = true;
                cboOpUnitCreate.Focus();
                cboDocBaseNameCopy.Top = txtVariantTextCopy.Top;
                cboDocBaseNameCopy.Left = txtVariantTextCopy.Left;
                cboProductCopy.Left = txtVariantTextCopy.Left;
                cboProductCopy.Top = txtVariantTextCopy.Top;
                Cursor = Cursors.Default;
                txtEDMSUserID.Text = Environment.UserName.ToUpper();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Error: " + ex.Message, MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void cboOpUnitCreate_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strDocType;
            if (cboOpUnitCreate.Items.Count <= 0) return;
            var selectedOpUnit = (TblOpUnit)cboOpUnitCreate.SelectedItem;
            mstrOpUnitCreate = selectedOpUnit.OpUnit;
            mintOpUnitIDCreate = selectedOpUnit.OpUnitID;
            mstrOpUnitTypeCreate = selectedOpUnit.OpUnitType;
            if (cboDocTypeCreate.Items.Count <= 0)
            {
                strDocType = "";
            }
            else
            {
                var selectedDocTypeCreate = (OpUnitDocType)cboDocTypeCreate.SelectedItem;
                strDocType = selectedDocTypeCreate.DocType;
            }
            var opUnitDocTypes = _formManager.Fill_OpUnit_DocType(mintOpUnitIDCreate, mstrOpUnitTypeCreate);
            cboDocTypeCreate.DisplayMember = "DocType";
            cboDocTypeCreate.ValueMember = "DocTypeID";
            cboDocTypeCreate.DataSource = opUnitDocTypes;
            SetToolTip(cboDocTypeCreate, "cboDocTypeCreate: OpUnitID= " + mintOpUnitIDCreate + " OpUnitType= " + mstrOpUnitTypeCreate + " OpUnitDocTypes");
            if (!string.IsNullOrEmpty(strDocType))
            {
                // Todo ComboBox_Default(cboDocTypeCreate, strDocType)
            }
        }

        private void cboDocTypeCreate_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboDocTypeCreate.Items.Count <= 0)
            {
                mstrDocTypeCreate = "";
                mstrDocTypeAbbCreate = "";
                mintDocTypeIDCreate = -1;
                mblnUseWorkCenterCreate = false;
                mblnUseProductCreate = false;
                mblnUseFreeTextCreate = false;
            }
            else
            {
                OpUnitDocType opUnitDocType = (OpUnitDocType)cboDocTypeCreate.SelectedItem;
                mstrDocTypeCreate = opUnitDocType.DocType;
                mstrDocTypeAbbCreate = opUnitDocType.DocTypeAbb;
                mintDocTypeIDCreate = opUnitDocType.DocTypeID;
                mblnUseWorkCenterCreate = opUnitDocType.Includes_WorkCenter;
                mblnUseFreeTextCreate = opUnitDocType.Includes_FreeText;
                mblnUseDocBaseNameCreate = opUnitDocType.Includes_DocBaseName;
                mblnUseProductCreate = opUnitDocType.Includes_ProductIdent;
            }

            lblWCEquip.Visible = mblnUseWorkCenterCreate;
            cboWorkCenter.Visible = mblnUseWorkCenterCreate;
            lblOrderByCreate.Visible = mblnUseProductCreate;
            optProductOrderByTitleCreate.Visible = mblnUseProductCreate;
            optProductOrderByNumberCreate.Visible = mblnUseProductCreate;
            lblProduct.Visible = mblnUseProductCreate;
            cboProductCreate.Visible = mblnUseProductCreate;

            if (mblnUseProductCreate)
            {
                if (!mblnInitializing)
                {
                    cboProductCreate__LOAD();
                }
            }
            lblStepCode.Visible = false;
            txtStepCode.Visible = false;

            if (mblnUseDocBaseNameCreate)
            {
                lblFreeTextCreate.Text = MODULEID;
                lblFreeTextCreate.Visible = true;
                txtFreeTextCreate.Visible = false;
                //cboDocBaseNameCreate.Height = 18; // design modifies, so not req mostly
                //cboDocBaseNameCreate.Left = 102;
                //cboDocBaseNameCreate.Top = 138;
                //cboDocBaseNameCreate.Width = 246;
                cboDocBaseNameCreate.Visible = true;
                cboDocBaseNameCreate.DropDownStyle = ComboBoxStyle.DropDownList;
                // todo
                cboDocBaseNameCreate.Visible = true;
                cboDocBaseNameCreate.Enabled = true;
                //cboDocBaseNameCreate.locked = false; //Not req (.enabled = true is achive the same func)
                cboDocBaseNameCreate__LOAD();
            }
            else
            {
                lblFreeTextCreate.Text = FREETEXT;
                lblFreeTextCreate.Visible = mblnUseFreeTextCreate;
                cboDocBaseNameCreate.Visible = false;
                //txtFreeTextCreate.Height = 18; // design modifies, so not req mostly
                //txtFreeTextCreate.Left = 102;
                //txtFreeTextCreate.Top = 138;
                //txtFreeTextCreate.Width = 246;
                txtFreeTextCreate.Visible = mblnUseFreeTextCreate;
            }
            if (!mblnUseWorkCenterCreate && !mblnUseProductCreate && mintDocTypeIDCreate != -1)
            {
                lblFileNameCreate__LOAD();
            }
        }

        private void cboOpUnitCopy_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strDocType;
            if (cboOpUnitCopy.Items.Count <= 0) return;

            var selectedOpUnitCopy = (TblOpUnit)cboOpUnitCopy.SelectedItem;
            mstrOpUnitCopy = selectedOpUnitCopy.OpUnit;
            mintOpUnitIDCopy = selectedOpUnitCopy.OpUnitID;
            mstrOpUnitTypeCopy = selectedOpUnitCopy.OpUnitType;

            if (cboDocTypeCopy.Items.Count <= 0)
            {
                strDocType = "";
            }
            else
            {
                var selectedDocTypeCopy = (OpUnitDocType)cboDocTypeCopy.SelectedItem;
                strDocType = selectedDocTypeCopy.DocType;
            }
            var opUnitDocTypes = _formManager.Fill_OpUnit_DocType(mintOpUnitIDCopy, mstrOpUnitTypeCopy);
            cboDocTypeCopy.DisplayMember = "DocType";
            cboDocTypeCopy.ValueMember = "DocTypeID";
            cboDocTypeCopy.DataSource = opUnitDocTypes;
            SetToolTip(cboDocTypeCopy, "cboDocTypeCopy: OpUnitID= " + mintOpUnitIDCopy + " OpUnitType= " + mstrOpUnitTypeCopy + " OpUnitDocTypes");
            if (!string.IsNullOrEmpty(strDocType))
            {
                // Todo ComboBox_Default(cboDocTypeCreate, strDocType)
            }
        }

        private void cboDocTypeCopy_Change(object sender, EventArgs e)
        {
            if (mblnInitializing) return;
            if (cboDocTypeCopy.Items.Count <= 0)
            {
                mstrDocTypeCopy = "";
                mstrDocTypeAbbCopy = "";
                mintDocTypeIDCopy = -1;
                mblnUseWorkCenterCopy = false;
                mblnUseProductCopy = false;
                mblnUseDocBaseNameCopy = false;
                mblnUseFreeTextCopy = false;
            }
            else
            {
                OpUnitDocType opUnitDocType = (OpUnitDocType)cboDocTypeCopy.SelectedItem;
                mstrDocTypeCopy = opUnitDocType.DocType;
                mstrDocTypeAbbCopy = opUnitDocType.DocTypeAbb;
                mintDocTypeIDCopy = opUnitDocType.DocTypeID;
                mblnUseWorkCenterCopy = opUnitDocType.Includes_WorkCenter;
                mblnUseFreeTextCopy = opUnitDocType.Includes_FreeText;
                mblnUseDocBaseNameCopy = opUnitDocType.Includes_DocBaseName;
                mblnUseProductCopy = opUnitDocType.Includes_ProductIdent;
            }
            lblWCEquip.Visible = mblnUseWorkCenterCopy;
            cboWorkCenter.Visible = mblnUseWorkCenterCopy;
            lblStepCode.Visible = false;
            txtStepCode.Visible = false;
            if (optFileCopyVers.Checked)
            {
                lblVariantText.Text = "Variant Text:";
                txtVariantTextCopy.Visible = true;
                cboDocBaseNameCopy.Visible = false;
                ProductCopy_Visible(false);
            }
            else if (optFileCopyDiff.Checked)
            {
                if (mblnUseProductCopy)
                {
                    lblVariantText.Text = "Product:";
                    txtVariantTextCopy.Visible = false;
                    ProductCopy_Visible(true);
                    cboProductCopy__LOAD();
                    cboDocBaseNameCopy.Visible = false;
                }
                if (mblnUseDocBaseNameCopy)
                {
                    lblVariantText.Text = "Base Name";
                    txtVariantTextCopy.Visible = false;
                    ProductCopy_Visible(false);
                    cboDocBaseNameCopy.Visible = true;
                    cboDocBaseNameCopy.DropDownStyle = ComboBoxStyle.DropDownList;
                    // Todo
                    // cboDocBaseNameCopy.AutoCompleteMode = AutoCompleteMode.Suggest;
                    //cboDocBaseNameCopy.AutoCompleteSource = AutoCompleteSource.ListItems;
                    cboDocBaseNameCopy.Enabled = true;
                    cboDocBaseNameCopy__LOAD();
                }
            }
            else
            {
                MessageBox.Show("NO COPY OPTION SET");
            }
            cboDocLocation__FILL();
        }

        private void cboDocLocation__FILL()
        {
            cboDocLocation.Items.Clear();
            cboDocLocation.Items.Add(APPROVED);
            if (optFileCopyDiff.Checked)
            {
                cboDocLocation.Items.Add(BLOCKED);
                cboDocLocation.Items.Add(AUTHORING);
                cboDocLocation.Items.Add(LOCKED);
            }
            else if (optFileCopyVers.Checked)
            {
                if (CommonService.gblnUserType_ADMIN)
                {
                    cboDocLocation.Items.Add(BLOCKED);
                }
            }

            if (cboDocLocation.Items.Count > 0)
            {
                cboDocLocation.SelectedIndex = 0;
            }
        }

        private void cboDocLocation_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool blnIncludeLocked;
            bool blnIncludeAuthoring;
            bool blnIncludeApproved;
            bool blnIncludeArchived;
            bool blnIncludeBlocked;
            bool TargetVisible;
            bool EDMSVisible;
            if (cboDocLocation.Items.Count == 0) return;
            blnIncludeLocked = (cboDocLocation.Text == LOCKED);
            blnIncludeAuthoring = (cboDocLocation.Text == AUTHORING);
            blnIncludeApproved = (cboDocLocation.Text == APPROVED);
            blnIncludeBlocked = (cboDocLocation.Text == BLOCKED);
            blnIncludeArchived = false;
            var files = _formManager.Fill_Document(mintOpUnitIDCopy, mintDocTypeIDCopy, blnIncludeLocked, blnIncludeAuthoring, blnIncludeApproved, blnIncludeArchived, blnIncludeBlocked);
            cboDocNameCopy.DisplayMember = "DisplayDocumentName";
            cboDocNameCopy.ValueMember = "DocID";
            cboDocNameCopy.DataSource = files;
            //added for dropdown display with 2 column
            cboDocNameCopy.DrawMode = DrawMode.OwnerDrawFixed;
            cboDocNameCopy.DrawItem += cboDocNameCopy_DrawItem;
            cboDocNameCopy.Format += cboDocNameCopye_Format;

            var selectedDocTypeCopy = (OpUnitDocType)cboDocTypeCopy.SelectedItem;
            SetToolTip(cboDocNameCopy, "cboDocNameCopy: OpUnitID=" + mintOpUnitIDCopy + " OpUnitType=" + selectedDocTypeCopy.DocTypeID + " DocTypeID=" + mintDocTypeIDCopy + " InclLocked=" + blnIncludeLocked);
            //cboDocNameCopy.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            //cboDocNameCopy.AutoCompleteSource = AutoCompleteSource.ListItems;
            TargetVisible = (cboDocNameCopy.Items.Count > 0);
            cboDocNameCopy.Visible = TargetVisible;
            lblFileNameCopyLabel.Visible = TargetVisible;
            lblFileNameMessageCreate.Visible = TargetVisible;
            if (!TargetVisible)
            {
                SetFinishButtonVisibility("not ok");
            }
            EDMSVisible = (cboDocLocation.Text == APPROVED) || (cboDocLocation.Text == BLOCKED);
            lblEDMSUserID.Visible = EDMSVisible;
            txtEDMSUserID.Visible = EDMSVisible;
            lblEDMSPassword.Visible = EDMSVisible;
            txtEDMSPassword.Visible = EDMSVisible;
        }

        private void cboDocNameCopye_Format(object sender, ListControlConvertEventArgs e)
        {
            var documentDetails = (DocumentDetails)e.ListItem;
            e.Value = documentDetails.DisplayDocumentName;
        }

        private void cboDocNameCopy_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            if (e.Index >= 0)
            {
                var documentDetails = (DocumentDetails)cboDocNameCopy.Items[e.Index];

                System.Drawing.Font drawFont = e.Font;
                int padding = 180;  // Space between DocumentName and DocumentTitle columns

                e.Graphics.DrawString(documentDetails.DisplayDocumentName, drawFont, Brushes.Black, e.Bounds.X, e.Bounds.Y);
                e.Graphics.DrawString(documentDetails.DocumentTitle, drawFont, Brushes.Black, e.Bounds.X + padding, e.Bounds.Y);
            }

            e.DrawFocusRectangle();
        }


        private void ProductCopy_Visible(bool isVisible)
        {
            lblOrderByCopy.Visible = isVisible;
            optProductOrderByTitleCopy.Visible = isVisible;
            optProductOrderByNumberCopy.Visible = isVisible;
            cboProductCopy.Visible = isVisible;
        }

        private void cboProductCopy__LOAD()
        {
            bool isProductOrderByNumber = true;
            cboProductCopy.Visible = false;
            //Me.Repaint
            this.Refresh();

            if (optProductOrderByNumberCopy.Checked)
            {
                isProductOrderByNumber = true;
            }
            else if (optProductOrderByTitleCopy.Checked)
            {
                isProductOrderByNumber = false;
            }
            var materials = _formManager.Fill_OpUnit_Product(mintOpUnitIDCopy, mstrOpUnitTypeCopy, isProductOrderByNumber);
            cboProductCopy.DisplayMember = "MaterialDisplay";
            cboProductCopy.ValueMember = "MaterialID";
            cboProductCopy.DataSource = materials;
            SetToolTip(cboProductCopy, "cboProductCopy: OpUnitID= " + mintOpUnitIDCopy + " OpUnitType= " + mstrOpUnitTypeCopy + " Material");
            cboProductCopy.Visible = true;
        }

        private void cboDocBaseNameCopy__LOAD()
        {
            var docBaseNameCopyDetails = _formManager.Fill_DocBaseName(mintDocTypeIDCopy, true);
            cboDocBaseNameCopy.DisplayMember = "DocumentBaseName";
            cboDocBaseNameCopy.ValueMember = "DocumentTitle";
            cboDocBaseNameCopy.DataSource = docBaseNameCopyDetails;
        }

        private void cboDocBaseNameCreate__LOAD()
        {
            var docBaseNameCreateDetails = _formManager.Fill_DocBaseName(mintDocTypeIDCreate, true);
            cboDocBaseNameCreate.DisplayMember = "DocumentBaseName";
            cboDocBaseNameCreate.ValueMember = "DocumentTitle";
            cboDocBaseNameCreate.DataSource = docBaseNameCreateDetails;

            cboDocBaseNameCreate.DrawMode = DrawMode.OwnerDrawFixed;
            cboDocBaseNameCreate.DrawItem += cboDocBaseNameCreate_DrawItem;
            cboDocBaseNameCreate.Format += cboDocBaseNameCreate_Format;

            //cboDocBaseNameCreate.AutoCompleteMode = AutoCompleteMode.None;
            //cboDocBaseNameCreate.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        private void cboDocBaseNameCreate_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            if (e.Index >= 0)
            {
                var docBaseNameDetails = (DocBaseNameDetails)cboDocBaseNameCreate.Items[e.Index];

                System.Drawing.Font drawFont = e.Font;
                int padding = 180;  // Space between DocumentName and DocumentTitle columns

                e.Graphics.DrawString(docBaseNameDetails.DocumentBaseName, drawFont, Brushes.Black, e.Bounds.X, e.Bounds.Y);
                e.Graphics.DrawString(docBaseNameDetails.DocumentTitle, drawFont, Brushes.Black, e.Bounds.X + padding, e.Bounds.Y);
            }

            e.DrawFocusRectangle();
        }

        private void cboDocBaseNameCreate_Format(object sender, ListControlConvertEventArgs e)
        {
            var documentDetails = (DocBaseNameDetails)e.ListItem;
            e.Value = documentDetails.DocumentBaseName;
        }

        private void cboProductCreate__LOAD()
        {
            bool isProductOrderByNumber = true;
            cboProductCreate.Visible = false;
            cboProductCreate.Refresh();
            if (optProductOrderByTitleCreate.Checked)
            {
                isProductOrderByNumber = false;
            }
            else if (optProductOrderByNumberCreate.Checked)
            {
                isProductOrderByNumber = true;
            }

            materials = _formManager.Fill_OpUnit_Product(mintOpUnitIDCreate, mstrOpUnitTypeCreate, isProductOrderByNumber);
            //cboProductCreate.AutoCompleteMode = AutoCompleteMode.None;
            //cboProductCreate.AutoCompleteSource = AutoCompleteSource.None; // --------------
            cboProductCreate.DisplayMember = "MaterialDisplay";
            cboProductCreate.ValueMember = "MaterialID";
            cboProductCreate.DataSource = materials;
            SetToolTip(cboProductCreate, "cboProductCreate: OpUnitID= " + mintOpUnitIDCreate + " OpUnitType= " + mstrOpUnitTypeCreate + " Material");
            //cboProductCreate.AutoCompleteMode = AutoCompleteMode.Suggest;
            //cboProductCreate.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboProductCreate.Visible = true;
        }

        private void lblFileNameCreate__LOAD()
        {
            // TBD
            string MyProcName = "lblFileNameCreate__LOAD";
            if (mblnInitializing) return;

            lblFileNameCreate.Text = "Building Name...";

            if (!ValidateParams(txtVariantTextCopy.Text, "FreeText")) return;
            if (!ValidateParams(cboDocBaseNameCreate.Text, "DocBaseName")) return;

            var resultSet = _formManager.BuildDocumentName(mintOpUnitIDCreate, mintDocTypeIDCreate, mlngMaterialIDCreate, mlngWorkCenterID, txtFreeTextCreate.Text, cboDocBaseNameCreate.Text);


            lblFileNameCreate.Text = resultSet.DocumentReturn;
            string errorReturn = resultSet.ErrorReturn;
            lblFileNameMessageCreate.Text = errorReturn;
            SetFinishButtonVisibility(errorReturn);

            //Dictionary<string, object> item2 = resultSet.Item2;
            //if (item2.ContainsKey("@DocumentReturn"))
            //{
            //    string documentReturn = item2["@DocumentReturn"].ToString();
            //    lblFileNameCreate.Text = documentReturn;
            //}
            //if (item2.ContainsKey("@ErrorReturn"))
            //{
            //    string errorReturn = item2["@ErrorReturn"].ToString();
            //    lblFileNameMessageCreate.Text = errorReturn;
            //    SetFinishButtonVisibility(errorReturn);
            //}
        }

        private void InitializeForm()
        {
            foreach (Control ctrl in progressBarPanel.Controls) //pgb
            {
                if (ctrl is Label label)
                {
                    if (label.Tag == null || label.Tag.ToString() != "Exclude")
                    {
                        label.Click += ProgressLabel_Click;
                    }
                }
            }
        }

        private void optFileOptions_Click(object sender, EventArgs e)
        {
            RadioButton radioButton = sender as RadioButton;
            if (radioButton != null)
            {
                RadioButton fileOption = (RadioButton)sender;
                Cursor = Cursors.WaitCursor;
                switch (fileOption.Name)
                {
                    case "optFileCreate":
                        ShowCurrentTab(tabPage1);
                        break;
                    case "optFileCopyVers":
                    case "optFileCopyDiff":
                        ShowCurrentTab(tabPage2);
                        cboDocTypeCopy_Change(null, null);
                        SetToolTip(txtVariantTextCopy, "Allowed: A-Z, a-z, 0-9, hyphen: - underscore: _ space:  or a period: .");
                        break;
                }
                Cursor = Cursors.Default;
            }
        }

        private void ShowCurrentTab(TabPage tabPage)
        {
            while (mpgWizardPage.TabPages.Count > 0)
            {
                mpgWizardPage.TabPages.RemoveAt(mpgWizardPage.TabPages.Count - 1);
            }
            mpgWizardPage.TabPages.Add(tabPage);
            mpgWizardPage.Appearance = TabAppearance.Buttons;
            mpgWizardPage.SizeMode = TabSizeMode.Fixed;
            mpgWizardPage.ItemSize = new Size(0, 1);
            Opacity = 1;

            Document currentActiveDocument = Globals.ThisDocument.ActiveDocument;//currentDocument;
            currentActiveDocument.Application.StatusBar = "";
        }

        private void txtFreeTextCreate_TextChanged(object sender, EventArgs e)
        {
            lblFileNameCreate__LOAD();
        }

        private void optProductOrderByCreate_Click(object sender, EventArgs e)
        {
            cboProductCreate__LOAD();
        }

        private void cboProductCreate_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboProductCreate.Items.Count <= 0)
            {
                mlngMaterialIDCreate = 0;
                return;
            }
            var selectedcboProductCreate = (Material)cboProductCreate.SelectedItem;
            mlngMaterialIDCreate = selectedcboProductCreate.MaterialID;
            lblFileNameCreate__LOAD();
        }

        private void cboDocBaseNameCreate_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (cboDocBaseNameCreate.Items.Count > 0) { cboDocBaseNameCreate.SelectedIndex = 0; }
            var selectedItem = (DocBaseNameDetails)cboDocBaseNameCreate.SelectedItem;
            if (selectedItem != null)
            {
                cboDocBaseNameCreate.Text = selectedItem.DocumentBaseName;
            }

            lblFileNameCreate__LOAD();
        }

        private void cboDocNameCopy_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboDocNameCopy.Items.Count <= 0) return;
            if (optFileCopyDiff.Checked)
            {
                lblFileNameCopyDiff__LOAD();
            }
            else if (optFileCopyVers.Checked)
            {
                lblFileNameCopyVers__LOAD();
            }
        }

        private void lblFileNameCopyVers__LOAD()
        {
            string MyProcName = "lblFileNameCopyVers__LOAD";
            if (mblnInitializing) return;

            lblFileNameMessageCopy.Text = "Checking...";
            DocumentDetails docDetails = (DocumentDetails)cboDocNameCopy.SelectedItem;

            var resultSet = _formManager.BuildDocumentNameVersion(docDetails.DocID, txtVariantTextCopy.Text);

            lblFileNameCopy.Text = resultSet.DocumentReturn;
            string errorReturn = resultSet.ErrorReturn;
            lblFileNameMessageCopy.Text = errorReturn;
            SetFinishButtonVisibility(errorReturn);

            //Dictionary<string, object> item2 = resultSet.Item2;
            //if (item2.ContainsKey("@DocumentReturn"))
            //{
            //    string documentReturn = item2["@DocumentReturn"].ToString();
            //    lblFileNameCopy.Text = documentReturn;
            //}
            //if (item2.ContainsKey("@ErrorReturn"))
            //{
            //    string errorReturn = item2["@ErrorReturn"].ToString();
            //    lblFileNameMessageCopy.Text = errorReturn;
            //    SetFinishButtonVisibility(errorReturn);
            //}
        }

        private void lblFileNameCopyDiff__LOAD()
        {
            // TBD
            string MyProcName = "lblFileNameCopyDiff__LOAD";
            if (mblnInitializing) return;

            lblFileNameCopy.Text = "Building Name...";

            var resultSet = _formManager.BuildDocumentName(mintOpUnitIDCopy, mintDocTypeIDCopy, mlngMaterialIDCopy, mlngWorkCenterID, txtVariantTextCopy.Text, cboDocBaseNameCopy.Text);

            if (!ValidateParams(txtVariantTextCopy.Text, "FreeText")) return;
            if (!ValidateParams(cboDocBaseNameCopy.Text, "DocBaseName")) return;


            lblFileNameCopy.Text = resultSet.DocumentReturn;
            string errorReturn = resultSet.ErrorReturn;
            lblFileNameMessageCopy.Text = errorReturn;
            SetFinishButtonVisibility(errorReturn);

            //Dictionary<string, object> item2 = resultSet.Item2;
            //if (item2.ContainsKey("@DocumentReturn"))
            //{
            //    string documentReturn = item2["@DocumentReturn"].ToString();
            //    lblFileNameCopy.Text = documentReturn;
            //}
            //if (item2.ContainsKey("@ErrorReturn"))
            //{
            //    string errorReturn = item2["@ErrorReturn"].ToString();
            //    lblFileNameMessageCopy.Text = errorReturn;
            //    SetFinishButtonVisibility(errorReturn);
            //}
        }

        private void optProductOrderByCopy_Click(object sender, EventArgs e)
        {
            cboProductCopy__LOAD();
        }

        private void cboProductCopy_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboProductCopy.Items.Count <= 0)
            {
                mlngMaterialIDCopy = 0;
                return;
            }
            var selectedcboProductCopy = (Material)cboProductCopy.SelectedItem;
            mlngMaterialIDCopy = selectedcboProductCopy.MaterialID;
            lblFileNameCopyDiff__LOAD();
        }

        private void SetFinishButtonVisibility(string errorReturn)
        {
            cmdFinish.Enabled = (errorReturn == "OK");// commented for testing
            //cmdFinish.Enabled = true;
        }

        private bool ValidateParams(string sourceControl, string paramName)
        {
            switch (paramName)
            {
                case "FreeText":
                    if (sourceControl.Length > 30)
                    {
                        MessageBox.Show($"ERROR: @sourceControl MUST NOT BE MORE THAN 30, DATA LENGTH ERROR FOR @sourceControl", sourceControl);
                        return false;
                    }
                    return true;
                case "DocBaseName":
                    if (sourceControl.Length > 100)
                    {
                        MessageBox.Show($"ERROR: @sourceControl MUST NOT BE MORE THAN 100, DATA LENGTH ERROR FOR @sourceControl", sourceControl);
                        return false;
                    }
                    return true;
                default:
                    return false;
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }

        public static bool IsValidFreeText(string inputText)
        {
            if (string.IsNullOrEmpty(inputText))
            {
                return true;
            }

            foreach (char ch in inputText)
            {
                if (!char.IsLetter(ch) && !char.IsDigit(ch) && ch != '-' && ch != '_' && ch != ' ' && ch != '.')
                {
                    return false;
                }
            }

            return true;
        }

        private void FormFileNewWizard_Activated(object sender, EventArgs e)
        {
            if (cboOpUnitCreate.Items.Count < 1)
            {
                this.Close();
                CommonService.CleanUpBackgroundWordProcesses();
            }
            Cursor = Cursors.Default;
        }

        private void cboDocBaseNameCopy_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (cboDocBaseNameCopy.Items.Count > 0) { cboDocBaseNameCopy.SelectedIndex = 0; }
            lblFileNameCopyDiff__LOAD();
        }

        private void txtVariantTextCopy_TextChanged(object sender, EventArgs e)
        {
            if (txtVariantTextCopy.Text != (txtVariantTextCopy.Text).Trim())
            {
                txtVariantTextCopy.Text = (txtVariantTextCopy.Text).Trim();
            }
            lblFileNameCopyVers__LOAD();
        }

        private void cmdFinish_Click(object sender, EventArgs e)
        {
            string MyProcName = "cmdFinish_Click";
            Cursor = Cursors.WaitCursor;
            _tblOpUnit = (TblOpUnit)cboOpUnitCreate.SelectedItem;
            _opUnitDocType = (OpUnitDocType)cboDocTypeCreate.SelectedItem;
            SaveManager.SetOpUnit(_tblOpUnit, _opUnitDocType); // 2301
            if (optFileCreate.Checked)
            {
                DialogResult result = MessageBox.Show("Do you want to create a NEW DOCUMENT" + Environment.NewLine + Environment.NewLine + "\t" + lblFileNameCreate.Text, "File NEW", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    File_Create();
                }
            }
            else if (optFileCopyVers.Checked)
            {
                DialogResult result = MessageBox.Show("Do you want to COPY the existing Document" + Environment.NewLine + Environment.NewLine + "\t" + cboDocNameCopy.Text + Environment.NewLine + Environment.NewLine + "to" + Environment.NewLine + Environment.NewLine + "\t" + lblFileNameCopy.Text, "File Copy", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    File_CopyVers();
                }
            }
            else if (optFileCopyDiff.Checked)
            {
                DialogResult result = MessageBox.Show("Do you want to COPY the existing Document" + Environment.NewLine + Environment.NewLine + "\t" + cboDocNameCopy.Text + Environment.NewLine + Environment.NewLine + "to a Different BaseName Document:" + Environment.NewLine + Environment.NewLine + "\t" + lblFileNameCopy.Text, "File Copy", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    File_CopyDiff();
                }
            }
            Cursor = Cursors.Default;
        }

        private async void File_CopyDiff()
        {
            Document currentActiveDocument = Globals.ThisDocument.ActiveDocument;//currentDocument;

            try
            {
                string MyProcName = "File_CopyDiff";
                string CommonFilePath;
                int DocIndex;
                long DocID_Source;
                string DocumentName_Source;
                string DocumentPath_Source;
                string DocPasswordRead;
                string DocPasswordWrite;
                string DocPasswordProtect;
                string DocumentName_Target = string.Empty;
                string DocumentPath_Target;
                int SPParamCount;
                bool ReProtect;
                Document docTarget;
                //Range aStory;
                //Field aField;
                Template wtpTemplate;
                string strDocumentTemplateFullName;
                string strExpectedTemplateFullName;
                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorWait;
                currentActiveDocument.Application.StatusBar = " ";
                bool isValidUser = false;
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                var service = new QDOCSService(vaultDns, apiVersion);
                switch (cboDocLocation.Text)
                {
                    case APPROVED:
                        if (!ValidateUserCreds()) return;
                        string strPrevDir = Directory.GetCurrentDirectory(); // required?
                        string strCurrDir = currentActiveDocument.get_AttachedTemplate().Path;
                        Directory.SetCurrentDirectory(strCurrDir);                      
                        var userCheckResponse = await service.AuthenticateAsync(username, password); //Authentication using username, password using QDOCS 
                        var userMailId = await service.ValidateUserByEmailAsync(mailid); //Validating using the user mail Id

                        if (userCheckResponse.IsSuccess && userMailId.IsSuccess)
                        {
                            isValidUser = true;
                        }
                        else
                        {
                            isValidUser = false;
                            MessageBox.Show("Status Code: " + userCheckResponse.StatusCode + "\n" + (userCheckResponse.ErrorMessage != null ? "Error: " + userCheckResponse.ErrorMessage : " ") + "\n" + (userCheckResponse.ErrorDetails != null ? "Error Details: " + userCheckResponse.ErrorDetails : " "), "FAILED to connect to Quality DOCS - COPY process terminated", MessageBoxButtons.OK);
                            return;
                        }
                        Directory.SetCurrentDirectory(strPrevDir); // required??
                        break;
                    default:
                        break;
                }
                OpUnitDocType opUnitDocType = (OpUnitDocType)cboDocTypeCopy.SelectedItem;

                //Not Req
                //if(ModChangeHistory.gstrAppVersionMajor == "3")//(CommonService.gstrAppVersionMajor == "3") // Not req
                //{
                //    CommonFilePath = _formManager.AppStatus_Select("DocumentPath") + "\\" + cboOpUnitCopy.Text;
                //}
                //else
                //{
                //    CommonFilePath = opUnitDocType.DocumentPath;
                //}

                CommonFilePath = opUnitDocType.DocumentPath;

                DocumentDetails docDetails = (DocumentDetails)cboDocNameCopy.SelectedItem;
                DocID_Source = docDetails.DocID;
                DocumentName_Source = docDetails.DocumentName; //.Column(qdcColDocNameBase, DocIndex)
                DocPasswordRead = docDetails.DocPasswordRead;
                DocPasswordWrite = docDetails.DocPasswordWrite;
                DocumentPath_Source = CommonFilePath + "\\" + DocumentName_Source;
                currentActiveDocument.Application.StatusBar = "Copying to File " + lblFileNameCopy.Text;
                DocumentResponse docCheckResponse = null;
                bool isValidDoc = false;
                string srcDocName, srcDocExtension = string.Empty;

                switch (cboDocLocation.Text)
                {
                    case APPROVED:
                        if (isValidUser)
                        {
                            string docNameWithoutExt = Path.GetFileNameWithoutExtension(DocumentName_Source);
                            docCheckResponse = await service.GetDocumentIdByNameAsync(docNameWithoutExt);
                            if (docCheckResponse.IsSuccess)
                            {
                                isValidDoc = true;
                            }
                            else
                            {
                                isValidDoc = false;
                                MessageBox.Show("Document '" + DocumentName_Source + "' not found", MyProcName, MessageBoxButtons.OK);
                                MessageBox.Show((docCheckResponse.ErrorMessage != null ? "Error: " + DocumentName_Source + " - " + docCheckResponse.ErrorMessage : " ") + "\n" + (docCheckResponse.ErrorDetails != null ? "Error Details: " + docCheckResponse.ErrorDetails : " "), "FAILED to connect to Quality DOCS  - COPY process terminated", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
                        }
                        break;
                    default:
                        break;
                }

                TblOpUnit opUnit = (TblOpUnit)cboOpUnitCopy.SelectedItem;
                Material material = (Material)cboProductCopy.SelectedItem;
                DocBaseNameDetails docBaseNameDetails = (DocBaseNameDetails)cboDocBaseNameCopy.SelectedItem; //null exception
                string docBaseName = docBaseNameDetails?.DocumentBaseName ?? "";

                var resultSet = _formManager.File_CopyDiff(DocID_Source, opUnit.OpUnitID, opUnitDocType.DocTypeID, material.MaterialID, docBaseName); // docBaseNameDetails.DocumentBaseName);



                string errorReturn = resultSet.ErrorReturn;
                if (errorReturn != "OK") { MessageBox.Show(errorReturn); return; }
                DocumentName_Target = resultSet.DocumentReturn;

                //Dictionary<string, object> item2 = resultSet.Item2;
                //if (item2.ContainsKey("@ErrorReturn"))
                //{
                //    string errorReturn = item2["@ErrorReturn"].ToString();
                //    if (errorReturn != "OK") { MessageBox.Show(errorReturn); return; }
                //}
                //if (item2.ContainsKey("@DocumentReturn"))
                //{
                //    DocumentName_Target = item2["@DocumentReturn"].ToString();
                //}

                if (DocumentName_Target != lblFileNameCopy.Text)
                {
                    MessageBox.Show("WARNING: Expecting FileName: " + Environment.NewLine + Environment.NewLine + "\t" + lblFileNameCopy.Text + Environment.NewLine + Environment.NewLine + "Generated FileName:" + Environment.NewLine + Environment.NewLine + "\t" + DocumentName_Target, MyCompName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                DocumentPath_Target = CommonFilePath + "\\" + DocumentName_Target;
                switch (cboDocLocation.Text)
                {
                    case APPROVED:
                        if (isValidDoc)
                        {
                            try
                            {
                                string docObjectId = docCheckResponse.DocumentId;
                                await service.DownloadDocumentByIdAsync(docObjectId, DocumentPath_Target); //Exporting the document to path
                            }
                            catch (Exception ex)
                            {
                                Cursor = Cursors.Default;
                                MessageBox.Show("COPY Terminating, due to error in Quality DOCS Processing", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                MessageBox.Show("Error: Due to error in Quality DOCS Processing" + "\n" + "Error Details: " + ex.Message, "FAILED to connect to Quality DOCS - COPY process terminated", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                        break;
                    case AUTHORING:
                    case LOCKED:
                    case BLOCKED:
                        if (File.Exists(DocumentPath_Source)) // required?
                        {
                            File.Copy(DocumentPath_Source, DocumentPath_Target, overwrite: false);
                        }
                        break;
                    default:
                        MessageBox.Show("INVALID DOCLOCATION=" + cboDocLocation.Text);
                        return;
                }

                Document openedDocument = _documentManager.OpenDocument(DocumentPath_Target, DocPasswordRead, DocPasswordWrite);
                Globals.ThisDocument.ActiveDocument = openedDocument;
                Microsoft.Office.Interop.Word.Application application = openedDocument.Application;
                DocumentInfo documentInfo = new DocumentInfo();
                documentInfo = _documentManager.Document_SelectQPSObject(DocumentName_Target, documentInfo, blnRefresh: true);
                docTarget = openedDocument.Application.Documents[DocumentName_Target];
                DocPasswordProtect = documentInfo.GetClsDocument.DocPasswordProtect;
                CommonService.DocumentProtection_Off(docTarget, DocPasswordProtect);
                if (docTarget.TrackRevisions == true)
                {
                    docTarget.TrackRevisions = false;
                }
                docTarget.Fields.Update();

                Globals.ThisDocument.FileOpenDocumentName(DocumentName_Target);

                if (!string.IsNullOrEmpty(DocumentName_Target))
                {
                    application.Documents[DocumentName_Target].Activate();
                }

                strExpectedTemplateFullName = CommonService.Get_Template_Name(docTarget.Application, TemplateNameFormat.tnfNameAndPath);
                //Need to check:
                //strDocumentTemplateFullName = ActiveDocument.AttachedTemplate.FullName
                strDocumentTemplateFullName = application.ActiveDocument.get_AttachedTemplate().FullName;
                if (strDocumentTemplateFullName != strExpectedTemplateFullName)
                {
                    application.ActiveDocument.set_AttachedTemplate(strExpectedTemplateFullName);
                    wtpTemplate = application.ActiveDocument.get_AttachedTemplate();
                    if (wtpTemplate.FullName != strExpectedTemplateFullName)
                    {
                        CommonService.Error_Handler("0", "Failed to attach Template '" + strExpectedTemplateFullName + "'", MyCompName, "cmdFinish_Click");
                    }
                }
                RunEditProperty((qRequest)CommonService.qRequest_DBDownLoad);
                foreach (Range aStory in application.ActiveDocument.StoryRanges)
                {
                    foreach (Field aField in aStory.Fields)
                    {
                        aField.Update();
                    }
                }
                Globals.ThisDocument.CommandBar_Build(_tblOpUnit.OpUnitType, _opUnitDocType.DocTypeAbb, _tblOpUnit.OpUnitID, _opUnitDocType.DocTypeID);

                if (docTarget.ProtectionType == WdProtectionType.wdAllowOnlyFormFields)
                {
                    CommonService.DocumentProtection_Off(docTarget, DocPasswordProtect);
                }

                CommonService.DocumentProtection_On(docTarget, DocPasswordProtect);
                this.Close();
                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                this.Close();
                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorNormal;

                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "cmdFinish_Click");
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, ex.Source, "cmdFinish_Click");
                }
            }
            finally
            {
                CommonService.CleanUpBackgroundWordProcesses();
            }

        }


        private bool ValidateUserCreds()
        {
            string MyProcName = "ValidateUserCreds";

            if (string.IsNullOrEmpty(txtEDMSUserID.Text.Trim()))
            {
                txtEDMSUserID.Focus();
                MessageBox.Show("UserID is required", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);  // NewCode
                return false;
            }
            if (string.IsNullOrEmpty(txtEDMSPassword.Text))
            {
                txtEDMSPassword.Focus();
                MessageBox.Show("Password is required", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);  // NewCode
                return false;
            }
            int failedLogin = 0; // 2301
            if (Ldapauth(txtEDMSUserID.Text.Trim(), txtEDMSPassword.Text))
            {
                return true;
            }
            else
            {
                txtEDMSUserID.Text = Environment.UserName.ToUpper();
                txtEDMSPassword.Text = "";
                failedLogin++;
                if (failedLogin > 5)
                {
                    MessageBox.Show("Authentication Request denied due to multiple unsuccessful login attempts.\n Please unlock your NT account using https://password.pfizer.com/ to access the application.", "Authentication Request Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                MessageBox.Show("Invalid Login Credentials!!");
                txtEDMSUserID.Text = Environment.UserName.ToUpper();
                txtEDMSUserID.Enabled = false;
                return false;
            }
            // return true;  // 2301
        }


       
        public bool Ldapauth(string strUserID, string strPwd)
        {
            string TargetServer = ConfigurationManager.AppSettings["LDAPServer"].ToString();
            int ldapPort = Convert.ToInt32(ConfigurationManager.AppSettings["LDAPPort"]);
            string ldapRoot = ConfigurationManager.AppSettings["LDAPRoot"];
            string LdapUser = ConfigurationManager.AppSettings["SvcAccUsr"];
            string srvAccPwd = ConfigurationManager.AppSettings["SvcAccPwd"];
            string ldapMaxRes = ConfigurationManager.AppSettings["LDAPMaxRes"];

            const string processName = "Ldapauth";
            bool isContinueProcessing = false;
            string userDN = string.Empty;

            try
            {

                var creds = new NetworkCredential(LdapUser, srvAccPwd);
                var identifier = new LdapDirectoryIdentifier(TargetServer, ldapPort);
                using (LdapConnection connection = new LdapConnection(identifier, creds))
                {
                    connection.AuthType = AuthType.Basic;
                    connection.SessionOptions.ProtocolVersion = 3;
                    connection.SessionOptions.SecureSocketLayer = true;
                    connection.SessionOptions.VerifyServerCertificate += (sender, certificate) =>
                    {
                        return certificate.Subject.Contains($"CN={TargetServer},");
                    };

                    // Request mail attribute explicitly
                    string searchFilter = $"(&(objectClass=person)(cn={strUserID}))";
                    var searchRequest = new SearchRequest(ldapRoot, searchFilter, SearchScope.Subtree, new string[] { "mail" });  // <-- Added mail attribute
                    SearchResponse searchResponse = (SearchResponse)connection.SendRequest(searchRequest);
                    if (searchResponse.Entries.Count > 0)
                    {
                        var entry = searchResponse.Entries[0];
                        // Get Distinguished Name
                        userDN = entry.DistinguishedName;
                        // Get Email
                        if (entry.Attributes["mail"] != null)
                        {
                            mailid = entry.Attributes["mail"][0].ToString();
                        }
                        // Bind using user credentials
                        connection.Bind(new NetworkCredential(userDN, strPwd));
                        isContinueProcessing = true;
                    }
                    else
                    {
                        isContinueProcessing = false;
                    }
                }
            }
            catch (Exception ex)
            {
                string argpErrNum = ex.HResult.ToString();
                string argpErrDesc = ex.Message.ToString() + " Code Line: " + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' '));
                CommonService.Error_Handler(argpErrNum, argpErrDesc, "LDAP Authentication Failure", processName, false);
            }
            return isContinueProcessing;
        }

        private async void File_CopyVers()
        {
            Document currentActiveDocument = Globals.ThisDocument.ActiveDocument;

            try
            {
                string MyProcName = "File_CopyVers";
                //Dim objHourGlass As clsHourGlass;
                string CommonFilePath;
                int DocIndex;
                long DocID_Source;
                string DocumentName_Source;
                string DocumentPath_Source;
                string DocPasswordRead;
                string DocPasswordWrite;
                string DocPasswordProtect;
                string DocumentName_Target = string.Empty;
                string DocumentPath_Target;
                int SPParamCount;
                bool ReProtect;
                Document docTarget;
                //Dim aStory As Range;
                //Dim aField As Field;
                Template wtpTemplate;
                string strDocumentTemplateFullName;
                string strExpectedTemplateFullName;
                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorWait;
                currentActiveDocument.Application.StatusBar = " ";

                if (!FreeTextValidate(txtVariantTextCopy.Text))
                {
                    MessageBox.Show("INVALID CHARACTERS IN VARIANT TEXT" + Environment.NewLine + Environment.NewLine + "Allowable characters in Document Names are:" + Environment.NewLine + Environment.NewLine + "     A-Z" + Environment.NewLine + "     a-z" + Environment.NewLine + "     0-9" + Environment.NewLine + "     Hyphen: -" + Environment.NewLine + "     Underscore: _" + Environment.NewLine + "     Period: ." + Environment.NewLine + "     or Space", "MBR Authoring - Incompatible Quality DOCS Characters", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtVariantTextCopy.Focus();
                    return;
                }
                bool isValidUser = false;
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                //QDOCS service
                var service = new QDOCSService(vaultDns, apiVersion);
                switch (cboDocLocation.Text)
                {
                    case APPROVED:
                        if (!ValidateUserCreds()) return;
                        string strPrevDir = Directory.GetCurrentDirectory(); // required?
                        string strCurrDir = currentActiveDocument.get_AttachedTemplate().Path;
                        Directory.SetCurrentDirectory(strCurrDir);                      
                        var userCheckResponse = await service.AuthenticateAsync(username, password);
                        //mailid = "Lynette.Radoe@pfizer.com";//Authentication using username, password using QDOCS 
                        var userMailId = await service.ValidateUserByEmailAsync(mailid); //Validating using the user mail Id
                        if (userCheckResponse.IsSuccess && userMailId.IsSuccess)
                        //if (userCheckResponse.IsSuccess)
                        {
                            isValidUser = true;
                        }
                        else
                        {
                            isValidUser = false;
                            MessageBox.Show("Status Code: " + userCheckResponse.StatusCode + "\n" + (userCheckResponse.ErrorMessage != null ? "Error: " + userCheckResponse.ErrorMessage : " ") + "\n" + (userCheckResponse.ErrorDetails != null ? "Error Details: " + userCheckResponse.ErrorDetails : " "), "FAILED to connect to Quality DOCS - COPY process terminated", MessageBoxButtons.OK);
                            return;
                        }
                        Directory.SetCurrentDirectory(strPrevDir); // required??
                        break;
                    default:
                        break;
                }

                OpUnitDocType opUnitDocType = (OpUnitDocType)cboDocTypeCopy.SelectedItem;
                CommonFilePath = opUnitDocType.DocumentPath;

                //Not Req
                //if (ModChangeHistory.gstrAppVersionMajor == "3") // Not req
                //{
                //    CommonFilePath = _formManager.AppStatus_Select("DocumentPath") + "\\" + cboOpUnitCopy.Text;
                //}
                //else
                //{
                //    CommonFilePath = opUnitDocType.DocumentPath;
                //}

                CommonFilePath = opUnitDocType.DocumentPath;

                DocumentDetails docDetails = (DocumentDetails)cboDocNameCopy.SelectedItem;
                DocID_Source = docDetails.DocID;
                DocumentName_Source = docDetails.DocumentName; //.Column(qdcColDocNameBase, DocIndex)
                DocPasswordRead = docDetails.DocPasswordRead;
                DocPasswordWrite = docDetails.DocPasswordWrite;
                DocumentPath_Source = CommonFilePath + "\\" + DocumentName_Source;
                currentActiveDocument.Application.StatusBar = "Copying to File " + lblFileNameCopy.Text;

                DocumentResponse docCheckResponse = null;
                bool isValidDoc = false;
                string srcDocName, srcDocExtension = string.Empty;

                switch (cboDocLocation.Text)
                {
                    case APPROVED:
                        if (isValidUser)
                        {
                            //(srcDocName, srcDocExtension) = SplitFileName(DocumentName_Source);

                            //docCheckResponse = await IsValidPDOCSDocument(pdocsService, DocumentName_Source);
                            string docNameWithoutExt = Path.GetFileNameWithoutExtension(DocumentName_Source);

                            docCheckResponse = await service.GetDocumentIdByNameAsync(docNameWithoutExt);
                            if (docCheckResponse.IsSuccess)
                            {
                                isValidDoc = true;
                            }
                            else
                            {
                                isValidDoc = false;

                                MessageBox.Show("Document '" + DocumentName_Source + "' not found", MyProcName, MessageBoxButtons.OK);
                                MessageBox.Show((docCheckResponse.ErrorMessage != null ? "Error: " + DocumentName_Source + " - " + docCheckResponse.ErrorMessage : " ") + "\n" + (docCheckResponse.ErrorDetails != null ? "Error Details: " + docCheckResponse.ErrorDetails : " "), "FAILED to connect to Quality DOCS - COPY process terminated", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
                        }
                        break;
                    default:
                        break;
                }
                var resultSet = _formManager.File_CopyVers(DocID_Source, txtVariantTextCopy.Text);

                string errorReturn = resultSet.ErrorReturn;
                if (errorReturn != "OK") { MessageBox.Show(errorReturn); return; }
                DocumentName_Target = resultSet.DocumentReturn;

                //Dictionary<string, object> item2 = resultSet.Item2;
                //if (item2.ContainsKey("@ErrorReturn"))
                //{
                //    string errorReturn = item2["@ErrorReturn"].ToString();
                //    if (errorReturn != "OK") { MessageBox.Show(errorReturn); return; }
                //}
                //if (item2.ContainsKey("@DocumentReturn"))
                //{
                //    DocumentName_Target = item2["@DocumentReturn"].ToString();
                //}

                if (DocumentName_Target != lblFileNameCopy.Text)
                {
                    MessageBox.Show("WARNING: Expecting FileName: " + Environment.NewLine + Environment.NewLine + "\t" + lblFileNameCopy.Text + Environment.NewLine + Environment.NewLine + "Generated FileName:" + Environment.NewLine + Environment.NewLine + "\t" + DocumentName_Target, MyCompName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                DocumentPath_Target = CommonFilePath + "\\" + DocumentName_Target;
                switch (cboDocLocation.Text)
                {
                    case APPROVED:
                        if (isValidDoc)
                        {
                            try
                            {                                
                                string docObjectId = docCheckResponse.DocumentId;
                                await service.DownloadDocumentByIdAsync(docObjectId, DocumentPath_Target); //Exporting the document to path
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("COPY Terminating, due to error in Quality DOCS Processing", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                MessageBox.Show("Error: Due to error in Quality DOCS Processing" + "\n" + "Error Details: " + ex.Message, "FAILED to connect to Quality DOCS - COPY process terminated", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                        break;
                    case AUTHORING:
                    case LOCKED:
                    case BLOCKED:
                        if (File.Exists(DocumentPath_Source))
                        {
                            File.Copy(DocumentPath_Source, DocumentPath_Target, overwrite: false);
                        }
                        break;
                    default:
                        MessageBox.Show("INVALID DOCLOCATION=" + cboDocLocation.Text);
                        break;
                }

                Document openedDocument = _documentManager.OpenDocument(DocumentPath_Target, DocPasswordRead, DocPasswordWrite);
                Globals.ThisDocument.ActiveDocument = openedDocument;
                Microsoft.Office.Interop.Word.Application application = openedDocument.Application;
                DocumentInfo documentInfo = new DocumentInfo();
                documentInfo = _documentManager.Document_SelectQPSObject(DocumentName_Target, documentInfo, blnRefresh: true);
                docTarget = openedDocument.Application.Documents[DocumentName_Target];
                DocPasswordProtect = documentInfo.GetClsDocument.DocPasswordProtect;
                CommonService.DocumentProtection_Off(docTarget, DocPasswordProtect);
                if (docTarget.TrackRevisions)
                {
                    docTarget.TrackRevisions = false;
                }
                docTarget.Fields.Update();
                Globals.ThisDocument.FileOpenDocumentName(DocumentName_Target);

                if (!string.IsNullOrEmpty(DocumentName_Target))
                {
                    openedDocument.Application.Documents[DocumentName_Target].Activate();
                }

                strExpectedTemplateFullName = CommonService.Get_Template_Name(docTarget.Application, TemplateNameFormat.tnfNameAndPath);
                //Need to check:
                //strDocumentTemplateFullName = ActiveDocument.AttachedTemplate.FullName
                strDocumentTemplateFullName = application.ActiveDocument.get_AttachedTemplate().FullName;
                if (strDocumentTemplateFullName != strExpectedTemplateFullName)
                {
                    application.ActiveDocument.set_AttachedTemplate(strExpectedTemplateFullName);
                    wtpTemplate = application.ActiveDocument.get_AttachedTemplate();
                    if (wtpTemplate.FullName != strExpectedTemplateFullName)
                    {
                        CommonService.Error_Handler("0", "Failed to attach Template '" + strExpectedTemplateFullName + "'", MyCompName, "cmdFinish_Click");
                    }
                }
                RunEditProperty((qRequest)CommonService.qRequest_DBDownLoad);
                foreach (Range aStory in application.ActiveDocument.StoryRanges)
                {
                    foreach (Field aField in aStory.Fields)
                    {
                        aField.Update();
                    }
                }

                Globals.ThisDocument.CommandBar_Build(_tblOpUnit.OpUnitType, _opUnitDocType.DocTypeAbb, _tblOpUnit.OpUnitID, _opUnitDocType.DocTypeID);

                if (docTarget.ProtectionType == WdProtectionType.wdAllowOnlyFormFields)
                {
                    CommonService.DocumentProtection_Off(docTarget, DocPasswordProtect);
                }
                if (docTarget.TrackRevisions == false)
                {
                    docTarget.TrackRevisions = true;

                }
                CommonService.DocumentProtection_On(docTarget, DocPasswordProtect);
                this.Close();
                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                this.Close();
                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorNormal;

                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "cmdFinish_Click");
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "cmdFinish_Click");
                }
            }
            finally
            {
                CommonService.CleanUpBackgroundWordProcesses();
            }
        }

        private qReturn RunEditProperty(qRequest qrqRequest, QWriteWhat qwwWriteWhat = 0, bool pblnReplaceTable = false)
        {
            qReturn runProperty = qReturn.qReturn_Cancel;
            FormEditProperty editProperty = new FormEditProperty();
            Document targetDocument = Globals.ThisDocument.ActiveDocument;
            editProperty.SetTargetDocumentName(targetDocument?.Name, targetDocument?.FullName);

            StepVarManager.Form_Push(this);
            editProperty.QLet(this.Top, this.Left + mpgWizardPage.Left, mqpsEditPropertyParms);
            switch (qrqRequest)
            {
                case qRequest.qRequest_DBDownLoad:
                    runProperty = editProperty.QDBDownload();
                    break;
            }

            if (CommonService.QReturn_QGet(MyCompName, runProperty))
            {
                editProperty.QGet(mlngMyTop, mlngMyLeft, mqpsEditPropertyParms);
                mlngMyLeft = mlngMyLeft - mpgWizardPage.Left;
            }

            editProperty.QClose(); //070125
            StepVarManager.Form_Pop(this);
            return runProperty;
        }

        public bool FreeTextValidate(string inputText)
        {
            if (string.IsNullOrEmpty(inputText))
            {
                return true;
            }

            return _regex.IsMatch(inputText);
        }

        private void cboDocumentSupercedes__LOAD()
        {
            bool blnIncludeLocked;
            bool blnIncludeAuthoring;
            bool blnIncludeApproved;
            bool blnIncludeArchived;
            blnIncludeLocked = true;
            blnIncludeAuthoring = true;
            blnIncludeApproved = true;
            blnIncludeArchived = true;
            var files = _formManager.Fill_Document(mintOpUnitIDCreate, mintDocTypeIDCreate, blnIncludeLocked, blnIncludeAuthoring, blnIncludeApproved, blnIncludeArchived, false);
            cboDocumentSupercedes.DisplayMember = "DocumentName";
            cboDocumentSupercedes.ValueMember = "DocID";
            cboDocumentSupercedes.DataSource = files;
            cboDocumentSupercedes.Items.Insert(0, string.Empty);
            cboDocumentSupercedes.SelectedIndex = 0;
            var selectedDocTypeCreate = (OpUnitDocType)cboDocTypeCreate.SelectedItem;
            SetToolTip(cboDocumentSupercedes, "cboDocumentSupercedes: OpUnitID=" + mintOpUnitIDCreate + " OpUnitType=" + selectedDocTypeCreate.DocTypeID + " DocTypeID=" + mintDocTypeIDCreate + " InclLocked=" + blnIncludeLocked);
        }

        private void File_Create()
        {
            string MyProcName = "File_Create";
            Document currentActiveDocument = Globals.ThisDocument.ActiveDocument;//currentDocument;
            var wordApp = currentActiveDocument.Application;
            try
            {
                //string MyProcName = "File_Create";
                //Dim objHourGlass As clsHourGlass;
                long lngWCID;
                string strFilePath;
                string strTemplateFullName = string.Empty;
                long lngDocTypeID;
                long lngRecTypeID;
                int intResponse;
                //Dim objCmd As New ADODB.Command
                long lngSPMat;
                string strSPFree;
                string strDocumentTitle;
                long intSPDocID;
                string strSPDocName = string.Empty;
                string strSPDocPasswordWrite = "";
                string strSPDocPasswordRead = "";
                int intSPParamCount;



                if (wordApp.ActiveWindow.View.SplitSpecial == WdSpecialPane.wdPaneNone)
                {
                    wordApp.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
                    wordApp.ActiveWindow.ActivePane.View.Zoom.PageFit = WdPageFit.wdPageFitBestFit;
                }
                else
                {
                    wordApp.ActiveWindow.View.Type = WdViewType.wdPrintView;
                    wordApp.ActiveWindow.View.Zoom.PageFit = WdPageFit.wdPageFitBestFit;
                }
                wordApp.ActiveWindow.View.ShowBookmarks = true; // Todo: Check if needed on file open 
                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorWait;
                currentActiveDocument.Application.StatusBar = " ";

                lngRecTypeID = mintDocTypeIDCreate;
                lngWCID = mlngWorkCenterID;

                OpUnitDocType opUnitDocType = (OpUnitDocType)cboDocTypeCreate.SelectedItem;
                strFilePath = opUnitDocType.DocumentPath;
                currentActiveDocument.Application.StatusBar = "Creating File " + lblFileNameCreate.Text;
                mlngSuperDocID = 0;
                if (cboDocumentSupercedes.Visible)
                {
                    DocumentDetails docDetails = (DocumentDetails)cboDocumentSupercedes.SelectedItem;
                    // Todo not required it seems since docDetails.DocID is already an int.
                    //if (IsNumeric(cboDocumentSupercedes.Column(qdcColDocID, cboDocumentSupercedes.ListIndex)) {
                    mlngSuperDocID = docDetails.DocID;
                    //}
                }
                //cmdFinishOverwrite:
                strSPFree = txtFreeTextCreate.Text;
                if (mblnUseDocBaseNameCreate)
                {
                    DocBaseNameDetails docBaseNameDetails = (DocBaseNameDetails)cboDocBaseNameCreate.SelectedItem;
                    strDocumentTitle = docBaseNameDetails.DocumentTitle;
                }
                else
                {
                    Material material = (Material)cboProductCreate.SelectedItem;
                    strDocumentTitle = material.MaterialTitle;
                }
                long workCenterID = 0;
                if (mblnUseWorkCenterCreate)
                {
                    workCenterID = lngWCID;
                }
                long materialID = 0;
                if (mblnUseProductCreate)
                {
                    materialID = mlngMaterialIDCreate;
                }

                if (!ValidateParams(txtFreeTextCreate.Text, "FreeText")) return;
                string freeText = string.Empty;
                //if (lblFreeTextCreate.Text == FREETEXT)
                //{
                //    freeText = txtFreeTextCreate.Text;
                //}
                //else if (lblFreeTextCreate.Text == MODULEID)
                //{
                //    freeText = cboDocBaseNameCreate.Text;
                //}

                var resultSet = _formManager.File_Create(mintOpUnitIDCreate, lngRecTypeID, strDocumentTitle, workCenterID, materialID, txtFreeTextCreate.Text, cboDocBaseNameCreate.Text, mlngSuperDocID);


                string errorReturn = resultSet.ErrorReturn;
                if (errorReturn != "OK") { MessageBox.Show(errorReturn); return; }
                strSPDocName = resultSet.DocumentNameReturn;
                if (strSPDocName != lblFileNameCreate.Text)
                {
                    MessageBox.Show("WARNING: Expecting FileName: " + Environment.NewLine + Environment.NewLine + "\t" + lblFileNameCreate.Text + Environment.NewLine + Environment.NewLine + "Generated FileName:" + Environment.NewLine + Environment.NewLine + "\t" + strSPDocName, MyCompName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                strSPDocPasswordWrite = CommonService.TranslatePassword(resultSet.DocPasswordWriteReturn);
                strSPDocPasswordRead = CommonService.TranslatePassword(resultSet.DocPasswordReadReturn);

                string version = wordApp.Application.Version;  // NewCode3
                int intVersion = int.Parse(version.Substring(0, version.IndexOf('.')));
                if (intVersion >= CommonService.gintWordVersion2007)
                {
                    strTemplateFullName = CommonService.Get_Template_Name(wordApp.Application, TemplateNameFormat.tnfReview) + "m";
                }
                else
                {
                    strTemplateFullName = CommonService.Get_Template_Name(wordApp.Application, TemplateNameFormat.tnfReview);
                }

                mdocNewDocument = wordApp.Documents.Add(strTemplateFullName);
                //////
                Template temp = mdocNewDocument.get_AttachedTemplate();
                string strDocumentTemplateFullName = temp.FullName;

                if (strDocumentTemplateFullName != CommonService.templateFullName)
                {
                    mdocNewDocument.set_AttachedTemplate(CommonService.templateFullName);
                }

                CommonService.Style_Copy_All(mdocNewDocument, _documentManager);

                wordApp.ChangeFileOpenDirectory(strFilePath);
                Clean();
                mdocNewDocument.SaveAs(FileName: strSPDocName, FileFormat: WdSaveFormat.wdFormatDocument, LockComments: false, Password: strSPDocPasswordRead, WritePassword: strSPDocPasswordWrite, AddToRecentFiles: true, ReadOnlyRecommended: false, EmbedTrueTypeFonts: false, SaveNativePictureFormat: false, SaveFormsData: false, SaveAsAOCELetter: false);
                wordApp.DocumentBeforeClose += new DocumentManager().Document_BeforeClose; // 2301
                wordApp.DocumentBeforeSave += new DocumentManager().Document_BeforeSave; // 2301

                _documentManager.Document_SelectQPSObject(strSPDocName, null, true);
                Globals.ThisDocument.FileOpenDocumentName(strSPDocName);

                if (!string.IsNullOrEmpty(strSPDocName))
                {
                    mdocNewDocument.Activate();
                }
                Globals.ThisDocument.ActiveDocument = mdocNewDocument;

                FormIntroPagesMBR formIntroPagesMBR = new FormIntroPagesMBR();// (mdocNewDocument);
                formIntroPagesMBR.Show();

                Globals.ThisDocument.CommandBar_Build(_tblOpUnit.OpUnitType, _opUnitDocType.DocTypeAbb, _tblOpUnit.OpUnitID, _opUnitDocType.DocTypeID);
                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorNormal;
                this.Close();


            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                this.Close();
                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorNormal;
                if (ex.Data["errNum"] != null)  // NewCode1
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    MessageBox.Show("The following Error occurred " + ex.Message + ". Please contact MBR Authoring support.");
                }
            }
            finally
            {
                CommonService.CleanUpBackgroundWordProcesses();
            }
        }

        public void Clean()//(Document doc)
        {
            string MyProcName = "Clean";
            if (mdocNewDocument.ProtectionType != WdProtectionType.wdNoProtection)
            {
                mdocNewDocument.Unprotect("USCG1790");
            }

            // Select the entire content of the document
            mdocNewDocument.Application.Selection.WholeStory();

            // Delete one character at the current selection point
            mdocNewDocument.Application.Selection.Delete(Unit: WdUnits.wdCharacter, Count: 1);
            if (mdocNewDocument.ActiveWindow.View.SplitSpecial != WdSpecialPane.wdPaneNone)
            {
                CommonService.Error_Handler("913", "SplitSpecial=" + mdocNewDocument.Application.ActiveWindow.View.SplitSpecial, MyProcName, MyProcName);
                mdocNewDocument.Application.ActiveWindow.Panes[2].Close();
            }
            CommonService.ViewPage_Main(mdocNewDocument.Application);
        }

        private void Style_Copy_All()
        {
            throw new NotImplementedException();
        }

       

        //private async void PDOCS(PDOCSService _PDOCSService)
        //{
        //    try
        //    {
        //        // validate the user/service acc: User: Srvamr-pdocapid

        //        var userCheckResponse = await _PDOCSService.IsValidUser("baskag03"); // case-sensitive
        //        if (userCheckResponse.Success)
        //        {
        //            var docName = "F000012165-v21"; // doc name without extension; case-sensitive

        //            // Todo uncomment below
        //            //var docName = _PDOCSService.GetFileExtension("F000012165-v21.doc"); 
        //            //var docName= docName.FileName,
        //            var docExtension = ".doc";// docName.Extension;
        //            var docCheckResponse = await _PDOCSService.IsValidDocument(docName);
        //            if (docCheckResponse.Success)
        //            {
        //                string docObjectId = docCheckResponse.Response;
        //                string newFileName = "F000012165-v22";
        //                string newFileNameWithExtension = newFileName + docExtension;
        //                string filePath = @"\\AMRMCPW040\MBRDocs\" + newFileNameWithExtension;
        //                //string filePath = @"\\AMRMCPW040\MBRDocs\F000012165-v22.doc"; // Give the same file extension: .doc,.docx,...
        //                await _PDOCSService.ExportDocument(docObjectId, filePath);

        //            }
        //            else
        //            {
        //                MessageBox.Show((docCheckResponse.ErrorMessage != null ? "Error: " + docName + " - " + docCheckResponse.ErrorMessage : " ") + "\n" + (docCheckResponse.ErrorDetails != null ? "Error Details: " + docCheckResponse.ErrorDetails : " "), "FAILED to connect to PDOCS - COPY process terminated", MessageBoxButtons.OK);
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("Status Code: " + userCheckResponse.StatusCode + "\n" + (userCheckResponse.ErrorMessage != null ? "Error: " + userCheckResponse.ErrorMessage : " ") + "\n" + (userCheckResponse.ErrorDetails != null ? "Error Details: " + userCheckResponse.ErrorDetails : " "), "FAILED to connect to PDOCS - COPY process terminated", MessageBoxButtons.OK);
        //        }
        //    }

        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error: Due to error in PDOCS Processing" + "\n" + "Error Details: " + ex.Message, "FAILED to connect to PDOCS - COPY process terminated", MessageBoxButtons.OK);
        //    }
        //}
        //private async void PDOCS(PDOCSService _PDOCSService)
        //{
        //    try
        //    {
        //        // validate the user/service acc: User: Srvamr-pdocapid

        //        var userCheckResponse = await _PDOCSService.IsValidUser("baskag03"); // case-sensitive
        //        if (userCheckResponse.Success)
        //        {
        //            var docName = "F000012165-v21"; // doc name without extension; case-sensitive

        //            // Todo uncomment below
        //            //var docName = _PDOCSService.GetFileExtension("F000012165-v21.doc"); 
        //            //var docName= docName.FileName,
        //            var docExtension = ".doc";// docName.Extension;
        //            var docCheckResponse = await _PDOCSService.IsValidDocument(docName);
        //            if (docCheckResponse.Success)
        //            {
        //                string docObjectId = docCheckResponse.Response;
        //                string newFileName = "F000012165-v22";
        //                string newFileNameWithExtension = newFileName + docExtension;
        //                string filePath = @"\\AMRMCPW040\MBRDocs\" + newFileNameWithExtension;
        //                //string filePath = @"\\AMRMCPW040\MBRDocs\F000012165-v22.doc"; // Give the same file extension: .doc,.docx,...
        //                await _PDOCSService.ExportDocument(docObjectId, filePath);

        //            }
        //            else
        //            {
        //                MessageBox.Show((docCheckResponse.ErrorMessage != null ? "Error: " + docName + " - " + docCheckResponse.ErrorMessage : " ") + "\n" + (docCheckResponse.ErrorDetails != null ? "Error Details: " + docCheckResponse.ErrorDetails : " "), "FAILED to connect to PDOCS - COPY process terminated", MessageBoxButtons.OK);
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("Status Code: " + userCheckResponse.StatusCode + "\n" + (userCheckResponse.ErrorMessage != null ? "Error: " + userCheckResponse.ErrorMessage : " ") + "\n" + (userCheckResponse.ErrorDetails != null ? "Error Details: " + userCheckResponse.ErrorDetails : " "), "FAILED to connect to PDOCS - COPY process terminated", MessageBoxButtons.OK);
        //        }
        //    }

        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error: Due to error in PDOCS Processing" + "\n" + "Error Details: " + ex.Message, "FAILED to connect to PDOCS - COPY process terminated", MessageBoxButtons.OK);
        //    }
        //}

        public void Style_Fix(Document doc)
        {
            string strProtectionPassword = string.Empty;
            if (doc.ProtectionType == WdProtectionType.wdNoProtection)
            {
                strProtectionPassword = string.Empty;
            }
            else
            {
                //strProtectionPassword = Document_SelectPasswordProtect(doc.Name);
                CommonService.DocumentProtection_Off(doc, strProtectionPassword);
            }
            //Style_Copy_One(StyleName:= gstrStyle_MBRSteps);
            if (string.IsNullOrEmpty(strProtectionPassword))
            {
                CommonService.DocumentProtection_On(doc, strProtectionPassword);
            }
        }

        //public static (string baseName, string extension) SplitFileName(string docName)
        //{
        //    if (string.IsNullOrWhiteSpace(docName))
        //    {
        //        throw new ArgumentException("Document name cannot be null or empty.", nameof(docName));
        //    }

        //    string baseName = Path.GetFileNameWithoutExtension(docName);

        //    string extension = Path.GetExtension(docName);

        //    return (baseName, extension);
        //}
        private void cboDocBaseNameCreate_KeyDown(object sender, KeyEventArgs e)
        {
            //Todo:
            //AutoCompleteCombo(cboDocBaseNameCreate, KeyCode);
        }

        private void cboDocBaseNameCopy_KeyDown(object sender, KeyEventArgs e)
        {
            //Todo:
            // AutoCompleteCombo(cboDocNameCopy, KeyCode)
        }

        //protected bool Ldapauth(string sUserName, string sPassword)
        //{
        //    Cursor.Current = Cursors.WaitCursor;
        //    bool sReturnValue = false;
        //    string userDN = string.Empty;
        //    try
        //    {
        //        String[] ldapAttrs = new String[100];
        //        // this flag lets us know whether to keep going.  this way, we can
        //        // save our return until the end
        //        bool isContinueProcessing = true;

        //        string ldapType = ConfigurationManager.AppSettings["LDAPType"];
        //        string ldapServer = ConfigurationManager.AppSettings["LDAPServer"];
        //        string ldapPort = ConfigurationManager.AppSettings["LDAPPort"];
        //        string ldapRoot = ConfigurationManager.AppSettings["LDAPRoot"];
        //        string srvAccUser = ConfigurationManager.AppSettings["SvcAccUsr"];
        //        string srvAccPwd = ConfigurationManager.AppSettings["SvcAccPwd"];
        //        string ldapMaxRes = ConfigurationManager.AppSettings["LDAPMaxRes"];



        //        //LdapConnection conn = new LdapConnection();
        //        System.DirectoryServices.Protocols.LdapConnection conn = new System.DirectoryServices.Protocols.LdapConnection(new LdapDirectoryIdentifier(ldapServer, int.Parse(ldapPort)));
        //        System.DirectoryServices.Protocols.LdapConnection conn2 = new System.DirectoryServices.Protocols.LdapConnection(new LdapDirectoryIdentifier(ldapServer, int.Parse(ldapPort)));
        //        //LdapConnection conn2 = new LdapConnection();

        //        //bind to directory using service account
        //        /*****************************************/
        //        if (isContinueProcessing)
        //        {

        //            // connect to the directory
        //            //conn.Connect(ldapServer, int.Parse(ldapPort));
        //            conn.AuthType = AuthType.Basic;
        //            // bind
        //            try
        //            {
        //                conn.Bind(new NetworkCredential(srvAccUser, srvAccPwd));
        //                conn.SessionOptions.SecureSocketLayer = true;
        //                conn.SessionOptions.VerifyServerCertificate = new VerifyServerCertificateCallback((con, cert) => true);
        //                conn.Bind();
        //                //conn.Dispose();
        //            }
        //            catch (Exception ex)
        //            {
        //                isContinueProcessing = false;
        //            }

        //            //if (!conn.Bound)
        //            {
        //                //sReturnValue = "User not found in directory;";
        //                //isContinueProcessing = false;
        //            }
        //        }
        //        /*****************************************/

        //        //see if user exists, fetch dn
        //        /*****************************************/
        //        if (isContinueProcessing)
        //        {
        //            // setup filter
        //            string searchFilter = String.Format("(cn={0})", sUserName);

        //            // search for the record
        //            //LdapSearchResults lsr;
        //            ILdapSearchResults lsr;
        //            if (ldapMaxRes != null && ldapMaxRes != "")
        //            {
        //                LdapSearchConstraints lsc = new LdapSearchConstraints();
        //                lsc.MaxResults = int.Parse(ldapMaxRes);
        //                //lsr = conn.Search(ldapRoot, LdapConnection.SCOPE_SUB, searchFilter.ToString(), null, false, lsc);
        //                //lsr = conn.Search(ldapRoot, LdapConnection.ScopeSub, searchFilter.ToString(), null, false, lsc);

        //            }
        //            else
        //            {
        //                //lsr = conn.Search(ldapRoot, LdapConnection.ScopeSub, searchFilter.ToString(), ldapAttrs, false);
        //                //lsr = conn.Search(ldapRoot, LdapConnection.SCOPE_SUB, searchFilter.ToString(), ldapAttrs, false);
        //            }

        //            // return object if found, otherwise null
        //            LdapEntry resultEntry = null;

        //            //if (lsr.HasMore())
        //            ////if (lsr.hasMore())
        //            //{
        //            //    //we found a match
        //            //    //resultEntry = lsr.next();
        //            //    resultEntry = lsr.Next();
        //            //    userDN = resultEntry.Dn.ToString();
        //            //    //userDN = resultEntry.DN.ToString();
        //            //}
        //            //else
        //            //{
        //            //    //user was not found
        //            //    //sReturnValue = "User not found in directory;Please check that you have entered your Pfizer network id";
        //            //    isContinueProcessing = false;
        //            //}
        //        }
        //        /*****************************************/

        //        //bind again as this user
        //        /*****************************************/
        //        if (isContinueProcessing)
        //        {
        //            //conn2.Connect(ldapServer, int.Parse(ldapPort));

        //            // bind
        //            //try { conn2.Bind(userDN, sPassword); }
        //            //catch (Exception ex)
        //            //{

        //            //}

        //            //if (!conn2.Bound)
        //            //{
        //            //    //bind failed - must be incorrect password
        //            //    //sReturnValue = "Incorrect password; Please use your Pfizer network password.";
        //            //    isContinueProcessing = false;
        //            //}
        //            //else
        //            //{
        //            //    sReturnValue = true;
        //            //    // don't return here, clean up first
        //            //}
        //        }
        //        /*****************************************/

        //        //clean up and go home
        //        /*****************************************/
        //        //conn.Disconnect();
        //        //conn2.Disconnect();
        //        conn.Dispose();
        //        conn2.Dispose();


        //    }
        //    catch (Exception ex)
        //    {
        //        sReturnValue = false;
        //    }
        //    finally
        //    {
        //        Cursor.Current = Cursors.Default;
        //    }
        //    return sReturnValue;

        //}

        private void cboProductCreate_KeyDown(object sender, KeyEventArgs e)
        {
            return;
            //ComboBox_KeyDown(sender, e, typeof(Material));
            ComboBox comboBox = sender as ComboBox;

            switch (e.KeyCode)
            {
                case Keys.Down:
                    if (comboBox.SelectedIndex < comboBox.Items.Count - 1)
                    {
                        comboBox.SelectedIndex++;
                    }
                    e.Handled = true;
                    break;

                case Keys.Up:
                    if (comboBox.SelectedIndex > 0)
                    {
                        comboBox.SelectedIndex--;
                    }
                    e.Handled = true;
                    break;

                case Keys.Enter:
                case Keys.Tab:
                    // Handle Enter or Tab key
                    SendKeys.Send("{TAB}");
                    e.Handled = true;
                    break;

                case Keys.Escape:
                    comboBox.Text = string.Empty; // Clear text on Escape
                    e.Handled = true;
                    break;

                default:
                    // Handle character input
                    if (char.IsControl((char)e.KeyCode) || e.KeyCode == Keys.Shift)
                        break;

                    // Perform auto-completion
                    AutoCompleteCombo(comboBox, e.KeyCode);
                    break;
            }
        }
        private void AutoCompleteCombo(ComboBox comboBox, Keys keyCode)
        {
            string input = comboBox.Text;
            if (string.IsNullOrEmpty(input)) return;

            // Filter items based on current input
            //var filteredItems = materials.Where(item => item.MaterialDisplay.StartsWith(input, StringComparison.OrdinalIgnoreCase)).ToArray();

            var filteredItems1 = materials.Where(item => item.MaterialDisplay.StartsWith(input, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (filteredItems1.Count > 0)
            {
                var orgDataSrc = comboBox.DataSource;
                comboBox.DataSource = filteredItems1;
                //comboBox.Items.AddRange(filteredItems1.ToArray());
                comboBox.DisplayMember = "MaterialDisplay";

                comboBox.DroppedDown = true;
                comboBox.SelectionStart = input.Length;
                comboBox.SelectionLength = filteredItems1[0].MaterialDisplay.Length - input.Length;
                comboBox.DataSource = orgDataSrc;
                comboBox.Text = filteredItems1[0].MaterialDisplay;
            }
            else
            {
                comboBox.DroppedDown = false; // No matching items
            }
        }
        //cboDocBaseNameCreate_KeyDown, cboDocNameCopy_KeyDown, cboProductCreate_KeyDown, cboDocumentSupercedes__LOAD,  txtFreeTextCreate_KeyPress, txtVariantTextCopy_KeyPress, RunEditProperty
    }
}
