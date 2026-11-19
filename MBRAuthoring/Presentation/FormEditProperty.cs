using MBRAuthoring.Business;
using MBRAuthoring.DataAccess;
using MBRAuthoring.Models;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using NaturalSort.Extension;
using System.Diagnostics;

namespace MBRAuthoring.Presentation
{
    public partial class FormEditProperty : FormBase
    {
        const string MyCompName = "FormEditProperty";
        const string MyMsgboxTitle = "MBR Edit Property";
        //performance
        private System.Windows.Forms.Timer wordKeepAliveTimer;
        protected override Panel progressBarPanel => panel1;  //060825

 

        List<string> lstPropertyMessages = new List<string>();
        System.Data.DataTable dtPropertyAll = new System.Data.DataTable();
        System.Data.DataTable dtPropertyCategory = new System.Data.DataTable();
        System.Data.DataTable dtPropertyValue = new System.Data.DataTable();
        System.Data.DataTable dtDocPropertyRequired = new System.Data.DataTable();
        System.Data.DataTable dtDocProperty = new System.Data.DataTable();
        List<string> lstPropertyCat = new List<string>();
        //List<TblLookupData> lookupDataList = new List<TblLookupData>(); //130125 added
        List<TblLookupData> propertyValueLookupDataList = new List<TblLookupData>(); //180125
        List<TblLookupData> propertyUOMLookupDataList = new List<TblLookupData>(); //180125

        bool IsModule = true;
        bool mblnModuleRecord;
        bool mblnDocTrackChgsReq;

        public string PropertyName;
        public string PropertyDescription;
        public int PropertyIndexMax;
        public string PropertyDataType;
        public string PropertyIndex;
        public string PropertyInstance;
        public string PropertyValue;
        public string PropertyUOM;
        public string PropertyDataTypeName;
        public string PropertyUsage;
        public string PropertyUsageName;
        public string PropertyLookupType;
        public string PropertyUOMLookupType;
        public int PropertyTextMaxLen;
        public string PropertyValueAsName;
        string DocIsModule;
        string DocHasModule;
        string DocHasMaterial;
        string DocHasProduct;
        string DocHasProperty;
        string DocHasWorkCenter;
        string DocHasWorkCenterAlt;
        string DocPermitsQuantitySufficient;
        int PropertyValue_ValueColumn;
        int PropertyMaterialID;
        int PropertyTable_Start;
        int PropertyTable_End;
        string PropertyValueDisplay;
        int PropertyWorkCenterID;
        bool blnDBUpdateSuccessful;
        bool PropertyDecIntWithLookup;
        public IDocumentManager _documentManager = null;
        private Word.Document activeDocument = Globals.ThisDocument.ActiveDocument;
        qRequest mqrqCalledRequest;
        qReturn mqrtQReturn;
        MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
        FormManager frmManager = new FormManager();
        CommonService objCommonService = new CommonService();
        Word.Table tabProperty;
        const float sglMargin = 5;
        float sglPropertyValue_Top;
        float sglPropertyValue_Left;
        float sglPropertyValue_Height;
        float sglPropertyValue_Width;
        float sglPropertyAdd_Top;
        float sglPropertyUpdate_Top;
        float sglPropertyDescribe_Top;
        float sglPropertyDelete_Top;
        float sglDocPropertyLabl_Top;
        float sglDocPropertyList_Top;
        float sglDocPropertyList_Hgt;

        string mstrDocName;
        string mstrDocFullName;
        int mDocID;
        string mstrProtectionPassword;
        string mstrOpUnit;
        bool mblnCalledWizard;
        int mintOpUnitID;
        string mstrOpUnitType;
        int mintDocTypeID;
        int intScann = -1; //090125
        //bool mbWizardCopyCalled = false;

        public const int qmpColMasterPropertyInstance = 0;
        public const int qmpColMasterPropertyValue = 1;
        public const int qmpColMasterPropertyDocName = 2;
        public const int qmpColMasterPropertyDocLocation = 3;
        //modified as per required abc
        public const int qdpColPropertyValueRequired = 0;
        public const int qdpColPropertyName = 1;
        public const int qdpColPropertyIndex = 2;
        public const int qdpColPropertyInstance = 3;
        public const int qdpColPropertyValue = 4;
        public const int qdpColPropertyUOM = 5;
        public const int qdpColPropertyValueDisplay = 6;
        public const int qdpColWorkCenterID = 7;
        public const int qdpColMaterialID = 8;
        public const int qdpColCount = 9;
        /*public const int qdpColPropertyName = 0;
        public const int qdpColPropertyIndex = 1;
        public const int qdpColPropertyInstance = 2;
        public const int qdpColPropertyValue = 3;
        public const int qdpColPropertyUOM = 4;
        public const int qdpColPropertyValueDisplay = 5;
        public const int qdpColPropertyValueRequired = 6;
        public const int qdpColMaterialID = 7;
        public const int qdpColWorkCenterID = 8;        
        public const int qdpColCount = 9;*/

        public const int qpyColPropertyName = 0;
        public const int qluColLookupText = 2; //090125
        public const int qdlColMaterialDisplay = 0;
        const int cintParmsValidated = 24963;
        const int cintColCountExpected = 9;
        const int cintEditPropertyRun = 19420;
        int mintParmsValidated;
        public const int qdlColMaterialDBID = 1;
        int mlngMyTop;
        int mlngMyLeft;
        bool mblnCancel;
        bool isNew = false, isUpdate = false, syncSysProp = false, isPropertyDeleted = false;
        string strPropertyValueRequired;
        public enum PropertyValidationMode
        {
            PropertyValidation_Add = 4329,
            PropertyValidation_Update = 15902,
            PropertyValidation_Apply = 9072
        }

        public struct DocProperty
        {
            public string PropertyName;
            public string PropertyValue;
        }

        public string mstrDocumentTitle
        {
            get; set;
        }

        // 
        public const string PropertyDataType_Integer = "I";
        public const string PropertyDataType_Date = "C";
        public const string PropertyDataType_Decimal = "D";
        public const string PropertyDataType_Lookup = "L";
        public const string PropertyDataType_LookupText = "K";
        public const string PropertyDataType_Text = "T";
        public const string PropertyDataType_Paragraph = "P";
        public const string PropertyDataType_Material = "M";

        public const string cChangeNum = "ChangeNum";
        public const string cLotNum = "LotNbr";
        public const string cOrderNum = "OrderNum";
        public const string cPrintDate = "PrintDate";
        public const string cPrintTime = "PrintTime";
        public const string cProductID = "ProductID";
        public const string cProductName = "ProductName";
        public const string cManufacturing_Product_Code = "Manufacturing_Product_Code";
        public const string cNDC_Code = "NDC_Code";
        public const string cProduct_Identification_Code = "Product_Identification_Code";
        public const string cRCS_Sales_Code = "RCS_Sales_Code";
        public const string cGlobal_Trade_ID_Number = "Global_Trade_ID_Number";
        //added for GTIN pannej02
        public const string cVialLabelGlobalTradeIDNumber = "VialLabelGlobalTradeIDNumber";
        public const string cModuleTableRows = "ModuleTableRows";
        public const string cModuleTitle = "ModuleTitle";
        public const string cPrintPages = "PrintPages";
        public const string cDocumentTitle = "DocumentTitle";
        public const string cRecordType = "RecordType";
        public const string cFirstLotUsage = "FirstLotUsage";
        public const string cNotApplicable = "N/A";

        public FormEditProperty()
        {
            InitializeComponent();
            dgvDocProperty.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _documentManager = new DocumentManager();
            try
            {
                SetTargetDocumentName(activeDocument?.Name, activeDocument?.FullName);
            }
            catch
            {
            }
        }

        public void SetTargetDocumentName(string documentName, string documentFullName = "")
        {
            if (!string.IsNullOrWhiteSpace(documentName))
            {
                mstrDocName = documentName;
            }

            if (!string.IsNullOrWhiteSpace(documentFullName))
            {
                mstrDocFullName = documentFullName;
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            StopWordKeepAliveTimer();
            base.OnFormClosed(e);
        }
        private void FormEditProperty_FormClosed(object sender, FormClosedEventArgs e)
        {
           
            CommonService.CleanUpBackgroundWordProcesses();

        }

        private void FormEditProperty_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            const string MyProcName = "UserForm_Initialize";
            bool blnReProtect;
            mblnCancel = false;

            try
            {
                if (!TryResolveActiveDocumentForForm())
                {
                    throw new InvalidOperationException("Unable to resolve the target document.Close the form and try again.");
                }

                InitializeTabs(tabEditProperty, btnBack, btnNext, tabEdit, tabSource, tabUnused, tabMessages);//add two buttons
                ShowTab(0);

                foreach (Control ctrl in progressBarPanel.Controls) //06082025
                {
                    if (ctrl is Label label)
                    {
                        if (label.Tag == null || label.Tag.ToString() != "Exclude")
                        {
                            label.Click += ProgressLabel_Click;
                        }
                    }
                }

                if (!TryResolveActiveDocumentForForm())
                {
                    throw new InvalidOperationException("Unable to resolve the target document.Close the form and try again.");
                }

                GetActiveDocValues(); //060125

                _documentManager.Document_Has(mstrDocName, ref DocIsModule, ref DocHasModule, ref DocHasMaterial, ref DocHasProduct, ref DocHasProperty, ref DocHasWorkCenter, ref DocHasWorkCenterAlt, ref DocPermitsQuantitySufficient);
                mblnModuleRecord = (DocIsModule == "Y");
                GetPropertyCategory();
                ddlPropertyCategory.DataSource = null;
                lstPropertyCat.Insert(0, "[Show All]");
                ddlPropertyCategory.DataSource = lstPropertyCat;
                ddlPropertyCategory.SelectedIndex = 0;


                //PropertyUnused_Load(); //  comparison
                //PropertySource_Load();

                if (activeDocument.ProtectionType == Word.WdProtectionType.wdNoProtection)
                {
                    blnReProtect = false;
                }
                else
                {
                    Document_UnProtect();
                    blnReProtect = true;
                }
                tabProperty = CommonService.Table_Find_Property(pdocCurrent: activeDocument);


                if (tabProperty == null)
                {
                    PropertyTable_Start = 0;
                    PropertyTable_End = 0;
                }
                else
                {

                    PropertyTable_Start = tabProperty.Range.Start;
                    PropertyTable_End = tabProperty.Range.End;
                }

                if (blnReProtect)
                {
                    Document_Protect();
                }
                lblUpdateMessage.Top = dgvDocProperty.Top;
                lblUpdateMessage.Left = dgvDocProperty.Left;
                lblUpdateMessage.Height = dgvDocProperty.Height;
                lblUpdateMessage.Width = dgvDocProperty.Width;

                //mblnCalledWizard = false;//070125 need to check it for copy
                //mbWizardCopyCalled = false;

                sglPropertyValue_Height = txtPropertyValue.Height;
                sglPropertyValue_Width = txtPropertyValue.Width;
                sglPropertyValue_Top = txtPropertyValue.Top;
                sglPropertyValue_Left = txtPropertyValue.Left;
                sglPropertyAdd_Top = btnPropertyAdd.Top;
                sglPropertyUpdate_Top = btnPropertyUpdate.Top;
                sglPropertyDelete_Top = btnPropertyDelete.Top;
                sglPropertyDescribe_Top = btnPropertyDelete.Top;
                sglDocPropertyLabl_Top = lblDocProperty.Top;
                sglDocPropertyList_Top = dgvDocProperty.Top;
                sglDocPropertyList_Hgt = dgvDocProperty.Height;

                MessageClear();

                chkShowAllProperties.Visible = false;
                chkShowModules.Visible = false;
                ddlPropertyCategory.Visible = true;
                lblPropertyCategory.Visible = true;
                btnFinish.Enabled = false;
                dtDocPropertyRequired = objMBRAuthDAL.DocPropertyRequired_SelectAll(OpUnitID: mintOpUnitID, DocTypeID: mintDocTypeID);
                tabEditProperty.Visible = true;
                PropertyReset();

                SetToolTip(btnCancel, "Terminate Function with NO Output");
                SetToolTip(btnFinish, "Save Document, then Terminate Function");
                SetToolTip(btnPropertyAdd, "Add above selected Module");
                SetToolTip(btnBack, "Show Previous Wizard Page");
                SetToolTip(btnNext, "Show Next Wizard Page");
                SetToolTip(cmdPropertyValueBracket, "Make Property Value equal Braketed Variable Name");
                SetToolTip(dgvDocProperty, "dgvDocProperty: DocID=" + mDocID + " Property");
                SetToolTip(btnPropertyUpdate, "Update Property");
                SetToolTip(btnPropertyDelete, "Delete Property");
                SetToolTip(btnDescribe, "Describe Property");
                SetToolTip(dgvPropertySource, "dgvPropertySource : DocID=" + mDocID + " MasterPRopertySource expose source to user");
                SetToolTip(btnSyncSysProperties, "Synchronise Document's System Properties");
                SetToolTip(dgvDocPropertyUnused, "dgvDocPropertyUnused : DocID=" + mDocID + " MasterPRopertySource expose source to user");
                SetToolTip(btnUnusedPropertyDelete, "Delete SELCTED Unused Property from this Document");
                SetToolTip(btnUnusedPropertyDeleteAll, "Delete ALL Unused Property from this Document");
                SetToolTip(label0, "To skip to this step, click here.");
                SetToolTip(shp0, "To skip to this step, click here.");
                SetToolTip(shp1, "To skip to this step, click here.");
                SetToolTip(shp2, "To skip to this step, click here.");
                SetToolTip(shp3, "To skip to this step, click here.");
                SetToolTip(label1, "To skip to this step, click here.");
                SetToolTip(label2, "To skip to this step, click here.");
                SetToolTip(label3, "To skip to this step, click here.");
                SetToolTip(label4, "To skip to this step, click here.");
                SetToolTip(label5, "To skip to this step, click here.");

                Cursor.Current = Cursors.Default;
                //performance
                wordKeepAliveTimer = new System.Windows.Forms.Timer();
                wordKeepAliveTimer.Interval = 1 * 60 * 1000; // every 1 minutes (adjust if needed)
                wordKeepAliveTimer.Tick += WordKeepAliveTimer_Tick;
                wordKeepAliveTimer.Start();
                //end
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void FormEditProperty_Shown(object sender, EventArgs e)
        {
            string strActiveTransaction = string.Empty;
            const string myProcName = "FormEditProperty_Activate";

            if (mblnCancel)
            {
                this.Close();
                CommonService.CleanUpBackgroundWordProcesses();
                return;
            }

            UserForm__Positioning();

            strActiveTransaction = DatabaseHelper.Transaction_Active();
            if (!string.IsNullOrEmpty(strActiveTransaction) && strActiveTransaction != mstrDocName)
            {
                MessageBox.Show("CANNOT RUN THIS LOGIC, DUE TO ACTIVE DATABASE TRANSACTION FOR '" + strActiveTransaction, myProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.Close();
                CommonService.CleanUpBackgroundWordProcesses();
                return;
            }

            if (mblnCalledWizard)//070125
            {
                btnFinish.Text = "Close";
            }
            else
            {
                LoadDocumentProperty();
                btnFinish.Text = "Save";
                btnBack.Enabled = false;//030225
            }

            GetPropertyAll(); // cbopropertyall_load
            mblnCalledWizard = false;

        }

        private void lblReqCount_Click(object sender, EventArgs e)
        {
            MessageBox.Show("'   '   =  OK: no user action needed, and no application errors"
                + "\r\n" + "' * '    =  user must enter/change this value"
                + "\r\n" + "'!A'     =  for Authoring Property in BatchRecord, value must be NOT bracketed"
                + "\r\n" + "'!B'     =  for BatchRecord Property in Authoring, value must be bracketed"
                + "\r\n" + "' D'     =  bad parm: invalid @DocID"
                + "\r\n" + "'!D'     =  for 'ModuleTableRows','ModuleTitle' in Batch Record, property must not exist"
                + "\r\n" + "'!E'     =  for 'ModuleTableRows','ModuleTitle' in a Master, property must not exist"
                + "\r\n" + "' F'     =  bad parm: invalid @PropertyOf"
                + "\r\n" + "'!F'     =  for 'ModuleTableRows','ModuleTitle' in a Module, value must be bracketed"
                + "\r\n" + "'!G'     =  for 'PrintPages', value must always be 'n'"
                + "\r\n" + "'!L'     =  for 'LotNum' n Authoring, value must always be bracketed"
                + "\r\n" + "'!M'     =  for a 'MOD_xxx' Property, value must always be bracketed"
                + "\r\n" + "'!N'     =  for a 'MTL_xxx' Property, value must always be bracketed"
                + "\r\n" + "' P'     =  bad parm: invalid @PropertyName (or unknown PropertyUsage) "
                + "\r\n" + "'!P'     =  for 'PrintDate' and 'PrintTime', value must always be bracketed"
                + "\r\n" + "'!Q'     =  for various Master Properties in a BatchRecord, value must be NOT bracketed"
                + "\r\n" + "'!R'     =  for various Master Properties in a Module, value must be bracketed"
                + "\r\n" + "'!S'     =  for various Master Properties in a Master, value must be NOT bracketed"
                + "\r\n" + "'!T'     =  for 'DocumentTitle', value must be supplied"
                + "\r\n" + "'!X'     =  for Authoring Property in Module, value must be bracketed", "Requirement Indicator Key", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /*private void btnReset_Click(object sender, EventArgs e) // not used in old as it is disabled
        {
            btnDelete.Enabled = false;
            btnReset.Enabled = false;
            btnShowDetails.Enabled = false;
            lstMasterPropertySource.Locked = false;
            lblPropertyFocus.Caption = "";            
        }*/

        private void dgvPropertySource_CellDoubleClick(object sender, EventArgs e)
        {
            int ListIndex;
            {

                ListIndex = dgvPropertySource.CurrentCell.RowIndex; // 
                MessageBox.Show("Details for selected Property in Source Document:" + "\r\n" + "\r\n" +
                    "Property Name:" + "\r\n" + "\t" + dgvPropertySource.Rows[ListIndex].Cells[qmpColMasterPropertyInstance].Value + "\r\n" + "\r\n" +
                    "Property Value:" + "\r\n" + "\t" + dgvPropertySource.Rows[ListIndex].Cells[qmpColMasterPropertyValue].Value + "\r\n" + "\r\n" +
                    "Property in Document:" + "\r\n" + "\t" + dgvPropertySource.Rows[ListIndex].Cells[qmpColMasterPropertyDocName].Value + "\r\n" + "\r\n" +
                    "Document Location:" + "\r\n" + "\t" + dgvPropertySource.Rows[ListIndex].Cells[qmpColMasterPropertyDocLocation].Value);
            }

        }

        private void UserForm__SHUTDOWN(qReturn pqrtQReturn)
        {
            mqrtQReturn = pqrtQReturn;
            if (!mblnCalledWizard)
            {
                MessageBox.Show("ERROR: Called UserForm__SHUTDOWN(" + pqrtQReturn + ")" + ", but this is not a CalledWizard.", MyCompName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (this.Visible)
            {
                this.Hide();
                //int s = dgvDocProperty.ColumnCount; //0701
            }

        }

        private void UserForm__Positioning()
        {
            if (mlngMyTop > 0 & mlngMyLeft > 0)
            {
                this.Top = (int)mlngMyTop;
                this.Left = (int)mlngMyLeft;
            }
        }

        private void QInitialize(qRequest pqrqRequest)
        {
            const string MyProcName = "QInitialize";
            mblnCalledWizard = true;
            mqrqCalledRequest = pqrqRequest;

            switch (pqrqRequest)
            {
                case qRequest.qRequest_Display:
                    {
                        mintParmsValidated = 0;
                        break;
                    }

                case qRequest.qRequest_Validate:
                    {
                        mintParmsValidated = 0;
                        break;
                    }

                case qRequest.qRequest_Write:
                    {
                        if (mintParmsValidated != cintParmsValidated) //mintParmsValidated = 0; So throwing ex
                        {
                            //ToDo Information.Err.Raise(42981, MyCompName + "." + MyProcName, "Attempting QWrite with unvalidated data");
                            var ex = new Exception();
                            ex.Data["errNum"] = "42981";
                            ex.Data["errDesc"] = "Attempting QWrite with unvalidated data";
                            throw ex;
                        }
                        break;
                    }

                case qRequest.qRequest_DBDownLoad:
                    {
                        mintParmsValidated = 0;
                        break;
                    }
            }

            //Todo shpMapCollapse.Visible = true;
        }

        public qReturn QDisplay(QDisplay_FirstLast QFirstLast)
        {
            mblnCalledWizard = true;
            QInitialize(qRequest.qRequest_Display);
            this.ShowDialog();
            return mqrtQReturn;
        }
        public void QLet(int QTop, int QLeft, qpsEditPropertyParms p)
        {
            const string MyProcName = "QLet";
            mblnCalledWizard = true;
            mlngMyTop = QTop;
            mlngMyLeft = QLeft;
            if (p.intEditPropertyRun != 0 & p.intEditPropertyRun != cintEditPropertyRun)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41987";
                ex.Data["errDesc"] = "INVALID RUNCODE=" + p.intEditPropertyRun;
                throw ex;

            }
            if (p.intEditPropertyValidated != 0 & p.intEditPropertyValidated != cintParmsValidated)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41989";
                ex.Data["errDesc"] = "INVALID RUNCODE=" + p.intEditPropertyValidated;
                throw ex;
            }
            if (p.intEditPropertyRun == 0)
            {
                p.strEditPropertySelected = new string[0, 0];

            }


            mintParmsValidated = p.intEditPropertyValidated;
            if (p.strEditPropertySelected.GetUpperBound(1) > 0)
            {
                DataGridView_From_Array(dgvDocProperty, p.strEditPropertySelected);
            }
            else
            {
                if (dgvDocProperty.ColumnCount < qdpColCount) //070125
                {
                    if (dgvDocProperty.Columns.Count < qdpColCount)
                    {

                        GetActiveDocValues();
                        //dgvDocProperty.ColumnCount = qdpColCount; 220125
                        LoadDocumentProperty();
                    }
                }
                //060125
                /*GetActiveDocValues();
                dgvDocProperty.ColumnCount = qdpColCount;
                LoadDocumentProperty();*/

            }

        }

        public void QGet(int QTop, int QLeft, qpsEditPropertyParms p)
        {
            QTop = this.Top;
            QLeft = this.Left;
            p.intEditPropertyRun = cintEditPropertyRun;
            p.intEditPropertyValidated = mintParmsValidated;
            DataGridView_To_Array(dgvDocProperty, p.strEditPropertySelected);
        }

        public qReturn QValidate()
        {
            QInitialize(qRequest.qRequest_Validate);
            if (Valid_Entries())
            {
                return qReturn.qReturn_Valid;
            }
            else
            {
                this.Show();
                return mqrtQReturn;
            }
        }
        public qReturn QAdd(string pPropertyName, string pPropertyValue, int pPropertyIndex = 0, string pPropertyUOM = "")
        {
            int intScan;
            int DocPropertyListIndex;
            mblnCalledWizard = true;
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            QInitialize(pqrqRequest: qRequest.qRequest_Add);

            if (dtPropertyAll.Rows.Count == 0)
            {
                GetPropertyAll();
            }
            if (Property_Selected(pPropertyName: pPropertyName))
            {
            }
            else
            {
                MessageBox.Show("DID NOT FIND PropertyName=" + pPropertyName + " for ADD", MyCompName, MessageBoxButtons.OK);
                return mqrtQReturn;
            }

            if (PropertyName != pPropertyName)
            {
                MessageBox.Show("LOGIC FAILED: selected PropertyName='" + pPropertyName + "' should be '" + PropertyName + "'", MyCompName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (dgvDocProperty.ColumnCount < qdpColCount)
            {
                GetActiveDocValues();//060125
                LoadDocumentProperty();
                txtPropertyValue.Text = pPropertyValue;
            }

            spnPropertyIndex.Minimum = pPropertyIndex; //050225
            spnPropertyIndex.Value = pPropertyIndex;
            txtPropertyValue.Text = pPropertyValue;

            if (!Property_Valid(ppvmMode: PropertyValidationMode.PropertyValidation_Apply, vPropertyInstance: pPropertyName, vPropertyValue: pPropertyValue, vPropertyValueDisplay: PropertyValueDisplay, vPropertyMaterialID: -1, vPropertyWorkCenterID: -1, vPropertyIndex: pPropertyIndex)) // 220125
            {
                return mqrtQReturn;
            }
            isNew = true;
            PropertyDataApply(propertyName: PropertyName, propertyInstance: PropertyInstance, propertyIndex: Convert.ToString(pPropertyIndex), PropertyValueRequired: objMBRAuthDAL.GetProperty_ValueRequired(DocID: mDocID, PropertyName: PropertyName, PropertyValue: PropertyValue, PropertyIndex: System.Convert.ToInt32(PropertyIndex)), PropertyValue: PropertyValue, PropertyUOM: PropertyUOM, PropertyValueDisplay: PropertyValueDisplay, PropertyMaterialID: PropertyMaterialID, PropertyWorkCenterID: PropertyWorkCenterID);
            //PropertyDataApply(propertyName: pPropertyName,PropertyValueRequired: objMBRAuthDAL.GetProperty_ValueRequired(DocID: mDocID, PropertyName: pPropertyName, PropertyIndex: pPropertyIndex, PropertyValue: pPropertyValue),  PropertyValue: pPropertyValue, PropertyUOM: pPropertyUOM, PropertyValueDisplay: PropertyValueDisplay, PropertyMaterialID: -1, PropertyWorkCenterID: -1,propertyIndex:Convert.ToString(pPropertyIndex), propertyInstance : PropertyInstance); //170125
            lblReqCount__FILL();
            if (mqrtQReturn == qReturn.qReturn_Finish)
            {
                mintParmsValidated = cintParmsValidated;
            }
            return qReturn.qReturn_Valid;
        }

        public qReturn QDelete(string pstrPropertyName)
        {
            int intScan;
            int intMaterialCount;

            QInitialize(pqrqRequest: qRequest.qRequest_Delete);
            if (mqrtQReturn != qReturn.qReturn_Finish && mqrtQReturn != qReturn.qReturn_Cancel)
            {
                UserForm__SHUTDOWN(pqrtQReturn: qReturn.qReturn_Cancel);
            }
            return mqrtQReturn;
        }

        public qReturn QWrite(QWriteWhat pqwwWriteWhat, bool pblnReplaceTable = false)
        {
            const string MyProcName = "QWrite";
            QInitialize(qRequest.qRequest_Write);
            mblnCalledWizard = true;
            switch (pqwwWriteWhat)
            {
                case QWriteWhat.qWriteDBUpdate:
                    {
                        Setup_Database(Update: true, dgvDocProperty);
                        break;
                    }

                case QWriteWhat.qWriteMBRProperty:
                    {
                        Setup_Document_Property();
                        break;
                    }

                case QWriteWhat.qWriteMBRTable:
                    {
                        Setup_Document_Text(pblnReplaceTable: pblnReplaceTable);
                        break;
                    }

                default:
                    {
                        MessageBox.Show("UNEXPECTED REQUEST '" + pqwwWriteWhat + "'", MyCompName + "." + MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }

            return qReturn.qReturn_Written;
        }

        public qReturn QDBDownload()
        {
            const string MyProcName = "QDBDownload";
            mblnCalledWizard = true;
            QInitialize(qRequest.qRequest_DBDownLoad);
            Setup_Document_Property();
            // Setup_Document_Text(pblnReplaceTable: true); Not required; take more time to load
            return qReturn.qReturn_Finish;
        }

        public void QClose()
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
            
        }

        private void LoadDocumentProperty()
        {
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();

            try
            {

                Fill_Document_Property();
                lblReqCount__FILL();
                Fill_WordDocument_Property(pblnShowAllProperty: true, pdgvDocProperty: dgvDocProperty);
            }

            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "LoadDocumentProperty");
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "LoadDocumentProperty");
                }
            }

        }
        private void GetPropertyCategory()
        {
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();



            //lstPropertyCat.Insert(0, "[Show All]");
            dtPropertyCategory = objMBRAuthDAL.GetPropertyCategory();
            /*DataRow dr = dtPropertyCategory.NewRow();
            dr[0] = "[Show All]";
            dr[1] = "[Show All]";
            dtPropertyCategory.Rows.InsertAt(dr, 0);*/
            if (dtPropertyCategory != null && dtPropertyCategory.Rows.Count > 0)
            {
                for (int intRow = 0; intRow <= dtPropertyCategory.Rows.Count - 1; intRow++)
                {

                    lstPropertyCat.Add(dtPropertyCategory.Rows[intRow]["PropertyCategory"].ToString() + "|" + dtPropertyCategory.Rows[intRow]["PropertyCategoryDescription"].ToString().Trim());


                    ddlPropertyCategory.DisplayMember = "PropertyCategory";
                    ddlPropertyCategory.ValueMember = "PropertyCategoryDescription";
                    //ddlPropertyCategory.DataSource = dtPropertyCategory;
                    ddlPropertyCategory.DataSource = lstPropertyCat;
                    ddlPropertyCategory.Enabled = true;

                }
            }

        }

        private void ddlPropertyCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            SetToolTip(ddlPropertyCategory, "ddlPropertyCategory : PropertyCategory");
            GetPropertyAll();
            Cursor.Current = Cursors.Default;
        }

        private void GetPropertyAll()
        {
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            string strPropertyCategory;
            string PropertyCategoryValue = string.Empty;

            if (lstPropertyCat.Count == 0)
            {
                return;
            }
            if (ddlPropertyCategory.Text == "[Show All]")
            {
                PropertyCategoryValue = "";
            }
            else
            {
                strPropertyCategory = ddlPropertyCategory.Text.ToString();
                PropertyCategoryValue = strPropertyCategory.Split('|')[0].ToString();

                //PropertyCategoryValue = ddlPropertyCategory.SelectedValue.ToString();
                lblPromptDoubleClick.Text = "Double-Click a property to Update or Delete";
                lblTitle.Text = "Available Properties "; // 

                dtPropertyAll = frmManager.Document_Property_Select(DocIsModule, PropertyCategoryValue);
                if (dtPropertyAll != null && dtPropertyAll.Rows.Count > 0)
                {
                    ddlPropertyAll.DisplayMember = "PropertyName";
                    ddlPropertyAll.ValueMember = "PropertyName";
                    ddlPropertyAll.DataSource = dtPropertyAll;
                    ddlPropertyAll.Enabled = true;
                }

            }
        }

        private void ddlPropertyUOM_KeyDown(object sender, KeyEventArgs e)//comp
        {
            //Todo:   
            //AutoCompleteCombo(cboMaterial, e.KeyCode)
            SetToolTip(ddlPropertyUOM, ddlPropertyUOM.Text);
        }
        private void ddlPropertyAll_KeyDown(object sender, KeyEventArgs e)//comp
        {
            //Todo:   
            //AutoCompleteCombo(cboMaterial, e.KeyCode)
            SetToolTip(ddlPropertyAll, ddlPropertyAll.Text);
        }

        private void ddlPropertyAll_SelectedIndexChanged(object sender, EventArgs e)
        {
            string MyProcName = "ddlPropertyAll_SelectedIndexChanged";
            Cursor.Current = Cursors.WaitCursor;
            SetToolTip(ddlPropertyAll, "ddlPropertyAll : Property");
            try
            {
                FillPropertyDetails();
            }
            catch (Exception ex)
            {
                if (ex.Data["errNum"] != null) // NewCode2
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                    return;
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, CommonService.ErrorProcLatest(MyProcName));
                    return;
                }
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
            
            

        }
        private void FillPropertyDetails() // 
        {
            const string MyProcName = "ddlPropertyAll_Change";
            int PropertyListIndex;
            string PropertyValueAs;
            const string PropertyValueAs_TextBox = "SimpleTextbox";
            const string PropertyValueAs_ComboBox = "TrueCombobox";
            const string PropertyValueAs_ComboList = "ComboAsList";
            const string PropertyValueAs_Bracket = "ForceBrackets";
            const string PropertyValueAs_Disable = "DisableChg";
            bool PropertyIndexVisible = false;
            //string PropertyUsageName = string.Empty; //Comp
            //            bool IsPropertyDecIntWithLookup;
            PropertyListIndex = ddlPropertyAll.SelectedIndex;
            if (PropertyListIndex < 0)
            {
                return;
            }
            PropertyName = (dtPropertyAll.Rows[PropertyListIndex]["PropertyName"]).ToString();
            PropertyDescription = (dtPropertyAll.Rows[PropertyListIndex]["PropertyDescription"]).ToString();
            PropertyDataType = (dtPropertyAll.Rows[PropertyListIndex]["PropertyDataType"]).ToString();
            PropertyDataTypeName = (dtPropertyAll.Rows[PropertyListIndex]["PropertyDataTypeName"]).ToString();
            PropertyIndexMax = Convert.ToInt32(dtPropertyAll.Rows[PropertyListIndex]["PropertyIndexMax"]);
            PropertyUsage = (dtPropertyAll.Rows[PropertyListIndex]["PropertyUsage"]).ToString();
            PropertyLookupType = (dtPropertyAll.Rows[PropertyListIndex]["PropertyLookupType"]).ToString();
            PropertyUOMLookupType = (dtPropertyAll.Rows[PropertyListIndex]["PropertyUOMLookupType"]).ToString();
            PropertyTextMaxLen = Convert.ToInt32(dtPropertyAll.Rows[PropertyListIndex]["PropertyTextMaxLen"]);
            string strPropLookUp = Convert.ToString(dtPropertyAll.Rows[PropertyListIndex]["PropertyLookupType"]).ToString();
            if (left(strPropLookUp, 10) == "<NoLookup>")
            {
                PropertyDecIntWithLookup = false;
            }
            else
            {
                PropertyDecIntWithLookup = true;
            }

            UOM_LookupType = string.Empty;//170125 resolving INVALID LOOKUPTYPE '' F0323

            cmdPropertyValueSize.Visible = false;
            PropertyValueAs = "X";
            lblPropertyValue.Text = "Property Value" + "\r\n" + "(" + PropertyDataTypeName + ")";

            switch (PropertyUsage)
            {
                case "A":
                    {
                        PropertyUsageName = "assigned in Authoring";
                        if (mblnModuleRecord)
                        {
                            PropertyValueAs = PropertyValueAs_Bracket;
                        }
                        else
                        {
                            switch (PropertyDataType)
                            {
                                case PropertyDataType_Integer:
                                    {
                                        if (PropertyDecIntWithLookup)
                                        {
                                            PropertyValueAs = PropertyValueAs_ComboBox;
                                            ddlPropertyValue.Visible = true;
                                            Fill_Lookup(MyCompName, PropertyLookupType.Trim(), "ddlPropertyValue");
                                            ddlPropertyValue__UnassignedAddItem(ValueColumn: qluColLookupText, KeyColumn: -1);
                                        }
                                        else
                                            PropertyValueAs = PropertyValueAs_TextBox;
                                        break;
                                    }
                                case PropertyDataType_Date:
                                    {
                                        PropertyValueAs = PropertyValueAs_TextBox;
                                        break;
                                    }
                                case PropertyDataType_Decimal:
                                    {
                                        if (PropertyDecIntWithLookup)
                                        {
                                            PropertyValueAs = PropertyValueAs_ComboBox;
                                            ddlPropertyValue.Visible = true;
                                            Fill_Lookup(MyCompName, PropertyLookupType.Trim(), "ddlPropertyValue");
                                            ddlPropertyValue__UnassignedAddItem(ValueColumn: qluColLookupText, KeyColumn: -1);
                                        }
                                        else
                                            PropertyValueAs = PropertyValueAs_TextBox;
                                        break;
                                    }
                                case PropertyDataType_LookupText:
                                    {
                                        PropertyValueAs = PropertyValueAs_ComboBox;
                                        ddlPropertyValue.Visible = true;
                                        //txtPropertyValue.Visible = false; Comp
                                        PropertyValue_ValueColumn = qluColLookupText;
                                        Fill_Lookup(MyCompName, PropertyLookupType.Trim(), "ddlPropertyValue");
                                        ddlPropertyValue__UnassignedAddItem(ValueColumn: qluColLookupText, KeyColumn: -1);
                                        break;
                                    }
                                case PropertyDataType_Lookup:
                                    {
                                        PropertyValueAs = PropertyValueAs_ComboList;//Comp
                                        ddlPropertyValue.Visible = true;
                                        PropertyValue_ValueColumn = qluColLookupText;
                                        Fill_Lookup(MyCompName, PropertyLookupType.Trim(), "ddlPropertyValue");
                                        ddlPropertyValue__UnassignedAddItem(ValueColumn: qluColLookupText, KeyColumn: -1);
                                        break;
                                    }
                                case PropertyDataType_Text:
                                    {
                                        PropertyValueAs = PropertyValueAs_TextBox;
                                        break;
                                    }
                                case PropertyDataType_Paragraph:
                                    {
                                        PropertyValueAs = PropertyValueAs_TextBox;
                                        cmdPropertyValueSize.Visible = true;
                                        break;
                                    }
                                case PropertyDataType_Material:
                                    {
                                        PropertyValueAs = PropertyValueAs_ComboList;
                                        ddlPropertyValue.Visible = true;
                                        PropertyValue_ValueColumn = qdlColMaterialDisplay;
                                        Fill_Document_Material(mDocID);
                                        break;
                                    }
                                default:
                                    {
                                        MessageBox.Show("UNEXPECTED PropertyDataType" + PropertyDataType, MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        break;
                                    }
                            }
                        }
                        break;
                    }

                case "B":
                    {
                        PropertyUsageName = "assigned in BatchRecord";
                        PropertyValueAs = PropertyValueAs_Bracket;
                        break;
                    }
                case "S":
                    {
                        PropertyUsageName = "reserved for MBR";
                        switch (PropertyName)
                        {
                            case cChangeNum:
                                {
                                    PropertyValueAs = PropertyValueAs_TextBox;
                                    break;
                                }
                            case cLotNum:
                                {
                                    PropertyValueAs = PropertyValueAs_Bracket;
                                    break;
                                }
                            case cOrderNum:
                                {
                                    PropertyValueAs = PropertyValueAs_Bracket;
                                    break;
                                }
                            case cPrintDate:
                                {
                                    PropertyValueAs = PropertyValueAs_Bracket;
                                    break;
                                }
                            case cPrintTime:
                                {
                                    PropertyValueAs = PropertyValueAs_Bracket;
                                    break;
                                }
                            case cProductID:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cProductName:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cManufacturing_Product_Code:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cNDC_Code:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cProduct_Identification_Code:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cRCS_Sales_Code:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cGlobal_Trade_ID_Number:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            //added fro GTIN - PANNEJ02 - start
                            case cVialLabelGlobalTradeIDNumber:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            //added fro GTIN - PANNEJ02 - end
                            case cModuleTableRows:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cModuleTitle:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cPrintPages:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cDocumentTitle:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cRecordType:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            case cFirstLotUsage:
                                {
                                    PropertyValueAs = PropertyValueAs_Disable;
                                    break;
                                }
                            default:
                                {
                                    if (PropertyName.Substring(0, 4) == "MOD_" || PropertyName.Substring(0, 4) == "MTL_")
                                        PropertyValueAs = PropertyValueAs_Bracket;
                                    else
                                        MessageBox.Show("WARNING: System Property '" + PropertyName + "' not recognized by " + MyCompName + "  (assumed [" + PropertyName + "])", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    break;
                                };

                        }
                        break;
                    }

                default:
                    {
                        MessageBox.Show("UNEXPECTED PropertyUsage '" + PropertyUsage + "'", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        break;

                    }
            }

            switch (PropertyValueAs)
            {
                case PropertyValueAs_TextBox:
                    {
                        PropertyUsageName = PropertyUsageName + ", available for user update";
                        PropertyValueAsName = "user entry";
                        ddlPropertyValue.Visible = false;
                        txtPropertyValue.Visible = true;
                        txtPropertyValue.Enabled = true;
                        txtPropertyValue.Text = "";
                        txtPropertyValue.MaxLength = PropertyTextMaxLen;
                        cmdPropertyValueBracket.Visible = true;
                        break;
                    }
                case PropertyValueAs_Disable:
                    {
                        PropertyUsageName = PropertyUsageName + ", protected from user-entry";
                        PropertyUsageName = "disabled - protected from user modification";
                        txtPropertyValue.Visible = true;
                        txtPropertyValue.Enabled = false;
                        ddlPropertyValue.Visible = false;
                        cmdPropertyValueBracket.Visible = false;
                        //GTIN change - pannej02 start
                        if (PropertyName==cVialLabelGlobalTradeIDNumber)
                        {
                            txtPropertyValue.Text = "[VialGTIN]";
                        }
                        //GTIN change - pannej02 end
                        break;
                    }
                case PropertyValueAs_Bracket: //Modified block 070125
                    {
                        PropertyUsageName = PropertyUsageName + ", forced to be Name in brackets";
                        txtPropertyValue.Visible = true;
                        txtPropertyValue.Enabled = false;
                        txtPropertyValue__Forced();
                        ddlPropertyValue.Visible = false;
                        cmdPropertyValueBracket.Visible = false;
                        break;
                    }
                case PropertyValueAs_ComboBox:
                    {
                        PropertyValueAsName = "user-selection from list, with user-overtype allowed";
                        txtPropertyValue.Visible = false;
                        txtPropertyValue.Enabled = false;
                        cmdPropertyValueBracket.Visible = false;
                        ddlPropertyValue.DropDownStyle = ComboBoxStyle.DropDown;
                        break;
                    }
                case PropertyValueAs_ComboList:
                    {
                        PropertyValueAsName = "user-selection from list, with NO user-overtype allowed";


                        txtPropertyValue.Visible = false;
                        txtPropertyValue.Enabled = false;
                        cmdPropertyValueBracket.Visible = false;
                        if (PropertyDecIntWithLookup)
                            ddlPropertyValue.DropDownStyle = ComboBoxStyle.DropDown;
                        else
                            ddlPropertyValue.DropDownStyle = ComboBoxStyle.DropDownList;
                        break;
                    }
                default:
                    {
                        PropertyValueAsName = "other (" + PropertyValueAs + ")- protected from user modification";
                        txtPropertyValue.Visible = false;
                        txtPropertyValue.Enabled = false;
                        ddlPropertyValue.Visible = false;
                        cmdPropertyValueBracket.Visible = false;
                        ddlPropertyValue.Enabled = true;
                        /*                      //Comp
                                                PropertyUsageName = PropertyUsageName + ", ValueAs=" + PropertyValueAs;
                                                txtPropertyValue__Forced();*/
                        break;
                    }

            }

            if (PropertyIndexMax == 0)
            {
                PropertyIndexVisible = false;
                spnPropertyIndex.Visible = true;
                spnPropertyIndex.Minimum = 0;
                spnPropertyIndex.Maximum = 0;
                spnPropertyIndex.Text = "0";
            }
            else
            {
                PropertyIndexVisible = true;
                spnPropertyIndex.Visible = true;
                spnPropertyIndex.Maximum = PropertyIndexMax;
                spnPropertyIndex.Minimum = 1;
                spnPropertyIndex.Text = "1";
                /*
                                if (spnPropertyIndex.Maximum == PropertyIndexMax)
                                {

                                }
                                else
                                    spnPropertyIndex.Maximum = PropertyIndexMax;
                                if (spnPropertyIndex.Minimum == 1)
                                {

                                }
                                else
                                {
                                    spnPropertyIndex.Minimum = 1;
                                    spnPropertyIndex.Text = "1";
                                }*/

            }

            lblPropertyIndex.Visible = PropertyIndexVisible;
            spnPropertyIndex.Visible = PropertyIndexVisible;
            Property_UOM_Setup(TestProperty: true);
        }


        private void cmdPropertyValueBracket_Click(object sender, EventArgs e)
        {
            txtPropertyValue.Text = PropertyBracketed(Property_Instance_Unite(pPropertyName: PropertyName, pPropertyIndex: spnPropertyIndex.Text));
            ddlPropertyUOM.Text = "";
        }
        private void txtPropertyValue_TextChanged(object sender, EventArgs e)
        {
            if (txtPropertyValue.Text.ToUpper() == (PropertyBracketed(Property_Instance_Unite(pPropertyName: PropertyName, pPropertyIndex: spnPropertyIndex.Text))).ToUpper())
            {
                Property_UOM_Setup(TestProperty: false);
            }
            else
            {
                Property_UOM_Setup(TestProperty: true);
            }
            lblSpaceCount.Text = Convert.ToString(txtPropertyValue.MaxLength - txtPropertyValue.TextLength);
        }

        public void lblReqCount__FILL()
        {
            int ListIndex;
            int ReqCount;
            ReqCount = 0;
            for (ListIndex = 0; ListIndex <= this.dgvDocProperty.Rows.Count - 1; ListIndex++)
            {
                if (dgvDocProperty.Rows[ListIndex].Cells[qdpColPropertyValueRequired].Value.ToString().Trim() == "")
                {
                }
                else
                {
                    ReqCount = ReqCount + 1;
                }
            }

            if (ReqCount > 0)
            {
                lblReqCount.Text = Convert.ToString(ReqCount);
                lblReqCount.ForeColor = Color.FromArgb(255, 0, 0);
            }
            else
            {
                lblReqCount.Text = "0";
                lblReqCount.ForeColor = Color.FromArgb(0, 0, 0);
            }
        }

        string UOM_LookupType = string.Empty;
        /*private void Property_UOM_Setup(bool TestProperty) // 
        {
            bool PropertyUOMVisible;
            //string UOM_LookupType = string.Empty;

            if (TestProperty)
            {
                if (PropertyUOMLookupType == "")
                {
                    PropertyUOMVisible = false;
                }
                if (mblnModuleRecord)
                {
                    PropertyUOMVisible = false;
                }
                else
                {
                    PropertyUOMVisible = true;
                    if (UOM_LookupType != (PropertyUOMLookupType).Trim())
                    {
                        UOM_LookupType = PropertyUOMLookupType.Trim();
                        Fill_Lookup(this.Name, UOM_LookupType, "ddlPropertyUOM");
                    }
                    else
                    {
                    }

                    if (ddlPropertyUOM.Items.Count > 0)
                    {
                        ddlPropertyUOM.SelectedIndex = 0;
                    }
                }
            }
            else
            {
                PropertyUOMVisible = false;
            }

            lblPropertyUOM.Visible = PropertyUOMVisible;
            ddlPropertyUOM.Visible = PropertyUOMVisible;
        }*/

        private void Property_UOM_Setup(bool TestProperty) //170125
        {
            bool PropertyUOMVisible;
            //string UOM_LookupType = string.Empty;

            if (TestProperty)
            {
                if (string.IsNullOrEmpty(PropertyUOMLookupType))
                {
                    PropertyUOMVisible = false;
                }
                else
                {
                    if (mblnModuleRecord)
                    {
                        PropertyUOMVisible = false;
                    }
                    else
                    {
                        PropertyUOMVisible = true;

                        if (UOM_LookupType != PropertyUOMLookupType.Trim())
                        {
                            UOM_LookupType = PropertyUOMLookupType.Trim();
                            Fill_Lookup(this.Name, UOM_LookupType, "ddlPropertyUOM");
                        }
                    }

                    if (ddlPropertyUOM.Items.Count > 0)
                    {
                        ddlPropertyUOM.SelectedIndex = 0;
                    }
                }
            }
            else
            {
                PropertyUOMVisible = false;
            }

            lblPropertyUOM.Visible = PropertyUOMVisible;
            ddlPropertyUOM.Visible = PropertyUOMVisible;
        }

        private void ddlPropertyValue_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            switch (PropertyDataType)
            {
                case "K":
                    {
                        break;
                    }

                case "L":
                    {
                        break;
                    }

                case "M":
                    {
                        if (ddlPropertyUOM.Visible == true)
                        {
                            ddlPropertyUOM.Text = "";
                        }
                        break;
                    }

                default:
                    {
                        if (PropertyDecIntWithLookup)
                        {
                            break;
                        }
                        else
                        {
                            MessageBox.Show("PROGRAM ERROR: ComboBox is exposed for PropertyDataType=" + PropertyDataType, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        break;
                    }
            }
            Cursor.Current = Cursors.Default;
        }

        private void ddlPropertyValue__NotApplicableAddItem(int ValueColumn, int KeyColumn)
        {
            int ColumnIndex;
            {

                ddlPropertyValue.Visible = true;
                ddlPropertyValue.DataSource = null;
                propertyValueLookupDataList.Insert(0, new TblLookupData());

                ddlPropertyValue__NotApplicable_UnassignedValue();

                /*if (KeyColumn > -1)
                {
                    lookupDataList[0].LookupAttributeValue = "-1";
                    //ddlPropertyValue.SetItemText(KeyColumn, 0, "-1");
                }*/
            }
        }
        private void ddlPropertyValue__UnassignedAddItem(int ValueColumn, int KeyColumn) //  working 181025
        {
            int ColumnIndex;
            {
                ddlPropertyValue.Visible = true;
                ddlPropertyValue.DataSource = null;
                propertyValueLookupDataList.Insert(0, new TblLookupData()); //180125
                                                                            //ddlPropertyValue.Items.Insert(0, PropertyBracketed(Property_Instance_Unite(pPropertyName: PropertyName, pPropertyIndex: PropertyIndex)));

                ddlPropertyValue__UnassignedValue();

                if (KeyColumn > -1)
                {
                    propertyValueLookupDataList[0].LookupAttributeValue = "-1";
                    //ddlPropertyValue.SetItemText(KeyColumn, 0, "-1");
                }
            }
        }
        private void ddlPropertyValue__UnassignedValue() //working 170125
        {
            if (!ddlPropertyValue.Visible)
            {
                return;
            }
            if (propertyValueLookupDataList.Count == 0)
                return;
            else
            {
                propertyValueLookupDataList[0].LookupText = PropertyBracketed(Property_Instance_Unite(pPropertyName: PropertyName, pPropertyIndex: PropertyIndex));
                ddlPropertyValue_Load();
            }

        }
        private void ddlPropertyValue__NotApplicable_UnassignedValue()
        {
            if (!ddlPropertyValue.Visible)
            {
                return;
            }
            if (propertyValueLookupDataList.Count == 0)
            {
                return;
            }
            else
            {
                propertyValueLookupDataList[0].LookupAttributeValue = cNotApplicable;
                ddlPropertyValue_Load();
            }

        }
        private void ddlPropertyValue_Load()
        {

            if (propertyValueLookupDataList != null)
            {
                ddlPropertyValue.DataSource = propertyValueLookupDataList;
                ddlPropertyValue.DisplayMember = "LookupText";
                ddlPropertyValue.ValueMember = "LookupText";
                ddlPropertyValue.Enabled = true;

                //ddlPropertyValue.SelectedIndex = 0;
            }

        }
        private void btnBack_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            ShowPrevTab();
            int currentIndex = GetCurrentTabIndex();
            if (currentIndex == 2)
            { PropertyUnused_Load(); }

            if (currentIndex == 1)
            { PropertySource_Load(); }

            if (currentIndex == 0)
            { btnBack.Enabled = false; }

            if (currentIndex == 3)
            { btnNext.Enabled = false; }

            Cursor.Current = Cursors.Default;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            ShowNextTab();
            int currentIndex = GetCurrentTabIndex();
            if (currentIndex == 2)
            { PropertyUnused_Load(); }

            if (currentIndex == 1)
            { PropertySource_Load(); }

            if (currentIndex == 0)
            { btnBack.Enabled = false; }

            if (currentIndex == 3)
            { btnNext.Enabled = false; }

            Cursor.Current = Cursors.Default;
        }


        private void btnDescribe_Click(object sender, EventArgs e)
        {
            MessageBox.Show(PropertyDescription, "Description of Property '" + PropertyName + "'", MessageBoxButtons.OK);
        }

        private void btnPropertyAdd_Click(object sender, EventArgs e)
        {
            isNew = true;
            strPropertyValueRequired = string.Empty;
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            const string MyProcName = "btnPropertyAdd_Click";

            Cursor.Current = Cursors.WaitCursor;

            try
            {
                if (btnPropertyAdd.Text == "Reset")
                {
                    isNew = false;
                    PropertyReset();
                    return;
                }

                // 
                if (!Property_Valid(ppvmMode: PropertyValidationMode.PropertyValidation_Add, vPropertyInstance: PropertyInstance, vPropertyValue: PropertyValue, vPropertyValueDisplay: PropertyValueDisplay, vPropertyMaterialID: PropertyMaterialID, vPropertyWorkCenterID: PropertyWorkCenterID))
                //if (!Property_Valid(ppvmMode: PropertyValidationMode.PropertyValidation_Add, PropertyInstance: ref PropertyInstance, PropertyValue: ref PropertyValue, PropertyValueDisplay: ref PropertyValueDisplay, PropertyMaterialID: PropertyMaterialID, PropertyWorkCenterID: PropertyWorkCenterID))
                {
                    return;
                }

                PropertyDataApply(propertyName: PropertyName, propertyInstance: PropertyInstance, propertyIndex: PropertyIndex, PropertyValueRequired: objMBRAuthDAL.GetProperty_ValueRequired(DocID: mDocID, PropertyName: PropertyName, PropertyValue: PropertyValue, PropertyIndex: System.Convert.ToInt32(PropertyIndex)), PropertyValue: PropertyValue, PropertyUOM: PropertyUOM, PropertyValueDisplay: PropertyValueDisplay, PropertyMaterialID: PropertyMaterialID, PropertyWorkCenterID: PropertyWorkCenterID);
                lblReqCount__FILL();
                PropertyReset();
                btnFinish.Enabled = true;
            }


            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void btnPropertyUpdate_Click(object sender, EventArgs e)
        {
            //int intListIndex; AfterComp
            isUpdate = true;
            strPropertyValueRequired = string.Empty;
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            Cursor.Current = Cursors.WaitCursor;
            if (!Property_Valid(ppvmMode: PropertyValidationMode.PropertyValidation_Update, vPropertyInstance: PropertyInstance, vPropertyValue: PropertyValue, vPropertyValueDisplay: PropertyValueDisplay, vPropertyMaterialID: PropertyMaterialID, vPropertyWorkCenterID: PropertyWorkCenterID))
            //if (!Property_Valid(ppvmMode: PropertyValidationMode.PropertyValidation_Update, PropertyInstance: ref PropertyInstance, PropertyValue: ref PropertyValue, PropertyValueDisplay: ref PropertyValueDisplay, PropertyMaterialID: PropertyMaterialID, PropertyWorkCenterID: PropertyWorkCenterID))
            {
                return;
            }
            //intListIndex = dgvDocProperty.CurrentCell.RowIndex; //  AfterComp
            //strPropertyValueRequired = objMBRAuthDAL.GetProperty_ValueRequired(DocID: mDocID, PropertyName: PropertyName, PropertyValue: PropertyValue, PropertyIndex: System.Convert.ToInt32(PropertyIndex));
            PropertyDataApply(propertyName: PropertyName, propertyInstance: PropertyInstance, propertyIndex: PropertyIndex, PropertyValueRequired: objMBRAuthDAL.GetProperty_ValueRequired(DocID: mDocID, PropertyName: PropertyName, PropertyValue: PropertyValue, PropertyIndex: System.Convert.ToInt32(PropertyIndex)), PropertyValue: PropertyValue, PropertyUOM: PropertyUOM, PropertyValueDisplay: PropertyValueDisplay, PropertyMaterialID: PropertyMaterialID, PropertyWorkCenterID: PropertyWorkCenterID);//AfterComp
            //PropertyDataApply(propertyName: PropertyName, propertyInstance: PropertyInstance, propertyIndex: Convert.ToString(intListIndex), PropertyValueRequired: objMBRAuthDAL.GetProperty_ValueRequired(DocID: mDocID, PropertyName: PropertyName, PropertyValue: PropertyValue, PropertyIndex: System.Convert.ToInt32(PropertyIndex)), PropertyValue: PropertyValue, PropertyUOM: PropertyUOM, PropertyValueDisplay: PropertyValueDisplay, PropertyMaterialID: PropertyMaterialID, PropertyWorkCenterID: PropertyWorkCenterID);
            lblReqCount__FILL();
            PropertyReset();
            //btnDescribe.Top = 132;
            btnFinish.Enabled = true;
            Cursor.Current = Cursors.Default;
        }

        private string ddlPropertyUOM__DATA(string PropertyUOMLookupType, bool NullForBlank = false)
        {
            string PropertyUOMData;
            if (PropertyUOMLookupType == "" || (ddlPropertyUOM.Visible == false))
            {
                lblPropertyUOM.Visible = false;
                ddlPropertyUOM.Visible = false;
                if (NullForBlank)
                {
                    PropertyUOMData = null;
                }
                else
                {
                    PropertyUOMData = "";
                }
            }
            else
            {
                PropertyUOMData = ddlPropertyUOM.Text;
            }
            return PropertyUOMData;
        }

        private void btnSyncSysProperties_Click(object sender, EventArgs e)
        {
            int returnvalue;
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            Cursor.Current = Cursors.WaitCursor;
            syncSysProp = true;
            returnvalue = objMBRAuthDAL.SyncMickProperties(mDocID);
            LoadDocumentProperty();
            btnFinish_Click(sender, e);
            MessageBox.Show("This document's System Properties have been syncronized with values from MICK.\r\nSystem Properties Syncronized");
            Cursor.Current = Cursors.Default;
        }
        private bool Property_Valid(PropertyValidationMode ppvmMode, string vPropertyInstance, string vPropertyValue, string vPropertyValueDisplay, int vPropertyMaterialID, int vPropertyWorkCenterID, int vPropertyIndex = -1) //changed int PropertyIndex = 1 from -1 // 
        {
            const string MyProcName = "Property_Valid";
            string strErrorMessage;
            string strPropertyValueOut = string.Empty;
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            int PropertyValueIndex;
            const int cintColCount = 9;
            string ErrorMessage = string.Empty;

            bool Property_Valid = false;
            vPropertyMaterialID = -1;
            vPropertyWorkCenterID = -1;

            try
            {
                if (dgvDocProperty.Columns.Count != qdpColCount)
                {
                    strErrorMessage = "PROGRAM ERROR: dgvDocProperty.ColumnCount=" + dgvDocProperty.ColumnCount + ", expecting " + qdpColCount;
                    MessageAdd(strErrorMessage);
                    MessageBox.Show(strErrorMessage, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return Property_Valid;
                }

                if (qdpColCount != cintColCount)
                {
                    strErrorMessage = "PROGRAM ERROR: This logic expects " + cintColCount + " Columns, not " + qdpColCount;
                    MessageAdd(strErrorMessage);
                    MessageBox.Show(strErrorMessage, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return Property_Valid;
                }

                if (ddlPropertyAll.Text == "")
                {
                    if (ddlPropertyAll.Enabled)
                        ddlPropertyAll.Focus();
                    strErrorMessage = "No Property selected";
                    MessageAdd(strErrorMessage);
                    MessageBox.Show(strErrorMessage, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return Property_Valid;
                }

                if(ppvmMode == PropertyValidationMode.PropertyValidation_Update)
                    spnPropertyIndex.Text = dgvDocProperty.CurrentRow.Cells[qdpColPropertyIndex].Value.ToString();//040225

                vPropertyIndex = Convert.ToInt32(spnPropertyIndex.Text);//040225
                vPropertyInstance = Property_Instance_Unite(pPropertyName: PropertyName, pPropertyIndex: vPropertyIndex.ToString());//040225
                

                if (ppvmMode == PropertyValidationMode.PropertyValidation_Add)
                {

                    foreach (DataGridViewRow row in dgvDocProperty.Rows)
                    {
                        if (row.Cells[qdpColPropertyInstance].Value.ToString().Equals(vPropertyInstance))
                        {
                            strErrorMessage = "Property Instance '" + vPropertyInstance + "' already in List";
                            MessageAdd(strErrorMessage);
                            MessageBox.Show(strErrorMessage, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return Property_Valid;
                        }
                    }

                }

                if (txtPropertyValue.Visible)
                {
                    if (txtPropertyValue.Enabled)
                    {
                        if (txtPropertyValue.Text == "")
                        {
                            txtPropertyValue.Focus();
                            strErrorMessage = "Please enter a VALUE before adding Property '" + ddlPropertyAll.Text + "'.";
                            MessageAdd(strErrorMessage);
                            MessageBox.Show(strErrorMessage, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return Property_Valid;
                        }
                        else
                            vPropertyValue = txtPropertyValue.Text;
                    }
                    else
                        vPropertyValue = txtPropertyValue.Text;
                }
                else if (ddlPropertyValue.Visible == true)
                {
                    vPropertyValue = ddlPropertyValue.Text;
                    PropertyValueIndex = ddlPropertyValue.SelectedIndex;
                    switch (PropertyDataType)
                    {
                        case "M":
                            {
                                if (PropertyValueIndex < 0)
                                {

                                    MessageBox.Show("No Material Selected", "Property Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    return Property_Valid;

                                }
                                else
                                {
                                    //PropertyMaterialID = Convert.ToInt32(dtPropertyValue.Rows[0][qdlColMaterialDBID].ToString());
                                    vPropertyMaterialID = Convert.ToInt32(ddlPropertyValue.SelectedValue);
                                    break;
                                }
                            }
                    }
                }
                else
                {
                    MessageBox.Show("PROGRAM PropertyValue ERROR: neither TextBox nor Combobox is Visible");
                    return Property_Valid;

                }

                if (vPropertyValue == string.Empty)
                {
                    if (vPropertyInstance != "Manufacturing_Product_Code")
                    {
                    }
                    else if (vPropertyInstance != "NDC_Code")
                    {
                    }
                    else if (vPropertyInstance != "Product_Identification_Code")
                    {
                    }
                    else if (vPropertyInstance != "RCS_Sales_Code")
                    {
                    }
                    else if (vPropertyInstance != "Global_Trade_ID_Number")
                    {
                        MessageBox.Show("PROGRAM PropertyValue ERROR: 'PropertyValue' was not assigned");
                        return Property_Valid;
                    }
                    //added for GTIN - pannej02--- need to check
                    else if (vPropertyInstance != "VialLabelGlobalTradeIDNumber")
                    {
                        MessageBox.Show("PROGRAM PropertyValue ERROR: 'PropertyValue' was not assigned");
                        return Property_Valid;
                    }
                }

                PropertyUOM = ddlPropertyUOM__DATA(PropertyUOMLookupType, NullForBlank: false);
                spnPropertyIndex.Text = vPropertyIndex.ToString();// Comp

                //strPropertyValueOut = objMBRAuthDAL.PropertyValidate(ref ErrorMessage, mDocID, "A", PropertyName, Convert.ToInt32(spnPropertyIndex.Value), vPropertyValue, ddlPropertyUOM__DATA(PropertyUOMLookupType: PropertyUOMLookupType, NullForBlank: true), vPropertyMaterialID, vPropertyWorkCenterID);
                strPropertyValueOut = objMBRAuthDAL.PropertyValidate(ref ErrorMessage, mDocID, "A", PropertyName, Convert.ToInt32(vPropertyIndex), vPropertyValue, ddlPropertyUOM__DATA(PropertyUOMLookupType: PropertyUOMLookupType, NullForBlank: true), vPropertyMaterialID, vPropertyWorkCenterID);//040225

                if (ErrorMessage != "")
                {
                    MessageBox.Show(ErrorMessage, "Property Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }

                vPropertyValue = strPropertyValueOut;


                if (txtPropertyValue.Visible)
                {
                    txtPropertyValue.Text = vPropertyValue;
                }
                else//210125
                {//ddlPropertyValue.Visible
                    if (ddlPropertyValue.DropDownStyle == ComboBoxStyle.DropDown)
                    {
                        ddlPropertyValue.Text = vPropertyValue;
                    }
                    else
                    {
                    }
                }
                // 
                vPropertyValueDisplay = Property_ValueDisplay_Unite(pPropertyValue: vPropertyValue, pPropertyUOM: PropertyUOM);
                PropertyInstance = vPropertyInstance;
                PropertyValue = vPropertyValue;
                PropertyValueDisplay = vPropertyValueDisplay;
                PropertyMaterialID = vPropertyMaterialID;
                PropertyWorkCenterID = vPropertyWorkCenterID;
                PropertyIndex = vPropertyIndex.ToString();

                Property_Valid = true;

            }
            catch (Exception ex)
            {
                vPropertyValueDisplay = Property_ValueDisplay_Unite(pPropertyValue: vPropertyValue, pPropertyUOM: PropertyUOM);
                PropertyInstance = vPropertyInstance;
                PropertyValue = vPropertyValue;
                PropertyValueDisplay = vPropertyValueDisplay;
                PropertyMaterialID = vPropertyMaterialID;
                PropertyWorkCenterID = vPropertyWorkCenterID;
                PropertyIndex = vPropertyIndex.ToString();

                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                }
            }

            return Property_Valid;
        }

        /*private bool Property_Valid(PropertyValidationMode ppvmMode, ref string PropertyInstance, ref string PropertyValue, ref string PropertyValueDisplay, int PropertyMaterialID, int PropertyWorkCenterID, int PropertyIndex = -1) //changed int PropertyIndex = 1 from -1 // 
        {
            const string MyProcName = "Property_Valid";
            string strErrorMessage;
            string strPropertyValueOut = string.Empty;
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            int PropertyValueIndex;
            const int cintColCount = 9;
            string ErrorMessage = string.Empty;

            bool Property_Valid = false;
            PropertyMaterialID = -1;
            PropertyWorkCenterID = -1;

            try
            {
                if (dgvDocProperty.Columns.Count != qdpColCount)
                {
                    strErrorMessage = "PROGRAM ERROR: dgvDocProperty.ColumnCount=" + dgvDocProperty.ColumnCount + ", expecting " + qdpColCount;
                    MessageAdd(strErrorMessage);
                    MessageBox.Show(strErrorMessage, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return Property_Valid;
                }

                if (qdpColCount != cintColCount)
                {
                    strErrorMessage = "PROGRAM ERROR: This logic expects " + cintColCount + " Columns, not " + qdpColCount;
                    MessageAdd(strErrorMessage);
                    MessageBox.Show(strErrorMessage, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return Property_Valid;
                }

                if (ddlPropertyAll.Text == "")
                {
                    if (ddlPropertyAll.Enabled)
                        ddlPropertyAll.Focus();
                    strErrorMessage = "No Property selected";
                    MessageAdd(strErrorMessage);
                    MessageBox.Show(strErrorMessage, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return Property_Valid;
                }
                PropertyIndex = Convert.ToInt32(spnPropertyIndex.Value.ToString());
                PropertyInstance = Property_Instance_Unite(pPropertyName: PropertyName, pPropertyIndex: PropertyIndex.ToString());

                if (ppvmMode == PropertyValidationMode.PropertyValidation_Add)
                {

                    foreach (DataGridViewRow row in dgvDocProperty.Rows)
                    {
                        if (row.Cells[qdpColPropertyInstance].Value.ToString().Equals(PropertyInstance))
                        {
                            strErrorMessage = "Property Instance '" + PropertyInstance + "' already in List";
                            MessageAdd(strErrorMessage);
                            MessageBox.Show(strErrorMessage, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return Property_Valid;
                        }
                    }

                }

                PropertyValue = string.Empty;
                if (txtPropertyValue.Visible)
                {
                    if (txtPropertyValue.Enabled)
                    {
                        if (txtPropertyValue.Text == "")
                        {
                            txtPropertyValue.Focus();
                            strErrorMessage = "Please enter a VALUE before adding Property '" + ddlPropertyAll.Text + "'.";
                            MessageAdd(strErrorMessage);
                            MessageBox.Show(strErrorMessage, MyMsgboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return Property_Valid;
                        }
                        else
                            PropertyValue = txtPropertyValue.Text;
                    }
                    else
                        PropertyValue = txtPropertyValue.Text;
                }
                else if (ddlPropertyValue.Visible == true)
                {
                    PropertyValue = ddlPropertyValue.Text;
                    PropertyValueIndex = ddlPropertyValue.SelectedIndex;
                    switch (PropertyDataType)
                    {
                        case "M":
                            {
                                if (PropertyValueIndex < 0)
                                {

                                    MessageBox.Show("No Material Selected", "Property Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    return Property_Valid;

                                }
                                else
                                {
                                    //PropertyMaterialID = Convert.ToInt32(dtPropertyValue.Rows[0][qdlColMaterialDBID].ToString());
                                    PropertyMaterialID = Convert.ToInt32(ddlPropertyValue.SelectedValue);
                                    break;
                                }
                            }
                    }
                }
                else
                {
                    MessageBox.Show("PROGRAM PropertyValue ERROR: neither TextBox nor Combobox is Visible");
                    return Property_Valid;

                }

                if (PropertyValue == string.Empty)
                {
                    if (PropertyInstance != "Manufacturing_Product_Code")
                    {
                    }
                    else if (PropertyInstance != "NDC_Code")
                    {
                    }
                    else if (PropertyInstance != "Product_Identification_Code")
                    {
                    }
                    else if (PropertyInstance != "RCS_Sales_Code")
                    {
                    }
                    else if (PropertyInstance != "Global_Trade_ID_Number")
                    {
                        MessageBox.Show("PROGRAM PropertyValue ERROR: 'PropertyValue' was not assigned");
                        return Property_Valid;
                    }
                }

                PropertyUOM = ddlPropertyUOM__DATA(PropertyUOMLookupType, NullForBlank: false);
                spnPropertyIndex.Text = PropertyIndex.ToString();// Comp

                strPropertyValueOut = objMBRAuthDAL.PropertyValidate(ref ErrorMessage, mDocID, "A", PropertyName, Convert.ToInt32(spnPropertyIndex.Value), PropertyValue, ddlPropertyUOM__DATA(PropertyUOMLookupType: PropertyUOMLookupType, NullForBlank: true), PropertyMaterialID, PropertyWorkCenterID);

                if (ErrorMessage != "")
                {
                    MessageBox.Show(ErrorMessage, "Property Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }

                PropertyValue = strPropertyValueOut;


                if (txtPropertyValue.Visible)
                {
                    txtPropertyValue.Text = PropertyValue;
                }
                else if (ddlPropertyValue.Visible)
                {
                    ddlPropertyValue.Text = PropertyValue;
                }
                else
                {
                }
                // 
                PropertyValueDisplay = Property_ValueDisplay_Unite(pPropertyValue: PropertyValue, pPropertyUOM: PropertyUOM);
                *//*PropertyInstance = vPropertyInstance;
                PropertyValue = vPropertyValue;
                PropertyValueDisplay = vPropertyValueDisplay;
                PropertyMaterialID = vPropertyMaterialID;
                PropertyWorkCenterID = vPropertyWorkCenterID;
                PropertyIndex = vPropertyIndex.ToString();*//*

                Property_Valid = true;

            }
            catch (Exception ex)
            {
                *//*vPropertyValueDisplay = Property_ValueDisplay_Unite(pPropertyValue: vPropertyValue, pPropertyUOM: PropertyUOM);
                PropertyInstance = vPropertyInstance;
                PropertyValue = vPropertyValue;
                PropertyValueDisplay = vPropertyValueDisplay;
                PropertyMaterialID = vPropertyMaterialID;
                PropertyWorkCenterID = vPropertyWorkCenterID;
                PropertyIndex = vPropertyIndex.ToString();
*//*
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                }
            }

            return Property_Valid;
        }*/
        private string Property_Instance_Unite(string pPropertyName, string pPropertyIndex)
        {
            string Property_Instance_Unite = string.Empty;
            if (pPropertyIndex == "0")
            {
                Property_Instance_Unite = pPropertyName;
            }
            else
            {
                Property_Instance_Unite = pPropertyName + pPropertyIndex;
            }
            return Property_Instance_Unite;
        }

        private void Property_Instance_Split(string pPropertyInstance, string pPropertyName, string pPropertyIndex)
        {
            int InstanceScan;
            string PropertyInstanceChar;
            string IndexString;
            IndexString = "";
            for (InstanceScan = pPropertyInstance.Length; InstanceScan >= 1; InstanceScan += -1)
            {
                PropertyInstanceChar = pPropertyInstance.Substring(InstanceScan - 1, 1);//changed to get substring value//AfterComp 
                if (string.Compare(PropertyInstanceChar, "0") < 0)
                {
                    break;
                }
                if (string.Compare(PropertyInstanceChar, "9") > 0)

                {
                    break;
                }
                IndexString = PropertyInstanceChar + IndexString;
            }

            pPropertyName = left(pPropertyInstance, InstanceScan);
            if (IndexString == "")
            {
                pPropertyIndex = "0";
            }
            else
            {
                pPropertyIndex = IndexString;
            }
            //110125
            PropertyName = pPropertyName;
            PropertyIndex = pPropertyIndex;

        }

        private string Property_ValueDisplay_Unite(string pPropertyValue, string pPropertyUOM)
        {
            string Property_ValueDisplay_Unite = string.Empty;
            if (pPropertyUOM == "")
            {
                Property_ValueDisplay_Unite = pPropertyValue;
            }
            else
            {
                Property_ValueDisplay_Unite = pPropertyValue + " " + pPropertyUOM;
            }
            return Property_ValueDisplay_Unite;
        }

        //private void Property_ValueDisplay_Split(string pPropertyValueDisplay, string pPropertyValue, string pPropertyUOM, string pPropertyName)
        //{
        //    int ValueDisplayScan;
        //    string PropertyValueDisplayChar;
        //    string UOMString;
        //    try
        //    {

        //        if (!Property_Selected(pPropertyName: pPropertyName))
        //        {
        //            //throw new Exception("cannot find Property '" + pPropertyName + "'");
        //            if (!mblnCalledWizard)
        //            {
        //                MessageBox.Show("cannot find Property '" + pPropertyName + "'");
        //            }
        //            return;
        //        }

        //        //if (pPropertyValueDisplay.IndexOf(" ", 1) == 0 || PropertyUOMLookupType == "")
        //        if (pPropertyValueDisplay.IndexOf(" ") == -1 || PropertyUOMLookupType == "")
        //        {
        //            pPropertyValue = pPropertyValueDisplay;
        //            pPropertyUOM = "";
        //            PropertyValue = pPropertyValue; //050225
        //            PropertyUOM = pPropertyUOM;
        //            return;
        //        }

        //        UOMString = "";
        //        for (ValueDisplayScan = pPropertyValueDisplay.Length - 1; ValueDisplayScan >= 0; ValueDisplayScan--)
        //        {
        //            //PropertyValueDisplayChar = pPropertyValueDisplay[ValueDisplayScan - 1].ToString();
        //            PropertyValueDisplayChar = pPropertyValueDisplay[ValueDisplayScan].ToString();
        //            if (PropertyValueDisplayChar == " ")
        //            {
        //                break;
        //            }
        //            UOMString = PropertyValueDisplayChar + UOMString;
        //        }

        //        //pPropertyValue = left(pPropertyValueDisplay, ValueDisplayScan - 1);
        //        pPropertyValue = pPropertyValueDisplay.Substring(0, ValueDisplayScan); //040225
        //        pPropertyUOM = UOMString;

        //        PropertyValue = pPropertyValue; //040225
        //        PropertyUOM = pPropertyUOM;
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}

        // works
        private void Property_ValueDisplay_Split(string pPropertyValueDisplay, string pPropertyValue, string pPropertyUOM, string pPropertyName)

        {

            int ValueDisplayScan;

            string PropertyValueDisplayChar;

            string UOMString;

            try

            {



                if (!Property_Selected(pPropertyName: pPropertyName))

                {

                    //throw new Exception("cannot find Property '" + pPropertyName + "'");

                    if (!mblnCalledWizard)

                    {

                        MessageBox.Show("cannot find Property '" + pPropertyName + "'");

                    }

                    return;

                }



                //if (pPropertyValueDisplay.IndexOf(" ", 1) == 0 || PropertyUOMLookupType == "")

                if (pPropertyValueDisplay.IndexOf(" ") == -1 || PropertyUOMLookupType == "")

                {

                    pPropertyValue = pPropertyValueDisplay;

                    pPropertyUOM = "";

                    PropertyValue = pPropertyValue; //050225

                    PropertyUOM = pPropertyUOM;

                    return;

                }



                /*UOMString = "";

 

                for (ValueDisplayScan = pPropertyValueDisplay.Length - 1; ValueDisplayScan >= 0; ValueDisplayScan--)

                {

                    PropertyValueDisplayChar = pPropertyValueDisplay[ValueDisplayScan].ToString();

                    if (PropertyValueDisplayChar == " ")

                    {

                        break;

                    }

                    UOMString = PropertyValueDisplayChar + UOMString;

                }

 

                pPropertyValue = pPropertyValueDisplay.Substring(0, ValueDisplayScan); //040225

                pPropertyUOM = UOMString;*/



                pPropertyValue = "";



                for (ValueDisplayScan = 0; ValueDisplayScan <= pPropertyValueDisplay.Length - 1; ValueDisplayScan++)

                {



                    PropertyValueDisplayChar = pPropertyValueDisplay[ValueDisplayScan].ToString();

                    if (PropertyValueDisplayChar == " ")

                    {

                        break;

                    }

                    pPropertyValue = pPropertyValue + PropertyValueDisplayChar;

                }



                pPropertyUOM = pPropertyValueDisplay.Substring(ValueDisplayScan + 1); //040225;



                PropertyValue = pPropertyValue; //040225

                PropertyUOM = pPropertyUOM;

            }

            catch (Exception e)

            {

                throw e;

            }

        }

        private int PropertyKeyApply(string pPropertyName, string pPropertyIndex)
        {
            string PropertyInstance;
            int DesiredListIndex = -1;//Changed it from -1  
            int ActualListIndex;
            int PropertyKeyApply;


            PropertyInstance = Property_Instance_Unite(pPropertyName: pPropertyName, pPropertyIndex: pPropertyIndex);
            int.TryParse(PropertyInstance, out int result);

            foreach (DataGridViewRow row in dgvDocProperty.Rows)
            {
                if (row.Cells[qdpColPropertyInstance].Value != null && row.Cells[qdpColPropertyInstance].Value.ToString() == PropertyInstance)
                {
                    DesiredListIndex = row.Index;
                    break;
                }
                else
                    DesiredListIndex = dgvDocProperty.RowCount + 1; //added to add new rows
            }


            if (DesiredListIndex < dgvDocProperty.Rows.Count)
            {
                //DesiredListIndex is -1 so throwing error  
                if (dgvDocProperty.Rows[DesiredListIndex].Cells[qdpColPropertyInstance].Value.ToString().ToUpper() == PropertyInstance.ToString().ToUpper())
                {
                    ActualListIndex = DesiredListIndex;
                }
                else
                {

                    ActualListIndex = dgvDocProperty_AddItem(DesiredListIndex: DesiredListIndex);
                }
            }
            else
                ActualListIndex = dgvDocProperty_AddItem(DesiredListIndex: DesiredListIndex);
            {

                dgvDocProperty.Rows[ActualListIndex].Cells[qdpColPropertyName].Value = pPropertyName;
                dgvDocProperty.Rows[ActualListIndex].Cells[qdpColPropertyIndex].Value = pPropertyIndex;
                dgvDocProperty.Rows[ActualListIndex].Cells[qdpColPropertyInstance].Value = PropertyInstance;
            }

            PropertyKeyApply = ActualListIndex;
            return PropertyKeyApply;
        }



        private void PropertyDataApply(string propertyName, string propertyInstance, string propertyIndex, string PropertyValueRequired, string PropertyValue, string PropertyUOM, string PropertyValueDisplay, int PropertyMaterialID, int PropertyWorkCenterID)
        {
            if (qdpColCount != cintColCountExpected)
            {
                MessageBox.Show("EXPECTING " + cintColCountExpected + " NOT " + qdpColCount);
                return;
            }

            if (isNew)
            {
                DataRow dr = dtDocProperty.NewRow();
                dr["PropertyName"] = propertyName;
                dr["PropertyInstance"] = propertyInstance;
                dr["PropertyIndex"] = propertyIndex;
                dr["PropertyValueRequired"] = PropertyValueRequired;
                dr["PropertyValue"] = PropertyValue;
                dr["PropertyUOM"] = PropertyUOM;
                dr["PropertyValueDisplay"] = PropertyValueDisplay;
                dr["MaterialID"] = PropertyMaterialID;
                dr["WorkCenterID"] = PropertyWorkCenterID;

                if (mblnCalledWizard)//270125
                    dtDocProperty.Rows.Add(dr);
                else
                    dtDocProperty.Rows.InsertAt(dr, 0); //   

                dtDocProperty.AcceptChanges();
                dgvDocProperty.Refresh();
                //dgvDocProperty.Sort(dgvDocProperty.Columns["PropertyInstance"], ListSortDirection.Ascending); //030225
                /*dgvDocProperty.DataSource = null;
                dgvDocProperty.DataSource = dtDocProperty;*/ //030225
            }

            if (isUpdate)
            {
                //int i = dgvDocProperty.CurrentCell.RowIndex + 1;
                int i = dgvDocProperty.CurrentCell.RowIndex;
                dtDocProperty.Rows[i]["PropertyName"] = propertyName;
                dtDocProperty.Rows[i]["PropertyInstance"] = propertyInstance;
                dtDocProperty.Rows[i]["PropertyIndex"] = propertyIndex;
                dtDocProperty.Rows[i]["PropertyValueRequired"] = PropertyValueRequired;
                dtDocProperty.Rows[i]["PropertyValue"] = PropertyValue;
                dtDocProperty.Rows[i]["PropertyUOM"] = PropertyUOM;
                dtDocProperty.Rows[i]["PropertyValueDisplay"] = PropertyValueDisplay;
                dtDocProperty.Rows[i]["MaterialID"] = PropertyMaterialID;
                dtDocProperty.Rows[i]["WorkCenterID"] = PropertyWorkCenterID;
                dtDocProperty.AcceptChanges();
                dgvDocProperty.Refresh();
                /*dgvDocProperty.DataSource = null;
                dgvDocProperty.DataSource = dtDocProperty;*/ //030225
            }



            if (isNew)
                isNew = false;
            if (isUpdate)
                isUpdate = false;

        }

        private void btnPropertyDelete_Click(object sender, EventArgs e)
        {
            const string MyProcName = "cmdPropertyDelete_Click";
            int intListIndex;
            intListIndex = dgvDocProperty.CurrentCell.RowIndex;
            PropertyName = dgvDocProperty.Rows[intListIndex].Cells[qdpColPropertyName].Value.ToString();
            PropertyInstance = dgvDocProperty.Rows[intListIndex].Cells[qdpColPropertyInstance].Value.ToString();
            try
            {
                if (!Property_Delete_Valid(PropertyInstance: PropertyInstance, PropertyIndex: PropertyIndex))
                {
                    return;
                }
                if (PropertyUsage == "S")
                {
                    if (PropertyName == "FirstLotUsage")
                    {
                        DgvRow_Delete(pdgv_Delete: dgvDocProperty, pblnPrompt: true);

                    }
                    else
                    {
                        MessageBox.Show("'" + PropertyInstance + "' is a System Property." + "\r\n" + "System Properties cannot be deleted.", "System Property", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    DgvRow_Delete(pdgv_Delete: dgvDocProperty, pblnPrompt: true);

                }
                lblReqCount__FILL();
                PropertyReset();
                btnFinish.Enabled = true;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                }
            }
        }

        private bool Property_Delete_Valid(string PropertyInstance, string PropertyIndex)
        {

            bool Property_Delete_Valid = false;

            if (!TryResolveActiveDocumentForForm())
            {
                return false;
            }

            dtDocPropertyRequired = objMBRAuthDAL.DocPropertyRequired_SelectAll(mintOpUnitID, mintDocTypeID);
            if (!Property_Delete_Range_Valid(pstrPropertyName: PropertyInstance, prngToCheck: activeDocument.Range(), pstrDocumentPart: "BODY"))
            {
                return Property_Delete_Valid;
            }
            foreach (Word.Section sec in activeDocument.Sections)
            {
                foreach (Word.HeaderFooter phf in sec.Headers)
                {
                    if (!Property_Delete_HeaderFooter_Valid(pstrPropertyName: PropertyInstance, pphfHeaderFooter: phf, pstrDocumentPart: "HEADER"))
                    {
                        return Property_Delete_Valid;
                    }
                    else
                    {
                    }
                }

                foreach (Word.HeaderFooter phf in sec.Footers)
                {
                    if (!Property_Delete_HeaderFooter_Valid(pstrPropertyName: PropertyInstance, pphfHeaderFooter: phf, pstrDocumentPart: "FOOTER"))
                    {
                        return Property_Delete_Valid;
                    }
                    else
                    {
                    }
                }
            }

            if (objCommonService.IsPropertyRequired(PropertyInstance: PropertyInstance, dtReqProperty: dtDocPropertyRequired))
            {
                return Property_Delete_Valid;
            }
            return Property_Delete_Valid = true;
        }

        private bool Property_Delete_HeaderFooter_Valid(string pstrPropertyName, Word.HeaderFooter pphfHeaderFooter, string pstrDocumentPart)
        {
            bool Property_Delete_HeaderFooter_Valid = false;
            if (pphfHeaderFooter.LinkToPrevious == true)
            {
                Property_Delete_HeaderFooter_Valid = true;
                return Property_Delete_HeaderFooter_Valid;
            }

            Property_Delete_HeaderFooter_Valid = false;
            if (!Property_Delete_Range_Valid(pstrPropertyName: pstrPropertyName, prngToCheck: pphfHeaderFooter.Range, pstrDocumentPart: pstrDocumentPart))
            {
                return Property_Delete_HeaderFooter_Valid;
            }
            else
            {
            }

            return Property_Delete_HeaderFooter_Valid = true;
        }
        private bool Property_Delete_Range_Valid(string pstrPropertyName, Word.Range prngToCheck, string pstrDocumentPart)
        {

            string[] astrFieldWords;
            int intFieldWordScan;
            string strFieldPropertyName;
            string strErrorMessage;
            bool Property_Delete_Range_Valid = false;

            foreach (Word.Field fldScan in prngToCheck.Fields)
            {
                if (fldScan.Type != Word.WdFieldType.wdFieldDocProperty)
                {
                }
                else if (PropertyTable_Start < fldScan.Code.Start & PropertyTable_End > fldScan.Code.End)
                {
                }
                else
                {
                    try
                    {
                        astrFieldWords = fldScan.Code.Text.Trim().Split(' ');  
                    }
                    catch(Exception ex)
                    {
                        fldScan.Delete();
                        continue;
                    }
                    //astrFieldWords = fldScan.Code.ToString().Split(' ');
                    for (intFieldWordScan = astrFieldWords.GetLowerBound(0); intFieldWordScan <= astrFieldWords.GetUpperBound(0); intFieldWordScan++)
                    {
                        strFieldPropertyName = (astrFieldWords[intFieldWordScan]).ToUpper();
                        switch (strFieldPropertyName)
                        {
                            case "DOCPROPERTY":
                                {
                                    break;
                                }

                            case "":
                                {
                                    break;
                                }

                            default:
                                {
                                    if (strFieldPropertyName == pstrPropertyName.ToUpper())
                                    {
                                        strErrorMessage = "CANNOT DELETE - Property '" + pstrPropertyName + "' used in this Document " + pstrDocumentPart;
                                        MessageAdd(strErrorMessage);
                                        MessageBox.Show(strErrorMessage, "Cannot Delete Property", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        return Property_Delete_Range_Valid;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                    //break;
                                }
                        }
                    }
                }
            }
            Property_Delete_Range_Valid = true;

            return Property_Delete_Range_Valid;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            const string MyProcName = "cmdCancel_Click";
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (mblnCalledWizard)
                {
                    UserForm__SHUTDOWN(pqrtQReturn: qReturn.qReturn_Cancel);
                }
                else
                {
                    this.Close();
                }
                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                }
            }
            finally
            {
              Cursor = Cursors.Default;
              CommonService.CleanUpBackgroundWordProcesses();
                
            }
        }
        private void btnFinish_Click(object sender, EventArgs e)
        {
            const string MyProcName = "cmdFinish_Click";
            Cursor.Current = Cursors.WaitCursor;
            bool btnWordUpdate = false ;
            try
            {
                var sortedRows = dtDocProperty.AsEnumerable().OrderBy(row => row.Field<string>("PropertyInstance"), StringComparison.OrdinalIgnoreCase.WithNaturalSort()).ToList();

                DataTable sortedTable = dtDocProperty.Clone();
                foreach (var row in sortedRows)
                    sortedTable.ImportRow(row);

                dtDocProperty = sortedTable;
                dgvDocProperty.DataSource = dtDocProperty;
                dgvDocProperty.Refresh();

                if (btnFinish.Text == "Close")
                {
                    //Dissolved function PerformIntroPagesAction
                    mblnCalledWizard = true;
                    Define_DocumentProperties();
                    Setup_Database(true, dgvDocProperty);
                    if (blnDBUpdateSuccessful != true)
                    {
                        MessageBox.Show("Data save operation failed. Try again.", "Failed to Save", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    //added active doc calling part after db save
                    if (!TryResolveActiveDocumentForForm())
                    {
                        // throw new InvalidOperationException("Unable to access the target document, Close the form and Save again.");
                        MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        this.Close();
                        return;
                    }
                    ExecuteWordComWithRetry(() => activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorWait);
                    if (ExecuteWordComWithRetry(() => activeDocument.ProtectionType != Word.WdProtectionType.wdNoProtection)) //210125 RPC
                    {
                        Document_UnProtect();
                    }
                    this.Left = System.Windows.Forms.Application.OpenForms[0].Width - tabEditProperty.Left; //AfterComp

                    Setup_Document_Property();
                    Setup_Document_Text(pblnReplaceTable: false);
                    UpdateWordDocumentFields();
                    UserForm__SHUTDOWN(pqrtQReturn: qReturn.qReturn_Finish);
                    return;
                }
                else
                {
                    Setup_Database(Update: true, dgvOriginal: dgvDocProperty);
                    if (blnDBUpdateSuccessful != true)
                    {
                        MessageBox.Show("Data save operation failed. Try again.", "Failed to Save", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (!TryResolveActiveDocumentForForm())
                    {
                        btnWordUpdate = false;
                        //throw new InvalidOperationException("Unable to access the target document, Click on Refresh menu to update the document.");
                        if (blnDBUpdateSuccessful)
                        {

                            MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            this.Close();
                            return;
                        }
                        else
                        {
                            MessageBox.Show("Data save operation failed. Try again.", "Failed to Save", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    ExecuteWordComWithRetry(() => activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorWait);
                    if (ExecuteWordComWithRetry(() => activeDocument.ProtectionType != Word.WdProtectionType.wdNoProtection)) //210125 RPC
                    {
                        Document_UnProtect();
                    }
                    this.Left = System.Windows.Forms.Application.OpenForms[0].Width - tabEditProperty.Left; //AfterComp
                }

                

                if (!TryResolveActiveDocumentForForm())
                {
                    // throw new InvalidOperationException("Unable to access the target document, Close the form and Save again.");
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }

                Setup_Document_Property();
                tabProperty = CommonService.Table_Find(activeDocument, "Property List");
                if (tabProperty == null)
                {
                    MessageBox.Show("Property TABLE NOT FOUND - NO TABLE REPLACEMENT");
                }
                else
                {
                    Setup_Document_Text(pblnReplaceTable: true);
                }

                if (ExecuteWordComWithRetry(() => activeDocument.ProtectionType != Word.WdProtectionType.wdNoProtection))
                {
                    Document_UnProtect();
                }
                if (ExecuteWordComWithRetry(() => activeDocument.TrackRevisions == true))
                {
                    ExecuteWordComWithRetry(() => activeDocument.TrackRevisions = false);
                }
                UpdateWordDocumentFields(); //220125
                if (mblnDocTrackChgsReq == true)
                {
                    if (ExecuteWordComWithRetry(() => activeDocument.TrackRevisions == false))
                    {
                        ExecuteWordComWithRetry(() => activeDocument.TrackRevisions = true);
                    }
                }
                if (!TryResolveActiveDocumentForForm())
                {
                    //throw new InvalidOperationException("Unable to access the target document, Close the form and Save again.");
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }
                SaveManager.Document_Save_FileAndDatabase(pdocCurrent: activeDocument, docpasswordProtect: mstrProtectionPassword, pblnRestoreCommandBar: false);
                if (!TryResolveActiveDocumentForForm())
                {
                    //throw new InvalidOperationException("Unable to access the target document, Close the form and Save again.");
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }
                CommonService.DocumentProtection_Assure(pdocProtectedDocument: activeDocument, pstrProtectionPassword: mstrProtectionPassword);
                btnWordUpdate = true;
                Cursor.Current = Cursors.Default;
                ExecuteWordComWithRetry(() => activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal);               

                this.Close();
            }
            catch (Exception ex)
            {
                try
                {
                    if (TryResolveActiveDocumentForForm() && activeDocument?.Application != null)
                    {
                        activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;
                    }
                }
                catch
                {
                }

                Cursor = Cursors.Default;

                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                }
            }
            finally
            {
                /*activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;
                Cursor = Cursors.Default;*/
                //performancetuning
                StopWordKeepAliveTimer();
                if (blnDBUpdateSuccessful) { 

                    CommonService.CleanUpBackgroundWordProcesses();
            }
            }
        }

        private bool Valid_Entries()
        {
            int intListScan;
            bool ValidEntry = false;
            mintParmsValidated = 0;
            //for (intListScan = 0; intListScan <= dgvDocProperty.Rows.Count - 1; intListScan++)
            for (intListScan = 0; intListScan < dgvDocProperty.Rows.Count; intListScan++) //   
            {
                //dgvDocProperty.Rows[intListScan].Selected = true;
                dgvDocProperty.CurrentCell = dgvDocProperty.Rows[intListScan].Cells[0];
                DocPropertySelected();
                if (!Property_Valid(ppvmMode: PropertyValidationMode.PropertyValidation_Update, vPropertyInstance: PropertyInstance, vPropertyValue: PropertyValue, vPropertyValueDisplay: PropertyValueDisplay, vPropertyMaterialID: PropertyMaterialID, vPropertyWorkCenterID: PropertyWorkCenterID))
                //if (!Property_Valid(ppvmMode: PropertyValidationMode.PropertyValidation_Update, PropertyInstance: ref PropertyInstance, PropertyValue: ref PropertyValue, PropertyValueDisplay: ref PropertyValueDisplay, PropertyMaterialID: PropertyMaterialID, PropertyWorkCenterID : PropertyWorkCenterID))
                {
                    return ValidEntry;
                }
            }

            ValidEntry = true;
            mintParmsValidated = cintParmsValidated;
            return ValidEntry;
        }
        private void Document_Protect()
        {
            if (!TryResolveActiveDocumentForForm())
            {
                MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.Close();
                return;
            }

            CommonService.DocumentProtection_On(pdocProtectedDocument: activeDocument, pstrProtectionPassword: mstrProtectionPassword);
        }
        private void Document_UnProtect()
        {
            if (!TryResolveActiveDocumentForForm())
            {
                // throw new InvalidOperationException("Unable to resolve the target document.");
                
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                
            }

            CommonService.DocumentProtection_Off(pdocProtectedDocument: activeDocument, pstrProtectionPassword: mstrProtectionPassword);
        }

        private string left(string strCellText, int len)
        {
            string strText = strCellText.Substring(0, Math.Min(strCellText.Length, len));
            return strText;
        }
        public static int? NullIfNegative(int value)
        {
            return value < 0 ? (int?)null : value;
        }

        private const uint RpcCallRejectedErrorCode = 0x80010001;
        private const uint RpcServerUnavailableErrorCode = 0x800706BE;
        private const uint RpcDisconnectedErrorCode = 0x80010108;
        private const uint RpcSystemCallFailedErrorCode = 0x80010100;
        private const uint RpcServerBusyRetryLaterErrorCode = 0x8001010A;
        private const uint RpcServerUnavailableAltErrorCode = 0x800706BA;
        private const int WordComRetryCount = 6;
        private const int WordComRetryDelayMs = 150;

        private static bool IsTransientWordComException(COMException ex)
        {
            uint code = (uint)ex.ErrorCode;
            return code == RpcCallRejectedErrorCode ||
                   code == RpcServerUnavailableErrorCode ||
                   code == RpcDisconnectedErrorCode ||
                   code == RpcSystemCallFailedErrorCode ||
                   code == RpcServerBusyRetryLaterErrorCode ||
                   code == RpcServerUnavailableAltErrorCode;
        }

        private static void ExecuteWordComWithRetry(Action action, int retries = WordComRetryCount, int delayMs = WordComRetryDelayMs)
        {
            for (int attempt = 0; attempt < retries; attempt++)
            {
                try
                {
                    action();
                    return;
                }
                catch (COMException ex) when (IsTransientWordComException(ex))
                {
                    if (attempt == retries - 1)
                    {
                        throw;
                    }

                    System.Threading.Thread.Sleep(delayMs);
                }
            }
        }

        private static T ExecuteWordComWithRetry<T>(Func<T> action, int retries = WordComRetryCount, int delayMs = WordComRetryDelayMs)
        {
            for (int attempt = 0; attempt < retries; attempt++)
            {
                try
                {
                    return action();
                }
                catch (COMException ex) when (IsTransientWordComException(ex))
                {
                    if (attempt == retries - 1)
                    {
                        throw;
                    }

                    System.Threading.Thread.Sleep(delayMs);
                }
            }

            throw new InvalidOperationException("Failed to execute Word COM call after retries.");
        }

        private void Setup_Database(bool Update, DataGridView dgvOriginal)
        {
            const string MyProcName = "Setup_Database";
            int intMain;
            string PropertyInstanceList;
            int RecordsAffected;
            bool MyTransaction;
            string strPropertyUOM;
            MyTransaction = false;
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            try
            {
                if (Update)
                {
                    if (DatabaseHelper.Transaction_Active() == "")
                    {
                        DatabaseHelper.Transaction_Begin(mstrDocName);
                        MyTransaction = true;
                    }
                }
                else
                {
                    MessageBox.Show("unexpected call with NO-Update", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                //int dtCount = dtDocProperty.Rows.Count; // Gow11Aug
                //int dgvDocPropertyCount = dgvDocProperty.Rows.Count; // Gow11Aug

                //for (intMain = 0; intMain <= dgvDocProperty.Rows.Count - 1; intMain++)
                //{
                //    //DataRow row = dtDocProperty.Rows[intMain];

                //    //if(row.RowState == DataRowState.Unchanged)
                //    //    continue;

                //    if (dgvDocProperty.Rows[intMain].Cells[qdpColPropertyUOM].Value != null)
                //    {
                //        strPropertyUOM = dgvDocProperty.Rows[intMain].Cells[qdpColPropertyUOM].Value.ToString();
                //    }
                //    else
                //    {
                //        strPropertyUOM = "";
                //    }

                //    if (!objMBRAuthDAL.DocProperty_Apply(DocID: System.Convert.ToInt32(mDocID),
                //                           PropertyName: dgvDocProperty.Rows[intMain].Cells[qdpColPropertyName].Value.ToString(),
                //                           PropertyIndex: System.Convert.ToInt32(dgvDocProperty.Rows[intMain].Cells[qdpColPropertyIndex].Value.ToString()),
                //                           PropertyValue: dgvDocProperty.Rows[intMain].Cells[qdpColPropertyValue].Value.ToString(),
                //                           PropertyUOM: strPropertyUOM,
                //                           MaterialID: System.Convert.ToInt32(dgvDocProperty.Rows[intMain].Cells[qdpColMaterialID].Value.ToString()),
                //                           WorkCenterID: System.Convert.ToInt32(dgvDocProperty.Rows[intMain].Cells[qdpColWorkCenterID].Value.ToString())))
                //    {
                //        DatabaseHelper.Transaction_Rollback(mstrDocName);
                //        blnDBUpdateSuccessful = false;
                //    }

                //}

                // prod
                for (intMain = 0; intMain <= dgvDocProperty.Rows.Count - 1; intMain++)
                {
                    if (dgvDocProperty.Rows[intMain].Cells[qdpColPropertyUOM].Value != null)
                    {
                        strPropertyUOM = dgvDocProperty.Rows[intMain].Cells[qdpColPropertyUOM].Value.ToString();
                    }
                    else
                    {
                        strPropertyUOM = "";
                    }

                    if (!objMBRAuthDAL.DocProperty_Apply(DocID: System.Convert.ToInt32(mDocID),
                                           PropertyName: dgvDocProperty.Rows[intMain].Cells[qdpColPropertyName].Value.ToString(),
                                           PropertyIndex: System.Convert.ToInt32(dgvDocProperty.Rows[intMain].Cells[qdpColPropertyIndex].Value.ToString()),
                                           PropertyValue: dgvDocProperty.Rows[intMain].Cells[qdpColPropertyValue].Value.ToString(),
                                           PropertyUOM: strPropertyUOM,
                                           MaterialID: System.Convert.ToInt32(dgvDocProperty.Rows[intMain].Cells[qdpColMaterialID].Value.ToString()),
                                           WorkCenterID: System.Convert.ToInt32(dgvDocProperty.Rows[intMain].Cells[qdpColWorkCenterID].Value.ToString())))
                    {
                        DatabaseHelper.Transaction_Rollback(mstrDocName);
                        blnDBUpdateSuccessful = false;
                    }

                }
                if (dgvDocProperty.Rows.Count == 0)
                {
                    PropertyInstanceList = "('Delete all!')";
                }
                else
                {
                    PropertyInstanceList = "(";
                    for (intMain = 0; intMain <= dgvDocProperty.Rows.Count - 1; intMain++)
                    {
                        PropertyInstanceList = PropertyInstanceList + "'" + dgvDocProperty.Rows[intMain].Cells[qdpColPropertyInstance].Value + "',"; //200125

                    }
                    PropertyInstanceList = left(PropertyInstanceList, PropertyInstanceList.Length - 1) + ")";
                }
                objMBRAuthDAL.DocProperty_DeleteAllExcept(mDocID, PropertyInstanceList: PropertyInstanceList);
                if (MyTransaction)
                {
                    DatabaseHelper.Transaction_Commit(mstrDocName);
                    blnDBUpdateSuccessful = true;
                }
                else
                    blnDBUpdateSuccessful = false;


            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                    DatabaseHelper.Transaction_Rollback(mstrDocName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                    DatabaseHelper.Transaction_Rollback(mstrDocName);
                }
            }

        }
        private void Setup_Document_Text(bool pblnReplaceTable)
        {
            const string MyProcName = "Setup_Document_Text";
            int intTableRow = 0;
            int DocProperty_ListCount;
            int DocProperty_ListIndex;
            int TableRow_Count;
            //int TableRow_Index = -1; COMMENTED BY SOUNDAHR
            int TableRow_Index = 3;
            string PropertyName;
            const int cintColPropertyName = 1;
            const int cintColPropertyValue = 2;
            bool ReApplyProtection = false;
            string propertyValue = string.Empty;

            try
            {
                if (tabProperty == null) //060125
                {
                    tabProperty = CommonService.Table_Find(ExecuteWordComWithRetry(() => activeDocument), "Property List");
                }

                if (pblnReplaceTable)
                {
                    Document_UnProtect();
                    ReApplyProtection = true;
                    CommonService.Table_Delete(ptblToDelete: tabProperty, pdocCurrent: ExecuteWordComWithRetry(() => activeDocument), pblnSelect: true);
                }
                else
                {
                    ReApplyProtection = false;
                }
                DocProperty_ListCount = dgvDocProperty.Rows.Count;
                if (dgvDocProperty.Rows.Count > 0)
                {
                    TableRow_Count = DocProperty_ListCount;
                }
                else
                {
                    TableRow_Count = 1;
                }
                //CommonService.AutoText_Insert("ModuleList", ref intTableRow, intModuleCount, tableAutoText: tblModule, wordApp: activeDocument.Application);
                //CommonService.AutoText_Insert(autotextEntryName: "PropertyList", tableDataRows: TableRow_Count, tableDataRowFirst: ref TableRow_Index, tableAutoText: tabProperty, wordApp:activeDocument.Application);
                CommonService.AutoText_Insert("PropertyList", ref intTableRow, TableRow_Count, tableAutoText: tabProperty, wordApp: activeDocument.Application);
                {
                    tabProperty = CommonService.Table_Find_PropertyMBR(activeDocument);
                    tabProperty.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                }
                /*if (mblnCalledWizard)
                {

                    for (DocProperty_ListIndex = 1; DocProperty_ListIndex <= DocProperty_ListCount; DocProperty_ListIndex++)
                    {
                        //PropertyName = dgvDocProperty.Rows[DocProperty_ListIndex].Cells[qdpColPropertyName].Value.ToString();
                        PropertyName = dgvDocProperty.Rows[DocProperty_ListIndex].Cells[qdpColPropertyInstance].Value.ToString(); //qdp100125
                        tabProperty.Cell(TableRow_Index, cintColPropertyName).Range.Text = PropertyName;
                        tabProperty.Cell(TableRow_Index, cintColPropertyValue).Range.Text = " ";
                        CommonService.Field_InsertProperty(where: tabProperty.Cell(TableRow_Index, cintColPropertyValue).Range, propertyName: PropertyName);
                        TableRow_Index = TableRow_Index + 1;
                    }
                }
                else
                {*/
                for (DocProperty_ListIndex = 0; DocProperty_ListIndex <= DocProperty_ListCount - 1; DocProperty_ListIndex++)
                {
                    //PropertyName = dgvDocProperty.Rows[DocProperty_ListIndex].Cells[qdpColPropertyName].Value.ToString();
                    PropertyName = dgvDocProperty.Rows[DocProperty_ListIndex].Cells[qdpColPropertyInstance].Value.ToString();//qdp100125
                    tabProperty.Cell(TableRow_Index, cintColPropertyName).Range.Text = PropertyName;
                    tabProperty.Cell(TableRow_Index, cintColPropertyValue).Range.Text = " ";
                    CommonService.Field_InsertProperty(where: tabProperty.Cell(TableRow_Index, cintColPropertyValue).Range, propertyName: PropertyName);
                    TableRow_Index = TableRow_Index + 1;
                }
                //}

                if (ReApplyProtection)//comparison  
                {
                    Document_Protect();
                }
            }

            catch (Exception ex)
            {
                if (ReApplyProtection)
                {
                    Document_Protect();
                }
                else
                {
                }
            }
        }

        // works, but takes more time
        //private void Setup_Document_Property()
        //{
        //    const string MyProcName = "Setup_Document_Property";
        //    int intListBoxPropertyCount;
        //    int intListBoxPropertyScan;
        //    DocProperty[] adpyListBoxProperty;
        //    string strListBoxPropertyInstance;
        //    string strListBoxPropertyValue;
        //    int intListBoxAllIndex;
        //    int intDocumentPropertyCount;
        //    int intDocumentPropertyScan;
        //    DocProperty[] adpyDocumentProperty;
        //    string strDocumentPropertyName;
        //    string strDocumentPropertyValue;
        //    //Microsoft.Office.Core.DocumentProperties dpcDocProperties;
        //    dynamic dpcDocProperties;

        //    int intListScan;
        //    try
        //    {
        //        adpyListBoxProperty = new DocProperty[dgvDocProperty.Rows.Count]; //COmp
        //        intListBoxPropertyCount = 0;
        //        for (intListScan = 0; intListScan <= dgvDocProperty.Rows.Count - 1; intListScan++)
        //        {
        //            Array_Add(parrWork: ref adpyListBoxProperty, pintArrayMax: ref intListBoxPropertyCount, pstrPropertyName: dgvDocProperty.Rows[intListScan].Cells[qdpColPropertyInstance].Value.ToString(), pstrPropertyValue: dgvDocProperty.Rows[intListScan].Cells[qdpColPropertyValueDisplay].Value.ToString());
        //        }

        //        dpcDocProperties = activeDocument.CustomDocumentProperties; //RPC Exception 3rd tab.
        //        //dpcDocProperties = Globals.ThisDocument.ActiveDocument.CustomDocumentProperties;

        //        while (dpcDocProperties.Count > 0)
        //        {
        //            dpcDocProperties.Item(1).Delete();
        //        }

        //        //var originalSectionStatus = SaveOriginalLinkStatus(activeDocument.Application);
        //        //LinkAllSections(activeDocument.Application);
        //        if (mblnCalledWizard)
        //        {
        //            foreach (DocProperty property in adpyListBoxProperty)
        //            {
        //                if (property.PropertyName.ToUpper() == "CHANGENUM")
        //                {
        //                    CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, CommonService.cChangeNum_ValueDefault, activeDocument);//220125
        //                    continue;
        //                }

        //                CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, property.PropertyValue, activeDocument);
        //            }
        //        }
        //        else
        //        {
        //            foreach (DocProperty property in adpyListBoxProperty)
        //            {
        //                CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, property.PropertyValue, activeDocument);
        //            }
        //        }

        //        UpdateWordDocumentFields(); //220125

        //        /*foreach (Word.Section sec in activeDocument.Sections)
        //        {
        //            foreach (Word.Field aField in sec.Range.Fields)
        //            {
        //                aField.Update();
        //            }
        //            foreach (Word.HeaderFooter header in sec.Headers)
        //            {
        //                foreach (Word.Field aField in header.Range.Fields)
        //                {
        //                    aField.Update();
        //                }
        //            }
        //            foreach (Word.HeaderFooter footer in sec.Footers)
        //            {
        //                foreach (Word.Field aField in footer.Range.Fields)
        //                {
        //                    aField.Update();
        //                }
        //            }
        //        }*/

        //    }

        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message.ToString());
        //    }
        //}

        // To improve the performance
        private void Setup_Document_Property()
        {
            const string MyProcName = "Setup_Document_Property";
            int intListBoxPropertyCount;
            int intListBoxPropertyScan;
            DocProperty[] adpyListBoxProperty;
            string strListBoxPropertyInstance;
            string strListBoxPropertyValue;
            int intListBoxAllIndex;
            int intDocumentPropertyCount;
            int intDocumentPropertyScan;
            DocProperty[] adpyDocumentProperty;
            string strDocumentPropertyName;
            string strDocumentPropertyValue;
            //Microsoft.Office.Core.DocumentProperties dpcDocProperties;
            dynamic dpcDocProperties;

            int intListScan;
            try
            {
                if (!TryResolveActiveDocumentForForm())
                {
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }

                dpcDocProperties = activeDocument.CustomDocumentProperties;

                while (dpcDocProperties.Count > 0)
                {
                    dpcDocProperties.Item(1).Delete();
                }
                adpyListBoxProperty = new DocProperty[dgvDocProperty.Rows.Count]; //COmp
                intListBoxPropertyCount = 0;

                for (intListScan = 0; intListScan <= dgvDocProperty.Rows.Count - 1; intListScan++)
                {
                    Array_Add(parrWork: ref adpyListBoxProperty, pintArrayMax: ref intListBoxPropertyCount, pstrPropertyName: dgvDocProperty.Rows[intListScan].Cells[qdpColPropertyInstance].Value.ToString(), pstrPropertyValue: dgvDocProperty.Rows[intListScan].Cells[qdpColPropertyValueDisplay].Value.ToString());
                }

                if (mblnCalledWizard)
                {
                    foreach (DocProperty property in adpyListBoxProperty)
                    {
                        if (property.PropertyName.ToUpper() == "CHANGENUM")
                        {
                            CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, CommonService.cChangeNum_ValueDefault, activeDocument);//220125
                            continue;
                        }

                        CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, property.PropertyValue, activeDocument);
                    }
                }
                else
                {
                    foreach (DocProperty property in adpyListBoxProperty)
                    {
                        CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, property.PropertyValue, activeDocument);
                    }
                }
                UpdateWordDocumentFields(); //220125
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        //private void Setup_Document_Property()
        //{
        //    const string MyProcName = "Setup_Document_Property";
        //    int intListBoxPropertyCount;
        //    int intListBoxPropertyScan;
        //    DocProperty[] adpyListBoxProperty;
        //    string strListBoxPropertyInstance;
        //    string strListBoxPropertyValue;
        //    int intListBoxAllIndex;
        //    int intDocumentPropertyCount;
        //    int intDocumentPropertyScan;
        //    DocProperty[] adpyDocumentProperty;
        //    string strDocumentPropertyName;
        //    string strDocumentPropertyValue;
        //    //Microsoft.Office.Core.DocumentProperties dpcDocProperties;
        //    dynamic dpcDocProperties =null;
        //    int retries = 3;
        //    int delayMs = 200;
        //    Word.Document doc = null;

        //    int intListScan;
        //    try
        //    {
        //        for (int i = 0; i < retries; i++)
        //        {
        //            try
        //            {
        //                doc = GetActiveDocument();
        //                if (doc != null)
        //                {
        //                    dpcDocProperties = activeDocument.CustomDocumentProperties;
        //                    break;
        //                }
        //            }
        //            catch (System.Runtime.InteropServices.COMException comEx)
        //            {
        //                Debug.WriteLine($"Attempt {i + 1}: COMException - {comEx.Message}");
        //                System.Threading.Thread.Sleep(delayMs);
        //            }


        //            while (dpcDocProperties.Count > 0)
        //        {
        //            dpcDocProperties.Item(1).Delete();
        //        }
        //        adpyListBoxProperty = new DocProperty[dgvDocProperty.Rows.Count]; //COmp
        //        intListBoxPropertyCount = 0;

        //        for (intListScan = 0; intListScan <= dgvDocProperty.Rows.Count - 1; intListScan++)
        //        {
        //            Array_Add(parrWork: ref adpyListBoxProperty, pintArrayMax: ref intListBoxPropertyCount, pstrPropertyName: dgvDocProperty.Rows[intListScan].Cells[qdpColPropertyInstance].Value.ToString(), pstrPropertyValue: dgvDocProperty.Rows[intListScan].Cells[qdpColPropertyValueDisplay].Value.ToString());
        //        }

        //        if (mblnCalledWizard)
        //        {
        //            foreach (DocProperty property in adpyListBoxProperty)
        //            {
        //                if (property.PropertyName.ToUpper() == "CHANGENUM")
        //                {
        //                    CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, CommonService.cChangeNum_ValueDefault, activeDocument);//220125
        //                    continue;
        //                }

        //                CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, property.PropertyValue, activeDocument);
        //            }
        //        }
        //        else
        //        {
        //            foreach (DocProperty property in adpyListBoxProperty)
        //            {
        //                CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, property.PropertyValue, activeDocument);
        //            }
        //        }
        //        UpdateWordDocumentFields(); //220125
        //    }

        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message.ToString());
        //    }
        //}

        //private Word.Document GetActiveDocument()
        //{
        //    try
        //    {
        //        return Globals.ThisDocument.ActiveDocument;
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine("Error getting ActiveDocument: " + ex.Message);
        //        return null;
        //    }
        //}

        /*private void Setup_Document_Property()
        {
            const string MyProcName = "Setup_Document_Property";
            int intListBoxPropertyCount;
            int intListBoxPropertyScan;
            DocProperty[] adpyListBoxProperty;
            string strListBoxPropertyInstance;
            string strListBoxPropertyValue;
            int intListBoxAllIndex;
            int intDocumentPropertyCount;
            int intDocumentPropertyScan;
            DocProperty[] adpyDocumentProperty;
            string strDocumentPropertyName;
            string strDocumentPropertyValue;
            Microsoft.Office.Core.DocumentProperties dpcDocProperties;
            //dynamic dpcDocProperties;
            

            int intListScan;
            try
            {
                adpyListBoxProperty = new DocProperty[dgvDocProperty.Rows.Count]; //COmp
                intListBoxPropertyCount = 0;
                for (intListScan = 0; intListScan <= dgvDocProperty.Rows.Count - 1; intListScan++)
                {
                    Array_Add(parrWork: ref adpyListBoxProperty, pintArrayMax: ref intListBoxPropertyCount, pstrPropertyName: dgvDocProperty.Rows[intListScan].Cells[qdpColPropertyInstance].Value.ToString(), pstrPropertyValue: dgvDocProperty.Rows[intListScan].Cells[qdpColPropertyValueDisplay].Value.ToString());
                }

                dpcDocProperties = activeDocument.CustomDocumentProperties;
                adpyDocumentProperty = new DocProperty[dpcDocProperties.Count];

                intDocumentPropertyCount = 0;
                foreach (Microsoft.Office.Core.DocumentProperties dpyDocProperty in dpcDocProperties)
                {
                    Array_Add(parrWork: ref adpyDocumentProperty, pintArrayMax: ref intDocumentPropertyCount, pstrPropertyName: dpyDocProperty[0].Value, pstrPropertyValue: dpyDocProperty[1].Value);
                }
                intListBoxPropertyScan = 1;
                intDocumentPropertyScan = 1;

                
                while ((intListBoxPropertyScan <= intListBoxPropertyCount && intDocumentPropertyScan <= intDocumentPropertyCount))
                {
                    if (intListBoxPropertyScan > intListBoxPropertyCount)
                    {
                        strListBoxPropertyInstance = ((char)255).ToString();
                        strListBoxPropertyValue = "AFTEREOA";
                    }
                    else
                    {
                        strListBoxPropertyInstance = adpyListBoxProperty[intListBoxPropertyScan].PropertyName;
                        strListBoxPropertyValue = adpyListBoxProperty[intListBoxPropertyScan].PropertyValue;
                    }

                    if (intDocumentPropertyScan > intDocumentPropertyCount)
                    {
                        strDocumentPropertyName = ((char)255).ToString();
                        strDocumentPropertyValue = "AFTEREOA";
                    }
                    else
                    {
                        strDocumentPropertyName = adpyDocumentProperty[intDocumentPropertyScan].PropertyName;
                        strDocumentPropertyValue = adpyDocumentProperty[intDocumentPropertyScan].PropertyValue;
                    }

                    
                    if (string.Compare(strListBoxPropertyInstance, strDocumentPropertyName, true) == 0)
                    {
                        if (strListBoxPropertyValue == strDocumentPropertyValue)
                        {
                        }
                        else
                        {
                            CommonService.DocumentProperty_Update(pdpcDocProperties: dpcDocProperties, pstrDocumentPropertyName: strDocumentPropertyName, pstrDocumentPropertyValue: strListBoxPropertyValue);
                        }
                        intListBoxPropertyScan = intListBoxPropertyScan + 1;
                        intDocumentPropertyScan = intDocumentPropertyScan + 1;
                    }
                    
                    else if (string.Compare(strListBoxPropertyInstance, strDocumentPropertyName, true) < 0)
                    {
                        CommonService.DocumentProperty_Add(dpcDocProperties, DocumentPropertyName: strListBoxPropertyInstance, DocumentPropertyValue: strListBoxPropertyValue);
                        //CommonService.DocumentProperty_PropertyAdd(DocumentPropertyName: strListBoxPropertyInstance, DocumentPropertyValue: strListBoxPropertyValue, document: activeDocument);
                        intListBoxPropertyScan = intListBoxPropertyScan + 1;
                    }
                    
                    else if (string.Compare(strListBoxPropertyInstance, strDocumentPropertyName, true) > 0)
                    {
                        intListBoxAllIndex = DataGridView_Find(dgvControl: dgvDocProperty, intKeyColumn: qdpColPropertyInstance, strKeyValue: strDocumentPropertyName);
                        if (intListBoxAllIndex < 0)
                        {
                            CommonService.DocumentProperty_Delete(pdpcDocProperties: dpcDocProperties, pstrDocumentPropertyName: strDocumentPropertyName);
                        }
                        else
                        {
                        }

                        intDocumentPropertyScan = intDocumentPropertyScan + 1;
                    }
                    else
                    {
                        //Todo Information.Err.Raise(Number: 42909, Source: MyProcName, Description: "FAILED COMPARE: '" + strListBoxPropertyInstance + "' AND '" + strDocumentPropertyName + "'");
                        var ex = new Exception();
                        ex.Data["errNum"] = "42909";
                        ex.Data["errDesc"] = "FAILED COMPARE: '" + strListBoxPropertyInstance + "' AND '" + strDocumentPropertyName + "'";
                        throw ex;
                    }

                }
                return;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }*/

        // works, but takes more time
        //private void Array_Add(ref DocProperty[] parrWork, ref int pintArrayMax, string pstrPropertyName, string pstrPropertyValue)
        //{
        //    int intScan;
        //    int intArrayScan;

        //    for (intArrayScan = pintArrayMax; intArrayScan >= 1; intArrayScan--)
        //    {
        //        if (intArrayScan <= parrWork.Count())
        //        {
        //            if (string.Compare(parrWork[intArrayScan].PropertyName, pstrPropertyName) > 0)
        //            {
        //                parrWork[intArrayScan + 1] = parrWork[intArrayScan];
        //            }
        //            else
        //            {
        //                break;
        //            }
        //        }
        //    }

        //    //intArrayScan = intArrayScan + 1;
        //    intArrayScan = intArrayScan;

        //    if (intArrayScan <= parrWork.Count())
        //    {
        //        parrWork[intArrayScan].PropertyName = pstrPropertyName;
        //        parrWork[intArrayScan].PropertyValue = pstrPropertyValue;
        //        pintArrayMax = pintArrayMax + 1;
        //    }
        //    /*parrWork[intArrayScan].PropertyName = pstrPropertyName;
        //    parrWork[intArrayScan].PropertyValue = pstrPropertyValue;
        //    pintArrayMax = pintArrayMax + 1;*/
        //    /*int intScan;
        //    int intArrayScan;
        //    for (intArrayScan = pintArrayMax; intArrayScan >= 1; intArrayScan += -1)            
        //    {

        //        if (string.Compare(parrWork[intArrayScan].PropertyName, pstrPropertyName) > 0)
        //        {
        //            parrWork[intArrayScan + 1] = parrWork[intArrayScan];
        //        }
        //        else
        //        {
        //            break;
        //        }
        //    }

        //    intArrayScan = intArrayScan + 1;
        //    parrWork[intArrayScan].PropertyName = pstrPropertyName;
        //    parrWork[intArrayScan].PropertyValue = pstrPropertyValue;
        //    pintArrayMax = pintArrayMax + 1;*/
        //}

        // To improve the performance
        private void Array_Add(ref DocProperty[] parrWork, ref int pintArrayMax, string pstrPropertyName, string pstrPropertyValue)
        {
            //To improve performance 22/04/25
            if (pintArrayMax <= parrWork.Count())
            {
                parrWork[pintArrayMax].PropertyName = pstrPropertyName;
                parrWork[pintArrayMax].PropertyValue = pstrPropertyValue;
                pintArrayMax = pintArrayMax + 1;
            }

        }

        private void cmdPropertyValueSize_Click(object sender, EventArgs e)
        {
            bool blnPropertyValue_IsStandard;
            blnPropertyValue_IsStandard = (txtPropertyValue.Left == sglPropertyValue_Left);
            txtPropertyValue.MaxLength = PropertyTextMaxLen;
            txtPropertyValue_Large(blnPropertyValue_IsStandard);
            txtPropertyValue.Focus();
        }

        private void txtPropertyValue_Large(bool blnLarge)
        {
            bool blnOverlaidVisible;
            if (blnLarge)
            {
                blnOverlaidVisible = false;
            }
            else if (UserForm__MODE_Add())
            {
                blnOverlaidVisible = true;
            }
            else
            {
                blnOverlaidVisible = false;
            }
            lblUpdateMessage.Text = "";
            DocPropertyList_Visible(blnOverlaidVisible);


            txtPropertyValue.Multiline = blnLarge;

            //txtPropertyValue.EnterKeyBehavior = blnLarge;
            txtPropertyValue.MaxLength = PropertyTextMaxLen;


            if (blnLarge)
            {
                btnPropertyAdd.Top = (int)(tabEditProperty.Height - btnPropertyAdd.Height - sglMargin);
                btnPropertyUpdate.Top = (int)(tabEditProperty.Height - btnPropertyUpdate.Height - sglMargin);
                btnDescribe.Top = (int)(tabEditProperty.Height - btnDescribe.Height - sglMargin);
                btnPropertyDelete.Top = (int)(tabEditProperty.Height - btnPropertyDelete.Height - sglMargin);
                btnPropertyDelete.Visible = false;
                txtPropertyValue.Left = (int)(sglMargin);
                txtPropertyValue.Width = (int)(tabEditProperty.Width - (sglMargin * 4));
                txtPropertyValue.Top = (int)(lblPropertyValue.Top + lblPropertyValue.Height + (sglMargin * 2));
                txtPropertyValue.Height = (int)(btnPropertyUpdate.Top - txtPropertyValue.Top - sglMargin);
                lblSpace.Top = 110;
                lblSpaceCount.Top = 110;
                lblSpace.Visible = true;
                lblSpaceCount.Visible = true;
                txtPropertyValue.Font = new System.Drawing.Font(txtPropertyValue.Font.FontFamily, 14);

            }
            else
            {
                txtPropertyValue.Height = (int)sglPropertyValue_Height;
                txtPropertyValue.Width = (int)sglPropertyValue_Width;
                txtPropertyValue.Top = (int)sglPropertyValue_Top;
                txtPropertyValue.Left = (int)sglPropertyValue_Left;
                btnPropertyAdd.Top = (int)sglPropertyAdd_Top;
                btnPropertyUpdate.Top = (int)sglPropertyUpdate_Top;
                btnPropertyDelete.Top = (int)sglPropertyDelete_Top;
                btnDescribe.Top = (int)sglPropertyUpdate_Top;
                btnPropertyDelete.Visible = true;
                lblSpace.Visible = false;
                lblSpaceCount.Visible = false;
                txtPropertyValue.Font = new System.Drawing.Font(txtPropertyValue.Font.FontFamily, 8);
            }
        }

        private void DocPropertySelected()
        {
            const string MyProcName = "DocPropertySelected";
            int DocPropertyListIndex;
            int PropertyNameIndex;
            bool PropertyContainsCrLf;
            int PropertyValueIndex;
            int PropertyUOM_ListIndex;
            string UpdateMessage;
            int res = 0;
            DocPropertyListIndex = dgvDocProperty.CurrentCell.RowIndex;
            dgvDocProperty.Rows[DocPropertyListIndex].Selected = true;

            if (DocPropertyListIndex < 0)
                return;

            UpdateMessage = "";

            string selectedPropertyName = dgvDocProperty.Rows[DocPropertyListIndex].Cells[qdpColPropertyName].Value.ToString();
            /*string lastCharofProperty = selectedPropertyName.Last().ToString();
            if (int.TryParse(lastCharofProperty, out res))
            {
                selectedPropertyName = selectedPropertyName.Substring(0, selectedPropertyName.Length - 1);
            }
*/
            if (Property_Selected(pPropertyName: selectedPropertyName))
            {
            }
            else
            {
                DialogResult result = MessageBox.Show("ERROR: PropertyName " + PropertyName + " Not Found in Master-List" + "\r\n" + "do you want to REMOVE this Property from this DOCUMENT?", MyProcName, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    DgvRow_Delete(pdgv_Delete: dgvDocProperty, pblnPrompt: false);
                    btnFinish.Enabled = true;
                    return;//080125
                }
                else
                {

                }

                //return; //0801
            }

            ddlPropertyCategory.Enabled = false;
            ddlPropertyAll.Enabled = false;
            spnPropertyIndex.Enabled = false;

            if (spnPropertyIndex.Visible)
            {
                spnPropertyIndex.Minimum = Convert.ToInt32(dgvDocProperty.Rows[DocPropertyListIndex].Cells[qdpColPropertyIndex].Value);
                //spnPropertyIndex.Maximum = Convert.ToInt32(dgvDocProperty.Rows[DocPropertyListIndex].Cells[qdpColPropertyIndex].Value);//0801
                spnPropertyIndex.Text = dgvDocProperty.Rows[DocPropertyListIndex].Cells[qdpColPropertyIndex].Value.ToString();
                spnPropertyIndex.Value = Convert.ToInt32(spnPropertyIndex.Text);//0801
                UpdateMessage = UpdateMessage + "\r\n" + "Property Name and Index cannot be changed.  ";
            }
            else
            {
                UpdateMessage = UpdateMessage + "\r\n" + "Property Name cannot be changed; 'Index' is not applicable.  ";
            }
            UpdateMessage = UpdateMessage + "\r\n" + "Property '" + PropertyName + "' defined in MBR with this behavior:  ";
            UpdateMessage = UpdateMessage + "\r\n" + "Property Value is handled in Authoring as: " + PropertyValueAsName + ".  ";
            UpdateMessage = UpdateMessage + "\r\n" + "Property Value is " + PropertyUsageName + ".";

            PropertyValue = dgvDocProperty.Rows[DocPropertyListIndex].Cells[qdpColPropertyValue].Value.ToString();
            //PropertyIndex = dgvDocProperty.Rows[DocPropertyListIndex].Cells[qdpColPropertyIndex].Value.ToString();
            if (txtPropertyValue.Visible)
            {
                txtPropertyValue.Text = PropertyValue;
                if (txtPropertyValue.Enabled)
                {
                    //check
                    PropertyContainsCrLf = PropertyValue.IndexOf("\r\n") > 0;
                    txtPropertyValue_Large(PropertyContainsCrLf);
                    txtPropertyValue.Focus();
                }
                else
                {
                }
            }
            else
            {
                //ddlpropertyvalue is not populating in IntroPages
                PropertyValueIndex = ComboBox_Find(cboControl: ddlPropertyValue, intKeyColumn: qluColLookupText, strKeyValue: PropertyValue);
                if (PropertyValueIndex < 0)
                {
                    if (PropertyLookupType.Trim() == "NumericException")
                    {
                        ddlPropertyValue.Text = PropertyValue;
                        //ddlPropertyValue.SelectedValue = PropertyValue;
                    }
                    else if (PropertyValue == PropertyBracketed(Property_Instance_Unite(pPropertyName: PropertyName, pPropertyIndex: dgvDocProperty.Rows[DocPropertyListIndex].Cells[qdpColPropertyIndex].Value.ToString())))
                    {
                    }
                    else
                    {
                        UpdateMessage = UpdateMessage + "\r\n" + "WARNING: current PropertyValue '" + PropertyValue + "'" + "NOT FOUND in the Property Value-list '" + PropertyLookupType + "'.  ";
                    }
                }
                else
                {
                    ddlPropertyValue.SelectedIndex = PropertyValueIndex;
                }
                if (ddlPropertyValue.Enabled)
                {
                    ddlPropertyValue.Focus();
                }
            }

            if (dgvDocProperty.Rows[DocPropertyListIndex].Cells[qdpColPropertyUOM].Value != null)
            {
                PropertyUOM = dgvDocProperty.Rows[DocPropertyListIndex].Cells[qdpColPropertyUOM].Value.ToString();
            }
            else
            {
                PropertyUOM = "";
            }
            if (ddlPropertyUOM.Visible)
            {
                PropertyUOM_ListIndex = ComboBox_Locate(cboControl: ddlPropertyUOM, intKeyColumn: qluColLookupText, strKeyValue: PropertyUOM);
                if (PropertyUOM_ListIndex > -1)
                {
                    if (ddlPropertyUOM.Items.Count > 0) //171025 cbo is filling empty but selecting index 0 (mitigated).
                        ddlPropertyUOM.SelectedIndex = PropertyUOM_ListIndex;
                }
                else
                {
                    UpdateMessage = UpdateMessage + "\r\n" + "WARNING: DocPropertyUOM '" + PropertyUOM + "' not in Lookup List!  ";
                }
            }
            else
            {
            }

            btnPropertyAdd.Text = "Reset";

            if ((txtPropertyValue.Visible && txtPropertyValue.Enabled) || (ddlPropertyValue.Visible && ddlPropertyValue.Enabled))
            {
                UpdateMessage = UpdateMessage + "\r\n" + "Property Value can be changed.  ";
            }
            else
            {
                UpdateMessage = UpdateMessage + "\r\n" + "Property Value cannot be changed.  ";
            }
            if (ddlPropertyUOM.Visible)
            {
                if (ddlPropertyUOM.Enabled)
                {
                    UpdateMessage = UpdateMessage + "\r\n" + "Property Unit-of-Measure can be changed.  ";
                }
                else
                {
                    UpdateMessage = UpdateMessage + "\r\n" + "Property Unit-of-Measure cannot be changed.  ";
                }
            }
            else
            {
                UpdateMessage = UpdateMessage + "\r\n" + "Property Unit-of-Measure is not applicable.  ";
            }
            if ((txtPropertyValue.Visible && txtPropertyValue.Enabled) || (ddlPropertyValue.Visible && ddlPropertyValue.Enabled) || (ddlPropertyUOM.Visible && ddlPropertyUOM.Enabled))
            {
                btnPropertyUpdate.Enabled = true;
                UpdateMessage = UpdateMessage + "\r\n" + "<" + btnPropertyUpdate.Text + "> is allowed.  ";
            }
            else
            {
                btnPropertyUpdate.Enabled = false;
                UpdateMessage = UpdateMessage + "\r\n" + "<" + btnPropertyUpdate.Text + "> is not allowed, " + "because nothing can be changed.  ";
            }

            btnPropertyDelete.Enabled = true;
            lblUpdateMessage.Text = UpdateMessage;
            DocPropertyList_Visible(false);
        }

        private bool Property_Selected(string pPropertyName)
        {
            int PropertyNameIndex;
            bool Property_Selected = false;

            try
            {

                PropertyNameIndex = ComboBox_Find(cboControl: ddlPropertyAll, intKeyColumn: qpyColPropertyName, strKeyValue: pPropertyName);

                if (PropertyNameIndex < 0)
                {
                    if (ddlPropertyCategory.Text == "[Show All]") //100325
                    {
                    }
                    else
                    {
                        ddlPropertyCategory.Text = "[Show All]";
                        PropertyNameIndex = ComboBox_Find(cboControl: ddlPropertyAll, intKeyColumn: qpyColPropertyName, strKeyValue: pPropertyName);
                    }
                }

                if (PropertyNameIndex > -1)
                {
                    if (PropertyNameIndex <= ddlPropertyAll.Items.Count) //040225
                    {
                        ddlPropertyAll.SelectedIndex = PropertyNameIndex;
                        PropertyName = pPropertyName;
                        Property_Selected = true;
                    }

                }
                else
                {
                    ddlPropertyAll.SelectedIndex = -1;
                    PropertyName = pPropertyName;
                    Property_Selected = false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return Property_Selected;
        }

        //private void btnNext_Click(object sender, EventArgs e)
        //{

        //    string currentTabName = tabEditProperty.SelectedTab.Name;

        //    try
        //    {
        //        switch (currentTabName)
        //        {
        //            case "tabEdit":
        //                tabEdit.Visible = false;
        //                tabSource.Visible = true;
        //                tabUnused.Visible = false;
        //                tabMessages.Visible = false;
        //                tabEditProperty.SelectedTab = tabSource;
        //                break;
        //            case "tabSource":
        //                tabEdit.Visible = false;
        //                tabSource.Visible = false;
        //                tabUnused.Visible = true;
        //                tabMessages.Visible = false;
        //                tabEditProperty.SelectedTab = tabUnused;
        //                break;
        //            case "tabUnused":
        //                tabEdit.Visible = false;
        //                tabSource.Visible = false;
        //                tabUnused.Visible = false;
        //                tabMessages.Visible = true;
        //                tabEditProperty.SelectedTab = tabMessages;
        //                btnNext.Enabled = false;
        //                PropertyMessages_Load();
        //                break;
        //            //case "tabMessages":
        //            //    tabEdit.Visible = false;
        //            //    tabSource.Visible = false;
        //            //    tabUnused.Visible = false;
        //            //    tabMessages.Visible = true;
        //            //    break;
        //            default:

        //                break;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message.ToString());
        //    }
        //}

        private void Fill_WordDocument_Property(bool pblnShowAllProperty, DataGridView pdgvDocProperty)
        {

            string WORDPropertyName;
            string WORDPropertyValue;
            const string cstrDocPropertySource = "Document";
            int Index;
            dynamic dpcDocProperties;

            try
            {
                if (!TryResolveActiveDocumentForForm())
                {
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }

                dpcDocProperties = activeDocument.CustomDocumentProperties;
                foreach (dynamic dpyDocProperty in dpcDocProperties)
                {
                    WORDPropertyName = dpyDocProperty.Name;
                    //Index = DataGridView_Find(dgvControl: dgvDocProperty, intKeyColumn: qdpColPropertyIndex, strKeyValue: WORDPropertyName); //Comp old
                    Index = DataGridView_Find(dgvControl: pdgvDocProperty, intKeyColumn: qdpColPropertyInstance, strKeyValue: WORDPropertyName); // comp new
                    if (Index < 0)
                    {
                        WORDPropertyValue = dpyDocProperty.Value;
                        Index = DataGridView_Locate(dgvControl: pdgvDocProperty, intKeyColumn: qdpColPropertyName, strKeyValue: WORDPropertyName); //qdpcolpropertyName
                        MessageAdd("Add Document-Property '" + WORDPropertyName + "'");
                        DataRow dr = dtDocProperty.NewRow();
                        dtDocProperty.Rows.InsertAt(dr, Index);

                        Property_Instance_Split(pPropertyInstance: WORDPropertyName, pPropertyName: PropertyName, pPropertyIndex: PropertyIndex); //pPropertyName: PropertyName
                        pdgvDocProperty.Rows[Index].Cells[qdpColPropertyName].Value = PropertyName;//PropertyName
                        if (PropertyIndex == null)
                            PropertyIndex = Convert.ToString(0); //Added to resolve null exception
                        pdgvDocProperty.Rows[Index].Cells[qdpColPropertyIndex].Value = PropertyIndex;
                        pdgvDocProperty.Rows[Index].Cells[qdpColPropertyInstance].Value = WORDPropertyName;
                        Property_ValueDisplay_Split(pPropertyValueDisplay: WORDPropertyValue, pPropertyValue: PropertyValue, pPropertyUOM: PropertyUOM, pPropertyName: PropertyName); //pPropertyName: PropertyName
                        pdgvDocProperty.Rows[Index].Cells[qdpColPropertyValue].Value = PropertyValue;// 
                        pdgvDocProperty.Rows[Index].Cells[qdpColPropertyUOM].Value = PropertyUOM;
                        pdgvDocProperty.Rows[Index].Cells[qdpColPropertyValueDisplay].Value = WORDPropertyValue;
                        pdgvDocProperty.Rows[Index].Cells[qdpColMaterialID].Value = "-1";
                        pdgvDocProperty.Rows[Index].Cells[qdpColWorkCenterID].Value = "-1";

                    }
                    else
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                //this.Close();
                //CommonService.CleanUpBackgroundWordProcesses();
                CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "Fill_WordDocument_Property");
            }
            

        }

        /*private void Fill_WordDocument_Property(bool pblnShowAllProperty, DataGridView pdgvDocProperty) //020325
        {

            string WORDPropertyName;
            string WORDPropertyValue;
            const string cstrDocPropertySource = "Document";
            int Index;
            dynamic dpcDocProperties;

            try
            {
                dpcDocProperties = activeDocument.CustomDocumentProperties;
                foreach (dynamic dpyDocProperty in dpcDocProperties)
                {
                    WORDPropertyName = dpyDocProperty.Name;
                    Index = DataGridView_Find(dgvControl: pdgvDocProperty, intKeyColumn: qdpColPropertyInstance, strKeyValue: WORDPropertyName); // comp new
                    if (Index < 0)
                    {
                        WORDPropertyValue = dpyDocProperty.Value;
                        Index = DataGridView_Locate(dgvControl: pdgvDocProperty, intKeyColumn: qdpColPropertyName, strKeyValue: WORDPropertyName); //qdpcolpropertyName
                        MessageAdd("Add Document-Property '" + WORDPropertyName + "'");
                        DataRow dr = dtDocProperty.NewRow();
                        Property_Instance_Split(pPropertyInstance: WORDPropertyName, pPropertyName: PropertyName, pPropertyIndex: PropertyIndex); //pPropertyName: PropertyName
                        dr["PropertyName"] = PropertyName;
                        if (PropertyIndex == null)
                            PropertyIndex = Convert.ToString(0); //Added to resolve null exception
                        dr["PropertyIndex"] = PropertyIndex;
                        dr["PropertyInstance"] = WORDPropertyName;
                        Property_ValueDisplay_Split(pPropertyValueDisplay: WORDPropertyValue, pPropertyValue: PropertyValue, pPropertyUOM: PropertyUOM, pPropertyName: PropertyName); //pPropertyName: PropertyName
                        dr["PropertyValue"] = PropertyValue;
                        dr["PropertyUOM"] = PropertyUOM;
                        dr["PropertyValueDisplay"] = WORDPropertyValue;
                        dr["MaterialID"] = "-1";
                        dr["WorkCenterID"] = "-1";
                        dtDocProperty.Rows.InsertAt(dr, Index);
                        dtDocProperty.AcceptChanges();
                        pdgvDocProperty.Refresh();
                    }
                    else
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                this.Close();
            }


        }*/

        private void PropertySource_Load()
        {
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            System.Data.DataTable dtPropertySource = new System.Data.DataTable();


            dtPropertySource = objMBRAuthDAL.GetPropertySource(mDocID);
            if (dtPropertySource != null && dtPropertySource.Rows.Count > 0)
            {
                dtPropertySource.Columns.Remove("DocumentName");
                dgvPropertySource.DataSource = dtPropertySource;
                dgvPropertySource.AutoGenerateColumns = false;
                //dgvPropertySource.ColumnCount = 4;
                dgvPropertySource.Columns[0].Name = "PropertyInstance";
                dgvPropertySource.Columns[0].DataPropertyName = "PropertyInstance";
                dgvPropertySource.Columns[0].HeaderText = "Property Name";
                dgvPropertySource.Columns[1].Name = "PropertyValueDisplay";
                dgvPropertySource.Columns[1].DataPropertyName = "PropertyValueDisplay";
                dgvPropertySource.Columns[1].HeaderText = "Property Value";
                dgvPropertySource.Columns[2].Name = "DocumentNameDisplay";
                dgvPropertySource.Columns[2].DataPropertyName = "DocumentNameDisplay";
                dgvPropertySource.Columns[2].HeaderText = "Source Module";
                dgvPropertySource.Columns[3].Name = "DocLocationName";
                dgvPropertySource.Columns[3].DataPropertyName = "DocLocationName";
                dgvPropertySource.Columns[3].HeaderText = "Location";


            }
            else
            {
                MessageBox.Show("There are no properties available for this DocID");
            }


        }



        private void PropertyUnused_Load()
        {
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            System.Data.DataTable dtPropertyUnused = new System.Data.DataTable();

            dtPropertyUnused = objMBRAuthDAL.GetPropertyUnused(mDocID);
            dgvDocPropertyUnused.DataSource = null;
            if (dtPropertyUnused != null && dtPropertyUnused.Rows.Count > 0)
            {
                dgvDocPropertyUnused.DataSource = dtPropertyUnused;
                dgvDocPropertyUnused.AutoGenerateColumns = false;

                dgvDocPropertyUnused.Columns[0].Name = "PropertyIndex";
                dgvDocPropertyUnused.Columns[0].DataPropertyName = "PropertyIndex";
                dgvDocPropertyUnused.Columns[0].HeaderText = "";
                dgvDocPropertyUnused.Columns[1].Name = "PropertyName";
                dgvDocPropertyUnused.Columns[1].DataPropertyName = "PropertyName";
                dgvDocPropertyUnused.Columns[1].HeaderText = "Property Name";
                dgvDocPropertyUnused.Columns[2].Name = "PropertyValue";
                dgvDocPropertyUnused.Columns[2].DataPropertyName = "PropertyValue";
                dgvDocPropertyUnused.Columns[2].HeaderText = "Property Value";

                dgvDocPropertyUnused.Columns[0].Visible = false;

                dgvDocPropertyUnused.Columns[1].Width = 250;
                dgvDocPropertyUnused.Columns[2].Width = 250;


            }
            else
            {
                MessageBox.Show("There are no Unused properties available for this DocID");
            }


        }

        private void dgvPropertySource_DoubleClick(object sender, EventArgs e)
        {

            MessageBox.Show
            ("Details for selected Property in Source Document\n" +
             "Property Name:" + dgvPropertySource.SelectedRows[0].Cells["PropertyInstance"].Value + "\n" +
             "Property Value:" + dgvPropertySource.SelectedRows[0].Cells["PropertyValueDisplay"].Value + "\n" +
             "Property in Document:" + dgvPropertySource.SelectedRows[0].Cells["DocumentNameDisplay"].Value + "\n" +
             "Document Location:" + dgvPropertySource.SelectedRows[0].Cells["DocLocationName"].Value);

        }

        private void btnUnusedPropertyDelete_Click(object sender, EventArgs e)
        {
            int returnvalue;
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            Cursor.Current = Cursors.WaitCursor;
            isPropertyDeleted = true;
            foreach (DataGridViewRow row in dgvDocPropertyUnused.SelectedRows)
            {
                returnvalue = objMBRAuthDAL.DeleteSelectedProperty(mDocID, row.Cells["PropertyName"].Value.ToString(), Convert.ToInt32(row.Cells["PropertyIndex"].Value), row.Cells["PropertyValue"].Value.ToString());
            }
            Fill_Document_Property();
            Setup_Document_Property();
            PropertyUnused_Load();
            btnFinish.Enabled = true;
            Cursor.Current = Cursors.Default;
        }

        private void btnUnusedPropertyDeleteAll_Click(object sender, EventArgs e)
        {
            int returnvalue;
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            Cursor.Current = Cursors.WaitCursor;
            isPropertyDeleted = true;
            foreach (DataGridViewRow row in dgvDocPropertyUnused.Rows)
            {
                returnvalue = objMBRAuthDAL.DeleteSelectedProperty(mDocID, row.Cells["PropertyName"].Value.ToString(), Convert.ToInt32(row.Cells["PropertyIndex"].Value), row.Cells["PropertyValue"].Value.ToString());
            }
            Fill_Document_Property();
            Setup_Document_Property();
            PropertyUnused_Load();
            btnFinish.Enabled = true;
            Cursor.Current = Cursors.Default;
        }

        public int dgvDocProperty_AddItem(int DesiredListIndex = -812)
        {
            int intListIndex;
            int intColScan;
            if (DesiredListIndex == -812)
            {

                intListIndex = dgvDocProperty.Rows.Count - 1;
            }
            else
            {

                intListIndex = DesiredListIndex;
            }


            /*DataRow dr = dtDocModule.NewRow();
            dr["DocModuleID"] = -1;
            dr["ModuleSequence"] = (dtDocModule.Rows.Count) + 1;
            //dr["ModuleBaseName"] = dtModuleAll.Rows[ddlModuleAll.SelectedIndex]["DocumentBaseName"].ToString();
            dr["ModuleBaseName"] = dtModuleAll.Rows[ddlModuleAll.SelectedIndex - 1]["DocumentBaseName"].ToString(); //AutoSizeChanged due to mismatch b/while dt and ddl to fetch the correct index
            dr["ModuleUsage"] = strUsage;
            dr["ModuleRepetitions"] = Convert.ToInt32(txtRepetitions.Text);
            dr["Packet"] = strPacket;
            dr["ModuleTitle"] = txtModuleTitle.Text;
            dr["ModuleTableRows"] = TableRows;
            dr["DocumentTitle"] = dtModuleAll.Rows[ddlModuleAll.SelectedIndex]["DocumentTitle"].ToString();
            dtDocModule.Rows.Add(dr);
            dgvModuleSelected.DataSource = dtDocModule;*/

            dgvDocProperty.Rows.Add(); // to add new row.

            for (intColScan = 0; intColScan <= dgvDocProperty.ColumnCount - 1; intColScan++)
            {
                /*DataRow newBlankRow = new DataRow(); 
                dgvDocProperty.Rows[intListIndex].Add();*/

                dgvDocProperty.Rows[intListIndex].Cells[intColScan].Value = "";
            }
            return intListIndex;
        }

        private void PropertyReset()
        {

            dgvDocProperty.ClearSelection();

            ddlPropertyCategory.Enabled = true;

            if (ddlPropertyAll.Items.Count > 0) //170125
            {
                //ddlPropertyAll.SelectedIndex = 0;
                ddlPropertyAll_SelectedIndexChanged(ddlPropertyAll, EventArgs.Empty);
                ddlPropertyAll.Enabled = true;
                ddlPropertyAll.Focus();
            }


            /*if (ddlPropertyAll.SelectedIndex > 0)
            {
                ddlPropertyAll.SelectedIndex = 0;
                ddlPropertyAll.Enabled = true;
                ddlPropertyAll.Focus();
            }*/
            spnPropertyIndex.Enabled = true;
            btnPropertyAdd.Text = "Add";
            btnPropertyUpdate.Enabled = false;
            btnPropertyDelete.Enabled = false;
            DocPropertyList_Visible(true);
            txtPropertyValue_Large(false);
        }

        private void DocPropertyList_Visible(bool Visible)
        {
            lblDocProperty.Visible = Visible;
            dgvDocProperty.Visible = Visible;
            lblPromptDoubleClick.Visible = Visible;
            lblUpdateMessage.Visible = !Visible;
            lblUpdateMessage.Visible = Visible;
            lblReqCount.Visible = Visible;
        }
        private void MessageClear()
        {
            lstMessages.Items.Clear();
            MessageCount();
        }

        private void MessageAdd(string pstrMessage)
        {
            lstMessages.Items.Add(pstrMessage);
            MessageCount();
        }

        private void MessageCount()
        {
            lblMessages.Text = "Messages(" + lstMessages.Items.Count + ")";
        }


        private bool UserForm__MODE_Add()
        {
            bool UserForm__MODE_Add;
            UserForm__MODE_Add = (btnPropertyAdd.Text == "Add");
            return UserForm__MODE_Add;
        }

        private bool UserForm__MODE_Update() // not used on any menthod on old as well comparison
        {
            return !UserForm__MODE_Add();
        }

        public void txtPropertyValue__Forced()
        {
            string PropertyIndexSuffix;
            if (txtPropertyValue.Visible && !txtPropertyValue.Enabled)
            {
                if (spnPropertyIndex.Text == "0")
                {
                    PropertyIndexSuffix = "";
                }
                else
                {
                    PropertyIndexSuffix = spnPropertyIndex.Text;
                }
                txtPropertyValue.Text = "[" + PropertyName + PropertyIndexSuffix + "]";
            }
            else
            {
            }
        }

        private void spnPropertyIndex_ValueChanged(object sender, EventArgs e)
        {
            /*int pIndex = Convert.ToInt32(spnPropertyIndex.Text.Trim());
            pIndex = pIndex + 1;
            PropertyIndex = pIndex.ToString();
*/
            //PropertyIndex = spnPropertyIndex.Text;//070125
            spnPropertyIndex.Text = Convert.ToInt32(spnPropertyIndex.Value).ToString();
            PropertyIndex = spnPropertyIndex.Text;
            txtPropertyValue__Forced();
            ddlPropertyValue__UnassignedValue();

            if (PropertyDataType == PropertyDataType_Material) //200125
            {
                Fill_Document_Material(mDocID);
            }

        }



        //FormCommon Methods
        public void DataGridView_From_Array(DataGridView dgvTempGrid, string[,] pstrArray, DataGridView dgvTempGridAll = null, int pintKeyColumn = -1)
        {
            const string MyProcName = "DataGridView_From_Array";
            int intRow;
            int intRowAll;
            int intCol;
            int intColumnCount;
            int intRowCount;
            try
            {
                if (dgvTempGridAll == null & pintKeyColumn != -1)
                {
                    //Todo Information.Err.Raise(41611, MyCompName + "." + MyProcName, "Cannot specify KeyColumn=" + pintKeyColumn + " without 'All' ListBox");
                    var ex = new Exception();
                    ex.Data["errNum"] = "41611";
                    ex.Data["errDesc"] = "Cannot specify KeyColumn=" + pintKeyColumn + " without 'All' ListBox";
                    throw ex;
                    return;
                }

                if (!(dgvTempGridAll == null) & pintKeyColumn == -1)
                {
                    //Todo Information.Err.Raise(41612, MyCompName + "." + MyProcName, "Cannot specify ListBoxAll=" + plstListBoxAll.Name + " with KeyColumn=" + pintKeyColumn);
                    var ex = new Exception();
                    ex.Data["errNum"] = "41612";
                    ex.Data["errDesc"] = "Cannot specify ListBoxAll=" + dgvTempGrid.Name + " with KeyColumn=" + pintKeyColumn;
                    throw ex;
                    return;
                }

                dgvTempGrid.ClearSelection();

                //On Error GoTo Array_Is_Undefined


                intRowCount = pstrArray.GetUpperBound(1);
                intColumnCount = pstrArray.GetUpperBound(2) + 1;

                if (intColumnCount < 1)
                {
                    MessageBox.Show("Cannot process Array with Ubound(2)=" + intColumnCount, MyCompName + "." + MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (dgvTempGridAll == null)
                {
                    if (pintKeyColumn != -1)
                    {
                        MessageBox.Show("CANNOT SPECIFY KeyColumn=" + pintKeyColumn + " WITHOUT SPECIFYING 'GridAll'", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

                if (pintKeyColumn != -1)
                {
                    if (dgvTempGridAll == null)
                    {
                        MessageBox.Show("CANNOT SPECIFY GridAll=" + dgvTempGridAll.Name + " WITHOUT SPECIFYING 'KeyColumn'", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

                dgvTempGrid.ClearSelection();
                dgvTempGrid.ColumnCount = intColumnCount;
                for (intRow = 0; intRow <= intRowCount - 1; intRow++)
                {
                    dgvTempGrid.Rows.Add();
                    for (intCol = 0; intCol <= intColumnCount - 1; intCol++)
                    {
                        dgvTempGrid.Rows[intRow].Cells[intCol].Value = pstrArray[intRow + 1, intCol];
                    }
                }

                if (pintKeyColumn != -1)
                {
                    DataGridView_Exclude(dgvMaster: dgvTempGrid, dgvToExclude: dgvTempGridAll, pintKeyColumn: pintKeyColumn);
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                pstrArray = new string[0, 0];
                intRowCount = 0;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                }
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }
        public void DataGridView_To_Array(DataGridView dgvTempGrid, string[,] pstrArray)
        {
            const string MyProcName = "ListBox_To_Array";
            int intRow;
            int intCol;
            int intColumnCount;
            int intRowCount;
            intRowCount = dgvTempGrid.Rows.Count;
            intColumnCount = dgvTempGrid.Columns.Count;
            if (intColumnCount < 1)
            {
                MessageBox.Show("Cannot process Grid '" + dgvTempGrid.Name + "' with " + intColumnCount + " Columns", MyCompName + "." + MyProcName, MessageBoxButtons.OK);
                return;
            }

            pstrArray = new string[intRowCount + 1, intColumnCount - 1 + 1];
            for (intRow = 0; intRow <= intRowCount - 1; intRow++)
            {
                for (intCol = 0; intCol <= intColumnCount - 1; intCol++)
                {
                    pstrArray[intRow + 1, intCol] = dgvTempGrid.Rows[intRow].Cells[intCol].Value.ToString();
                }
            }
        }
        public void DataGridView_Exclude(DataGridView dgvMaster, DataGridView dgvToExclude, int pintKeyColumn)
        {
            int intRow;
            int intRowAll;
            for (intRow = 0; intRow <= dgvMaster.Rows.Count - 1; intRow++)
            {
                for (intRowAll = 0; intRowAll <= dgvToExclude.Rows.Count - 1; intRowAll++)
                {
                    if (dgvToExclude.Rows[intRowAll].Cells[pintKeyColumn].Value == dgvMaster.Rows[intRow].Cells[pintKeyColumn].Value)
                    {

                        dgvToExclude.Rows.RemoveAt(intRowAll);
                        intRowAll = dgvToExclude.Rows.Count;
                    }
                }
            }
        }
        public void DgvRow_Delete(System.Windows.Forms.DataGridView pdgv_Delete, bool pblnPrompt = true)
        {
            const string MyProcName = "ListBox_Delete";
            bool blnDelete;
            if (pdgv_Delete.CurrentCell.RowIndex == -1)
            {
                MessageBox.Show("An item must be selected in order to delete it.", MyProcName + " Error");
                return;
            }

            if (pblnPrompt)
            {
                DialogResult dialogRes = MessageBox.Show("Are you sure you want to delete this information?", MyProcName, MessageBoxButtons.YesNo);
                if (dialogRes == DialogResult.Yes)
                {
                    blnDelete = true;
                }
                else if (dialogRes == DialogResult.No)
                {
                    blnDelete = false;
                    return;
                }
                else
                {
                    blnDelete = true;
                }
                if (blnDelete)
                {
                    pdgv_Delete.Rows.RemoveAt(pdgv_Delete.CurrentCell.RowIndex);
                    dtDocProperty.AcceptChanges();
                    pdgv_Delete.DataSource = null;
                    pdgv_Delete.DataSource = dtDocProperty;
                    /*dgvDocProperty.Update();
                    dgvDocProperty.Refresh();*/
                }
            }
            else //080125
            {
                pdgv_Delete.Rows.RemoveAt(pdgv_Delete.CurrentCell.RowIndex);
                dtDocProperty.AcceptChanges();
                pdgv_Delete.DataSource = null;
                pdgv_Delete.DataSource = dtDocProperty;
                /*dgvDocProperty.Update();
                dgvDocProperty.Refresh();*/
            }
        }
        public void Fill_Lookup(string pstrFormName, string pstrLookupType, string ddlType)
        {
            const string MyProcName = "Fill_Lookup";

            string strLookupDefaultText = string.Empty;
            /*propertyValueLookupDataList = null; //  180125
            propertyValueLookupDataList = DatabaseHelper.LookupData_Select(lookupType: pstrLookupType, lookupDefault: strLookupDefaultText, orderBy: "1");*/
            //List<TblLookupData> lookupDataList = DatabaseHelper.LookupData_Select(lookupType: pstrLookupType, lookupDefault: strLookupDefaultText, orderBy: "1");

            if (propertyValueLookupDataList == null)
            {  // NewCode2
                var ex = new Exception();
                ex.Data["errNum"] = "0000";
                ex.Data["errDesc"] = "Issue on LookupData.";
                throw ex;
            }

            if (ddlType == "ddlPropertyValue")
            {
                propertyValueLookupDataList = null; //  180125
                propertyValueLookupDataList = DatabaseHelper.LookupData_Select(lookupType: pstrLookupType, lookupDefault: strLookupDefaultText, orderBy: "1");

                ddlPropertyValue.DataSource = propertyValueLookupDataList;
                ddlPropertyValue.DisplayMember = "LookupText";
                ddlPropertyValue.ValueMember = "LookupText";


            }
            if (ddlType == "ddlPropertyUOM")
            {
                propertyUOMLookupDataList = null; //  180125
                propertyUOMLookupDataList = DatabaseHelper.LookupData_Select(lookupType: pstrLookupType, lookupDefault: strLookupDefaultText, orderBy: "1");

                ddlPropertyUOM.DataSource = propertyUOMLookupDataList;  // 
                ddlPropertyUOM.DisplayMember = "LookupText";
                ddlPropertyUOM.ValueMember = "LookupText";

            }


        }

        public void Fill_Document_Material(int DocID)
        {
            const string MyProcName = "Fill_Document_Material";

            System.Data.DataTable dtDocMaterial = new System.Data.DataTable();
            int intDocMaterial;
            bool blnIncludeZeroAmount;
            int TrueDocID;
            if (DocID > 0)
            {
                TrueDocID = DocID;
                blnIncludeZeroAmount = true;
            }
            else
            {
                TrueDocID = DocID * -1;
                blnIncludeZeroAmount = true;
            }
            dtDocMaterial = objMBRAuthDAL.DocMaterial_Select(DocID: TrueDocID, IncludeZeroAmount: blnIncludeZeroAmount);

            dtDocMaterial = FillOtherMaterialDetails(dtDocMaterial, 0, "AddItem");
            dtDocMaterial = FillOtherMaterialDetails(dtDocMaterial, 1, cNotApplicable);


            ddlPropertyValue.DataSource = null;
            ddlPropertyValue.DataSource = dtDocMaterial;
            ddlPropertyValue.DisplayMember = "MaterialDisplay";
            ddlPropertyValue.ValueMember = "MaterialId";
            //ddlPropertyValue.SelectedIndex = 0;
        }

        public DataTable FillOtherMaterialDetails(DataTable dt, int insertPos, string strAdd)
        {
            if (strAdd == cNotApplicable)
            {

            }
            else
                strAdd = PropertyBracketed(Property_Instance_Unite(pPropertyName: PropertyName, pPropertyIndex: PropertyIndex));


            DataRow unAssignedItem = dt.NewRow();
            unAssignedItem[0] = strAdd;
            unAssignedItem[1] = -1;
            unAssignedItem[2] = strAdd;
            unAssignedItem[3] = "00000";
            unAssignedItem[4] = 0;
            unAssignedItem[5] = "'KG'";
            unAssignedItem[6] = -1; //180125
            unAssignedItem[7] = "Pkg Standard";
            unAssignedItem[8] = false;
            unAssignedItem[9] = "LABEL          ";
            unAssignedItem[10] = false;
            unAssignedItem[11] = false;
            unAssignedItem[12] = false;

            dt.Rows.InsertAt(unAssignedItem, insertPos);
            dt.AcceptChanges();

            return dt;
        }
        public int ComboBox_Find(System.Windows.Forms.ComboBox cboControl, int intKeyColumn, string strKeyValue, bool blnUCase = true)
        {
            string strKeyCompare;
            //int intScann;
            //GTIN Changes PANNEJ02 - soundhar - start
            //if (!mblnCalledWizard) //0801
            ////if (!mbWizardCopyCalled)
            //{
            //GTIN Changes PANNEJ02 - soundhar - end
            if (blnUCase)
                {
                    strKeyCompare = strKeyValue.ToUpper();
                }
                else
                {
                    strKeyCompare = strKeyValue;
                }
                intScann = 0;
                string strValue = "";
                for (intScann = cboControl.Items.Count - 1; intScann >= 0; intScann += -1)
                {
                    if (blnUCase)
                    {

                        //string strValue = cboControl.Items[intScan].ToString();
                        var strItem = cboControl.Items[intScann] as DataRowView;
                        if (strItem != null)
                        {
                            DataRow dataRow = strItem.Row;
                            var a = dataRow.ItemArray[intKeyColumn];
                            strValue = a.ToString();
                        }
                        else
                        {
                            var strItem2 = cboControl.Items[intScann];
                            if (strItem2 != null)
                            {
                                strValue = ((TblLookupData)cboControl.Items[intScann]).LookupText;
                            }
                        }


                        if (strValue.ToUpper() == strKeyCompare)
                        {
                            break;
                        }
                        else
                        {
                        }
                    }
                    else if (strValue == strKeyCompare)
                    {
                        break;
                    }
                    else
                    {
                    }
                }
            //GTIN Changes PANNEJ02 - soundhar - start
            // }
            //GTIN Changes PANNEJ02 - soundhar - end

            return intScann;
        }



        private int DataGridView_Find(System.Windows.Forms.DataGridView dgvControl, int intKeyColumn, string strKeyValue, bool blnUCase = true)
        {
            string strKeyCompare;
            int intScan;
            if (blnUCase)
            {
                strKeyCompare = strKeyValue.ToUpper();
            }
            else
            {
                strKeyCompare = strKeyValue;
            }
            intScan = 0;
            for (intScan = dgvControl.Rows.Count - 1; intScan >= 0; intScan--)
            {
                var row = dgvControl.Rows[intScan];
                string value = row.Cells[intKeyColumn].Value.ToString();
                if (blnUCase)
                {
                    if (value.ToUpper() == strKeyCompare)
                    {
                        break;
                    }
                }
                else
                {
                    if (value == strKeyCompare)
                    {
                        break;
                    }
                }
            }
            return intScan;
        }

        /*private int DataGridView_Find(System.Windows.Forms.DataGridView dgvControl, int intKeyColumn, string strKeyValue, bool blnUCase = true)
        {
            string strKeyCompare;
            int intScan;
            if (blnUCase)
            {
                strKeyCompare = strKeyValue.ToUpper();
            }
            else
            {
                strKeyCompare = strKeyValue;
            }
            intScan = 0;
            for (intScan = dgvControl.Items.Count - 1; intScan >= 0; intScan--)
            {
                var item = dgvControl.Items[intScan];
                string value = item.ToString();
                if (blnUCase)
                {
                    if (value.ToUpper() == strKeyCompare)
                    {
                        break;
                    }
                }
                else
                {
                    if (value == strKeyCompare)
                    {
                        break;
                    }
                }
            }
            return intScan;
        }*/

        /*public int DataGridView_Locate(System.Windows.Forms.DataGridView dgvControl, int intKeyColumn, string strKeyValue, bool blnLCase = true)
        {

            string strKeyCompare;
            int intScan = 0; // 
            if (blnLCase)
            {
                strKeyCompare = strKeyValue.ToLower();
            }
            else
            {
                strKeyCompare = strKeyValue;
            }

            //for (intScan = 0; intScan <= dgvControl.Rows.Count; intScan++)// 
            for (intScan = 0; intScan <= dgvControl.Rows.Count - 1; intScan++)
            {
                var row = dgvControl.Rows[intScan];
                string value = row.Cells[intKeyColumn].Value.ToString();
                if (blnLCase)
                {
                    if (value.ToLower() == strKeyCompare)
                    {
                        break;
                    }
                }
                else
                {
                    if (value == strKeyCompare)
                    {
                        break;
                    }
                }
            }

            return intScan;
        }*/

        public int DataGridView_Locate(System.Windows.Forms.DataGridView dgvControl, int intKeyColumn, string strKeyValue, bool blnLCase = true)
        {
            string strKeyCompare;
            int intScan = 0; // 
            if (blnLCase)
            {
                strKeyCompare = strKeyValue.ToLower();
            }
            else
            {
                strKeyCompare = strKeyValue;
            }

            for (intScan = 0; intScan <= dgvControl.Rows.Count; intScan++)// 
            {
                if (blnLCase)
                {
                    if (string.Compare(dgvControl.Rows[intScan].Cells[intKeyColumn].Value.ToString().ToLower(), strKeyCompare) < 0)
                    {
                    }
                    else
                    {
                        break;
                    }
                }
                else if (string.Compare(dgvControl.Rows[intScan].Cells[intKeyColumn].Value.ToString().ToLower(), strKeyCompare) < 0)
                {
                }

                else
                {
                    break;
                }
            }

            return intScan;
        }


        public int ComboBox_Locate(System.Windows.Forms.ComboBox cboControl, int intKeyColumn, string strKeyValue, bool blnLCase = true)
        {
            {
                string strKeyCompare;
                int intScan;
                if (blnLCase)
                {
                    strKeyCompare = strKeyValue.ToLower();
                }
                else
                {
                    strKeyCompare = strKeyValue;
                }
                string strValue = ""; //090125
                intScan = 0;
                for (intScan = 0; intScan <= cboControl.Items.Count - 1; intScan++)
                {

                    if (blnLCase)
                    {
                        TblLookupData res = (TblLookupData)cboControl.Items[intScan];

                        if (res != null)
                        {
                            strValue = res.LookupText.ToString();
                        }

                        if (string.Compare(strValue.ToLower(), strKeyCompare) < 0)
                        {
                        }
                        else
                        {
                            break;
                        }
                    }
                    else if (string.Compare(strValue.ToLower(), strKeyCompare) < 0)
                    {
                    }
                    else
                    {
                        break;
                    }
                }

                return intScan;
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {
            btnFinish_Click(sender, e);
        }

        private void shp1_Click(object sender, EventArgs e)
        {
            ShowTab(1);
        }

        private void shp2_Click(object sender, EventArgs e)
        {
            ShowTab(2);
        }

        private void shp3_Click(object sender, EventArgs e)
        {
            ShowTab(3);
        }

        private void label5_Click(object sender, EventArgs e)
        {
            btnFinish_Click(sender, e);
        }

        private void shp0_Click(object sender, EventArgs e)
        {
            ShowTab(0);
        }
        private void dgvDocProperty_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                DocPropertySelected();
                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "dgvDocProperty_CellDoubleClick");
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "dgvDocProperty_CellDoubleClick");
                }
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        public void GetActiveDocValues()
        {
            if (!TryResolveActiveDocumentForForm())
            {
                throw new InvalidOperationException("Unable to resolve the target document. los ethe form and try again.");
            }

            var result = _documentManager.Document_SelectAttributes(activeDocument.Name);
            mstrDocName = activeDocument.Name;
            mstrDocFullName = activeDocument.FullName;
            mDocID = result.DocumentID;
            mstrProtectionPassword = result.DocumentProtectionPassword;
            mintOpUnitID = result.DocumentOpUnitID;
            mintDocTypeID = result.DocumentDocTypeID; // 2301
            mstrOpUnitType = result.DocumentOpUnitType;
        }

        // Resolves the Word document this form should operate on.
        // Preference order:
        // 1) Reuse current activeDocument if it is still valid and matches.
        // 2) Find an open document by full name/name using Word windows/documents.
        // 3) Fall back to the currently active Word document.
        // Returns false when no suitable document can be resolved.
        private bool TryResolveActiveDocumentForForm()
        {
            try
            {
                // Fast-path: keep the existing reference when it is still usable.
                if (activeDocument != null)
                {
                    string currentDocumentName = activeDocument.Name;
                    if (string.IsNullOrWhiteSpace(mstrDocName) ||
                        string.Equals(currentDocumentName, mstrDocName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrWhiteSpace(mstrDocName))
                        {
                            mstrDocName = currentDocumentName;
                        }

                        return true;
                    }
                }
            }
            catch
            {
                // COM proxy may be stale; force a fresh resolve path.
                activeDocument = null;
            }

            // Retry a few times to handle transient Word COM timing issues.
            for (int retry = 0; retry < 3; retry++)
            {
                try
                {
                    Word.Document resolvedDocument = null;
                    Word.Application wordApp = GetWordApplicationForForm();

                    //// First try from the document that launched/hosts this add-in.
                    //if (Globals.ThisDocument?.ActiveDocument != null)
                    //{
                    //    resolvedDocument = FindOpenDocument(Globals.ThisDocument.ActiveDocument.Application, mstrDocName, mstrDocFullName);
                    //}

                    // Then try from the global Word application object.
                    if (resolvedDocument == null)
                    {
                        resolvedDocument = FindOpenDocument(wordApp, mstrDocName, mstrDocFullName);
                    }

                    // If a specific target name was provided but not found, retry briefly.
                    if (resolvedDocument == null && !string.IsNullOrWhiteSpace(mstrDocName))
                    {
                        if (retry < 2)
                        {
                            System.Threading.Thread.Sleep(100);
                            continue;
                        }

                        activeDocument = null;
                        return false;
                    }

                    // Last fallback: use Word's current active document.
                    if (resolvedDocument == null)
                    {
                        // Word.Application wordApp = Globals.ThisDocument?.Application;//commented
                        if (wordApp != null)
                        {
                            resolvedDocument = wordApp.ActiveDocument;
                        }
                    }

                    if (resolvedDocument == null)
                    {
                        activeDocument = null;
                        return false;
                    }

                    // Save the resolved target for downstream operations.
                    activeDocument = resolvedDocument;

                    if (string.IsNullOrWhiteSpace(mstrDocName))
                    {
                        mstrDocName = activeDocument.Name;
                    }

                    return true;
                }
                catch (COMException)
                {
                    // Transient COM failures can occur while Word is switching documents.
                    activeDocument = null;
                    if (retry < 2)
                    {
                        System.Threading.Thread.Sleep(100);
                        continue;
                    }

                    return false;
                }
                catch
                {
                    activeDocument = null;
                    return false;
                }
            }

            activeDocument = null;
            return false;
        }

        private Word.Application GetWordApplicationForForm()
        {
            try
            {
                if (activeDocument != null)
                {
                    return activeDocument.Application;
                }
            }
            catch
            {
            }

            try
            {
                if (Globals.ThisDocument?.Application != null)
                {
                    return Globals.ThisDocument.Application;
                }
            }
            catch
            {
            }

            try
            {
                return Marshal.GetActiveObject("Word.Application") as Word.Application;
            }
            catch
            {
                return null;
            }
        }

        // Attempts to locate an already-open document using both open windows and the document list.
        private static Word.Document FindOpenDocument(Word.Application wordApp, string documentName, string documentFullName)
        {
            if (wordApp == null)
            {
                return null;
            }

            Word.Document windowDocument = FindOpenDocumentByWindows(wordApp, documentName, documentFullName);
            if (windowDocument != null)
            {
                return windowDocument;
            }

            return FindOpenDocumentByDocuments(wordApp, documentName, documentFullName);
        }

        // Scans open Word windows first, which typically reflects visible user context.
        private static Word.Document FindOpenDocumentByWindows(Word.Application wordApp, string documentName, string documentFullName)
        {
            for (int windowIndex = 1; windowIndex <= wordApp.Windows.Count; windowIndex++)
            {
                Word.Window openWindow = wordApp.Windows[windowIndex];
                Word.Document openDocument = openWindow.Document;

                if (DocumentMatches(openDocument, documentName, documentFullName))
                {
                    return openDocument;
                }
            }

            return null;
        }

        // Scans Word's documents collection as a secondary lookup path.
        private static Word.Document FindOpenDocumentByDocuments(Word.Application wordApp, string documentName, string documentFullName)
        {
            for (int documentIndex = 1; documentIndex <= wordApp.Documents.Count; documentIndex++)
            {
                Word.Document openDocument = wordApp.Documents[documentIndex];
                if (DocumentMatches(openDocument, documentName, documentFullName))
                {
                    return openDocument;
                }
            }

            return null;
        }

        // Matches by full path when available; otherwise falls back to document name.
        private static bool DocumentMatches(Word.Document openDocument, string documentName, string documentFullName)
        {
            if (openDocument == null)
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(documentFullName))
            {
                return string.Equals(openDocument.FullName, documentFullName, StringComparison.OrdinalIgnoreCase);
            }

            if (!string.IsNullOrWhiteSpace(documentName))
            {
                return string.Equals(openDocument.Name, documentName, StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        /*private void ddlPropertyValue__UnassignedAddItem(int ValueColumn, int KeyColumn) //  working 181025
        {
            int ColumnIndex;
            {
                ddlPropertyValue.Visible = true;
                ddlPropertyValue.DataSource = null;
                ddlPropertyValue.Items.Insert(0, PropertyBracketed(Property_Instance_Unite(pPropertyName: PropertyName, pPropertyIndex: PropertyIndex)));                
            }
        }
        private void ddlPropertyValue__UnassignedValue() //working 170125
        {
            if (!ddlPropertyValue.Visible)
            {
                return;
            }
            if (lookupDataList.Count == 0)
                return;
            else
            {
                lookupDataList[0].LookupText = PropertyBracketed(Property_Instance_Unite(pPropertyName: PropertyName, pPropertyIndex: PropertyIndex));
                ddlPropertyValue_Load();
            }

        }
        private void ddlPropertyValue__NotApplicable_UnassignedValue()
        {
            if (!ddlPropertyValue.Visible)
            {
                return;
            }
            if (lookupDataList.Count == 0)
            {
                return;
            }
            else
            {
                ddlPropertyValue.Items.Add(cNotApplicable);
                lookupDataList[0].LookupAttributeValue = cNotApplicable;
                ddlPropertyValue_Load();
            }

        }
        */

        public class SectionLinkStatus
        {
            public bool PrimaryHeaderLink { get; set; }
            public bool FirstPageHeaderLink { get; set; }
            public bool EvenPagesHeaderLink { get; set; }
            public bool PrimaryFooterLink { get; set; }
            public bool FirstPageFooterLink { get; set; }
            public bool EvenPagesFooterLink { get; set; }
        }

        public Dictionary<int, SectionLinkStatus> SaveOriginalLinkStatus(Word.Application wordApp)
        {
            Word.Document doc = wordApp.ActiveDocument;
            var sectionStatuses = new Dictionary<int, SectionLinkStatus>();


            foreach (Word.Section section in doc.Sections)
            {
                var status = new SectionLinkStatus
                {
                    PrimaryHeaderLink = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious,
                    FirstPageHeaderLink = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious,
                    EvenPagesHeaderLink = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious,
                    PrimaryFooterLink = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious,
                    FirstPageFooterLink = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious,
                    EvenPagesFooterLink = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious
                };

                sectionStatuses.Add(section.Index, status);
            }

            return sectionStatuses;
        }

        public void LinkAllSections(Word.Application wordApp)
        {
            Word.Document doc = wordApp.ActiveDocument;


            foreach (Word.Section section in doc.Sections)
            {
                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = true;
                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;

                section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
                section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = true;
                section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;


                section.Range.Fields.Update();
            }
        }

        public void RestoreOriginalLinkStatus(Word.Application wordApp, Dictionary<int, SectionLinkStatus> originalStatuses)
        {
            Word.Document doc = wordApp.ActiveDocument;


            foreach (Word.Section section in doc.Sections)
            {
                if (originalStatuses.ContainsKey(section.Index))
                {
                    var status = originalStatuses[section.Index];

                    section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = status.PrimaryHeaderLink;
                    section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = status.FirstPageHeaderLink;
                    section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = status.EvenPagesHeaderLink;

                    section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = status.PrimaryFooterLink;
                    section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = status.FirstPageFooterLink;
                    section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = status.EvenPagesFooterLink;


                    section.Range.Fields.Update();
                }
            }

        }
        //dissolved this function in finish click event to perform db updates first before occurence of RPC error
        //public void PerformIntroPagesAction()
        //{
        //    mblnCalledWizard = true;
        //    Define_DocumentProperties();
        //    Setup_Database(true, dgvDocProperty);
        //    Setup_Document_Property();
        //    Setup_Document_Text(pblnReplaceTable: false);
        //    UpdateWordDocumentFields();
        //    UserForm__SHUTDOWN(pqrtQReturn: qReturn.qReturn_Finish);
        //}
        private void Define_DocumentProperties()
        {
            const string MyProcName = "Define_DocumentProperties";
            string ProductName;
            string ProductID;
            string RecordType;
            string ProductMfgProductCode;
            string sNDC_Code;
            string sProduct_Identification_Code;
            string sRCS_Sales_Code;
            string sGlobal_Trade_ID_Number;
            //added for GTIN - PANNEJ02
            string sVialLabelGlobalTradeIDNumber;
            string sDocTitle;
            //Dim arsDocPropertyRequired As ADODB.Recordset;
            const string DO_NOT_USE = "DO NOT USE";
            try
            {
                DocumentInfo documentInfo = new DocumentInfo();
                var docInfoResult = _documentManager.Document_SelectQPSObject(mstrDocName, documentInfo);

                if (DocIsModule == "Y")
                {
                    ProductName = PropertyBracketed(CommonService.cProductName);
                    ProductID = PropertyBracketed(CommonService.cProductID);
                }
                else
                {
                    ProductName = docInfoResult.GetClsDocument.mstrDocumentPath;
                    ProductID = docInfoResult.GetClsDocument.mstrMaterialEDPNUM;
                }
                DocumentType docType = _formManager.DocType_Select(mintDocTypeID);
                RecordType = docType.RecordType;

                //QAdd(pPropertyName: CommonService.cDocumentTitle, pPropertyValue: mstrDocName);//mstrDocumentTitle
                QAdd(pPropertyName: CommonService.cDocumentTitle, pPropertyValue: mstrDocumentTitle); //Get & set value from IntroPagesMBR to set DocumentTitle.
                QAdd(pPropertyName: CommonService.cProductName, pPropertyValue: ProductName);
                QAdd(pPropertyName: CommonService.cProductID, pPropertyValue: ProductID);
                QAdd(pPropertyName: CommonService.cRecordType, pPropertyValue: RecordType);

                /*RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cDocumentTitle, PropertyValue: txtDocumentTitle.Text);
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cProductName, PropertyValue: ProductName);
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cProductID, PropertyValue: ProductID);
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cRecordType, PropertyValue: RecordType);*/
                if (DocIsModule == "Y")
                {

                }
                else
                {
                    ProductMfgProductCode = docInfoResult.GetProductDetail.Manufacturing_Product_Code;
                    sNDC_Code = docInfoResult.GetProductDetail.NDC_Code;
                    sProduct_Identification_Code = docInfoResult.GetProductDetail.Product_Identification_Code;
                    sRCS_Sales_Code = docInfoResult.GetProductDetail.RCS_Sales_Code;
                    sGlobal_Trade_ID_Number = docInfoResult.GetProductDetail.Global_Trade_ID_Number;
                    //GTIN property -pannej02
                    sVialLabelGlobalTradeIDNumber = docInfoResult.GetProductDetail.VialLabelGlobalTradeIDNumber;

                    QAdd(pPropertyName: CommonService.cManufacturing_Product_Code, pPropertyValue: ProductMfgProductCode);
                    QAdd(pPropertyName: CommonService.cNDC_Code, pPropertyValue: sNDC_Code);
                    QAdd(pPropertyName: CommonService.cProduct_Identification_Code, pPropertyValue: sProduct_Identification_Code);
                    QAdd(pPropertyName: CommonService.cRCS_Sales_Code, pPropertyValue: sRCS_Sales_Code);
                    QAdd(pPropertyName: CommonService.cGlobal_Trade_ID_Number, pPropertyValue: sGlobal_Trade_ID_Number);
                    //added for GTIN property -pannej02
                    QAdd(pPropertyName: CommonService.cVialLabelGlobalTradeIDNumber, pPropertyValue: sVialLabelGlobalTradeIDNumber);

                    /*RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cManufacturing_Product_Code, PropertyValue: ProductMfgProductCode);
                    RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cNDC_Code, PropertyValue: sNDC_Code);
                    RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cProduct_Identification_Code, PropertyValue: sProduct_Identification_Code);
                    RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cRCS_Sales_Code, PropertyValue: sRCS_Sales_Code);
                    RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cGlobal_Trade_ID_Number, PropertyValue: sGlobal_Trade_ID_Number);*/
                }

                QAdd(pPropertyName: CommonService.cOrderNum, pPropertyValue: PropertyBracketed(CommonService.cOrderNum));
                QAdd(pPropertyName: CommonService.cLotNum, pPropertyValue: PropertyBracketed(CommonService.cLotNum));
                QAdd(pPropertyName: CommonService.cPrintDate, pPropertyValue: PropertyBracketed(CommonService.cPrintDate));
                QAdd(pPropertyName: CommonService.cPrintTime, pPropertyValue: PropertyBracketed(CommonService.cPrintTime));
                QAdd(pPropertyName: CommonService.cPrintPages, pPropertyValue: CommonService.cPrintPages_ValueDefault);
                QAdd(pPropertyName: CommonService.cChangeNum, pPropertyValue: CommonService.cChangeNum_ValueDefault);

                /*RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cOrderNum, PropertyValue: PropertyBracketed(CommonService.cOrderNum));
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cLotNum, PropertyValue: PropertyBracketed(CommonService.cLotNum));
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cPrintDate, PropertyValue: PropertyBracketed(CommonService.cPrintDate));
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cPrintTime, PropertyValue: PropertyBracketed(CommonService.cPrintTime));
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cPrintPages, PropertyValue: CommonService.cPrintPages_ValueDefault);
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cChangeNum, PropertyValue: CommonService.cChangeNum_ValueDefault);*/

            }
            catch (Exception ex)
            {
                if (ex.Data["errNum"] != null)  // NewCode1
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                }
            }
        }

        public void UpdateWordDocumentFields()
        {
            if (!TryResolveActiveDocumentForForm())
            {
                return;
            }

            foreach (Word.Section sec in activeDocument.Sections)
            {
                foreach (Word.Field aField in sec.Range.Fields)
                {
                    aField.Update();
                }
                foreach (Word.HeaderFooter header in sec.Headers)
                {
                    foreach (Word.Field aField in header.Range.Fields)
                    {
                        aField.Update();
                    }
                }
                foreach (Word.HeaderFooter footer in sec.Footers)
                {
                    foreach (Word.Field aField in footer.Range.Fields)
                    {
                        aField.Update();
                    }
                }
            }
        }

        public void Fill_Document_Property()
        {
            dtDocProperty = objMBRAuthDAL.GetDocumentProperty(mDocID);

            dtDocProperty.Columns.Remove("CountOfMasters");
            dtDocProperty.Columns.Remove("CountOfModules");
            //dgvDocProperty.DataSource = dtDocProperty;
            
            bindingSourceData.DataSource = dtDocProperty;//100925
            dgvDocProperty.DataSource = bindingSourceData;
            bindingSourceData.ResetBindings(false);
            dgvDocProperty.Refresh();

            dgvDocProperty.AutoGenerateColumns = false;

            dgvDocProperty.Columns[qdpColPropertyValueRequired].Name = "PropertyValueRequired";
            dgvDocProperty.Columns[qdpColPropertyValueRequired].DataPropertyName = "PropertyValueRequired";
            dgvDocProperty.Columns[qdpColPropertyValueRequired].HeaderText = "";
            dgvDocProperty.Columns[qdpColPropertyName].Name = "PropertyName";
            dgvDocProperty.Columns[qdpColPropertyName].DataPropertyName = "PropertyName";
            dgvDocProperty.Columns[qdpColPropertyName].HeaderText = "Property Name";
            dgvDocProperty.Columns[qdpColPropertyValueDisplay].Name = "PropertyValueDisplay";
            dgvDocProperty.Columns[qdpColPropertyValueDisplay].DataPropertyName = "PropertyValueDisplay";
            dgvDocProperty.Columns[qdpColPropertyValueDisplay].HeaderText = "Property Value";
            dgvDocProperty.Columns[qdpColPropertyInstance].Name = "PropertyInstance";
            dgvDocProperty.Columns[qdpColPropertyInstance].DataPropertyName = "PropertyInstance";
            dgvDocProperty.Columns[qdpColPropertyInstance].HeaderText = "PropertyInstance";
            dgvDocProperty.Columns[qdpColPropertyValue].Name = "PropertyValue";
            dgvDocProperty.Columns[qdpColPropertyValue].DataPropertyName = "PropertyValue";
            dgvDocProperty.Columns[qdpColPropertyValue].HeaderText = "PropertyValue";
            dgvDocProperty.Columns[qdpColPropertyUOM].Name = "PropertyUOM";
            dgvDocProperty.Columns[qdpColPropertyUOM].DataPropertyName = "PropertyUOM";
            dgvDocProperty.Columns[qdpColPropertyUOM].HeaderText = "PropertyUOM";
            dgvDocProperty.Columns[qdpColPropertyIndex].Name = "PropertyIndex";
            dgvDocProperty.Columns[qdpColPropertyIndex].DataPropertyName = "PropertyIndex";
            dgvDocProperty.Columns[qdpColPropertyIndex].HeaderText = "PropertyIndex";
            dgvDocProperty.Columns[qdpColWorkCenterID].Name = "WorkCenterID";
            dgvDocProperty.Columns[qdpColWorkCenterID].DataPropertyName = "WorkCenterID";
            dgvDocProperty.Columns[qdpColWorkCenterID].HeaderText = "WorkCenterID";
            dgvDocProperty.Columns[qdpColMaterialID].Name = "MaterialID";
            dgvDocProperty.Columns[qdpColMaterialID].DataPropertyName = "MaterialID";
            dgvDocProperty.Columns[qdpColMaterialID].HeaderText = "MaterialID";

            dgvDocProperty.Columns[qdpColPropertyValueRequired].Width = 75; //020325
            dgvDocProperty.Columns[qdpColPropertyInstance].Width = 240;
            dgvDocProperty.Columns[qdpColPropertyValueDisplay].Width = 240;

            dgvDocProperty.Columns[qdpColPropertyName].Visible = false;
            dgvDocProperty.Columns[qdpColPropertyValue].Visible = false;
            dgvDocProperty.Columns[qdpColPropertyUOM].Visible = false;
            dgvDocProperty.Columns[qdpColPropertyIndex].Visible = false;
            dgvDocProperty.Columns[qdpColWorkCenterID].Visible = false;
            dgvDocProperty.Columns[qdpColMaterialID].Visible = false;
        }

        //Performance
        private void WordKeepAliveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!TryResolveActiveDocumentForForm())
                {
                    return;
                }

                _ = activeDocument.Application.Version;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Keep-alive failed: " + ex.Message);
            }
        }

        private void StopWordKeepAliveTimer()
        {
            if (wordKeepAliveTimer != null)
            {
                wordKeepAliveTimer.Stop();
                wordKeepAliveTimer.Tick -= WordKeepAliveTimer_Tick;
                wordKeepAliveTimer.Dispose();
                wordKeepAliveTimer = null;
            }
        }

    }
}