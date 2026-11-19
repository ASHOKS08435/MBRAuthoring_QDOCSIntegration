using MBRAuthoring.Business;
using MBRAuthoring.DataAccess;
using MBRAuthoring.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace MBRAuthoring.Presentation
{
    public partial class FormIntroPagesMBR : FormBase
    {

        private Word.Document currentActiveDocument;

        public IDocumentManager _documentManager = null;// ServiceLocator.GetDocumentManager();

        const string MyCompName = "frmIntroPagesMBR";
        string mstrDocName;
        private string mstrDocType;
        private string mstrOpUnit;
        private string mstrDocIsModule;
        private bool mblnCancel;
        private bool mblnActivateRun;
        private int mDocID;
        bool mblnCleaningRecord;
        bool mblnMasterRecord;
        int mintOpUnitID;
        string mstrOpUnitType;
        string mstrDocTypeAbb;
        int mintDocTypeID;
        string mstrProtectionPassword;
        bool mblnDocTrackChgsReq;
        string mstrWorkCenter;
        string mstrDocHasMaterial;
        private string mstrDocHasModule;
        private string mstrDocHasProduct;
        private string mstrDocHasProperty;
        private int mintParmsValidated;
        string mstrDocHasWorkCenter;
        string mstrDocHasWorkCenterAlt;
        string mstrDocPermitsQuantitySufficient;
        const int cintParmsValidated = 21883;
        bool mblnCalledWizard;
        int mlngMyTop;
        Document mdocMBR;
        int mlngMyLeft;
        QWriteWhat qWriteFinal;
        qpsEditMaterialParms mqpsEditMaterialParms;
        qpsEditModuleParms mqpsEditModuleParms;
        qpsEditPropertyParms mqpsEditPropertyParms;
        private Word.Application wordApp;

        public FormIntroPagesMBR()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();
            PositionForm();
        }

        private void PositionForm()
        {
            if (mlngMyTop > 0 && mlngMyLeft > 0)
            {
                this.Top = mlngMyTop;
                this.Left = mlngMyLeft;
            }
        }

        private void frmIntroPagesMBR_Load(object sender, EventArgs e)
        {

            InitializeForm();
            InitializeTabs(mpgWizardPage, cmdBack, cmdNext, tabPage1, tabPage3); // 2301
            ShowTab(0);
        }

        /*private void frmIntroPagesMBR_Activate(object sender, EventArgs e)
        {
            try
            {
            if (mblnCancel)
            {
                this.Close();
            }
            if (mblnActivateRun)
                return;
            else
                mblnActivateRun = true;
            txtDocumentTitle.Text = _documentManager.Document_SelectQPSObject(mstrDocName, null, true).GetClsDocument.DocumentTitle; // Todo true?
            txtDocumentTitle.Enabled = false;
            if (!mblnCalledWizard)
            {
                if (this.Left >= mpgWizardPage.Left)
                {
                    this.Left = this.Left - mpgWizardPage.Left;
                }
                else
                {
                    this.Left = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex.Data["errNum"] != null)  // NewCode1
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "frmIntroPagesMBR_Activate");
                }
            }
        }*/
        private void frmIntroPagesMBR_Activate(object sender, EventArgs e)
        {
            try
            {
                if (mblnCancel)
                {
                    this.Close();
                    CommonService.CleanUpBackgroundWordProcesses();
                }
                if (mblnActivateRun)
                    return;
                else
                    mblnActivateRun = true;
                txtDocumentTitle.Text = _documentManager.Document_SelectQPSObject(mstrDocName, null, true).GetClsDocument.DocumentTitle; // Todo true?
                txtDocumentTitle.Enabled = false;
                if (!mblnCalledWizard)
                {
                    if (this.Left >= mpgWizardPage.Left)
                    {
                        this.Left = this.Left - mpgWizardPage.Left;
                    }
                    else
                    {
                        this.Left = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex.Data["errNum"] != null)  // NewCode1
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "frmIntroPagesMBR_Activate");
                }
            }
        }
        private void InitializeForm()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                currentActiveDocument = Globals.ThisDocument.ActiveDocument;//currentDocument;
                mdocMBR = Globals.ThisDocument.ActiveDocument;//currentDocument;
                mstrDocName = currentActiveDocument.Name;
                var selectAttribute = _documentManager.Document_SelectAttributes(mstrDocName);
                mstrProtectionPassword = selectAttribute.DocumentProtectionPassword;

                _documentManager.Document_Has(mstrDocName, ref mstrDocIsModule, ref mstrDocHasModule, ref mstrDocHasMaterial, ref mstrDocHasProduct, ref mstrDocHasProperty, ref mstrDocHasWorkCenter, ref mstrDocHasWorkCenterAlt, ref mstrDocPermitsQuantitySufficient);

                lblProduct.Visible = false;
                txtProduct.Visible = false;
                txtOpUnit.Enabled = false;
                txtDocType.Enabled = false;
                txtProduct.Enabled = false;
                txtOpUnit.Text = selectAttribute.DocumentOpUnit;//mstrOpUnit
                txtDocType.Text = selectAttribute.DocumentOpUnitType; //mstrDocType
                lblDocLocationComment.Visible = false;
                txtDocLocationComment.Visible = false;
                lblChangeNum.Visible = false;
                txtChangeNum.Visible = false;

                //Disable and removed the second tab page
                mpgWizardPage.TabPages.RemoveAt(1);

                cmdEditMaterial.Visible = (mstrDocHasMaterial != "N");
                cmdEditModule.Visible = (mstrDocHasModule != "N");
                cmdEditProduct.Visible = (mstrDocHasProduct != "N");
                if (mstrDocIsModule == "Y")
                {
                    cmdEditProperty.Visible = (mstrDocHasProperty != "N");
                }
                else
                {
                    cmdEditProperty.Visible = false;
                }
                cmdEditWorkCenter.Visible = false;
                cmdEditWorkCenterAlt.Visible = false;
                cboSupercedesDoc.Visible = false;
                lblSupercedesDoc.Visible = false;
                mblnActivateRun = false;
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "frmIntroPagesMBR_Load - InitializeForm");

            }
            finally // merge
            {
                Cursor = Cursors.Default;
            }
        }


        private void cmdBack_Click(object sender, EventArgs e)
        {
            ShowPrevTab();
        }

        private void cmdEditMaterial_Click(object sender, EventArgs e)
        {

            switch (RunEditMaterial(qRequest.qRequest_Display))
            {
                case (qReturn.qReturn_Back):
                    if (cmdBack.Enabled)
                        cmdBack_Click(sender, e);
                    break;
                case (qReturn.qReturn_Next):
                    if (cmdNext.Enabled)
                        cmdNext_Click(sender, e);
                    break;
            }
        }

        private void cmdEditWorkCenterAlt_Click(object sender, EventArgs e)
        {
            switch (RunEditMaterial(qRequest.qRequest_Display))
            {
                case (qReturn.qReturn_Back):
                    if (cmdBack.Enabled)
                        cmdBack_Click(sender, e);
                    break;
                case (qReturn.qReturn_Next):
                    if (cmdNext.Enabled)
                        cmdNext_Click(sender, e);
                    break;
            }


        }

        private void cmdEditModule_Click(object sender, EventArgs e)
        {
            switch (RunEditModule(qRequest.qRequest_Display))
            {
                case (qReturn.qReturn_Back):
                    if (cmdBack.Enabled)
                        cmdBack_Click(sender, e);
                    break;
                case (qReturn.qReturn_Next):
                    if (cmdNext.Enabled)
                        cmdNext_Click(sender, e);
                    break;
            }
        }

        private void cmdEditProperty_Click(object sender, EventArgs e)
        {
            switch (RunEditProperty(qRequest.qRequest_Display))
            {
                case (qReturn.qReturn_Back):
                    if (cmdBack.Enabled)
                        cmdBack_Click(sender, e);
                    break;
                case (qReturn.qReturn_Next):
                    if (cmdNext.Enabled)
                        cmdNext_Click(sender, e);
                    break;
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }

        private void cmdFinish_Click(object sender, EventArgs e)
        {
            try
            {

                if (!Valid_Entries())
                    return;

                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorWait;

                this.Left = currentActiveDocument.Application.Width - mpgWizardPage.Left;
                Define_Variables();
                qWriteFinal = QWriteWhat.qWriteMBRTable;
                Setup_Database();
                Setup_Document();
                if (!Globals.ThisDocument.docDetails.ContainsKey(mdocMBR.Name)) // 2301
                {
                    Globals.ThisDocument.docDetails.Add(mdocMBR.Name, mstrProtectionPassword);
                }
                SaveManager.Document_Save_FileAndDatabase(mdocMBR, docpasswordProtect: mstrProtectionPassword); // 2301
                //currentActiveDocument.Protect(WdProtectionType.wdAllowOnlyFormFields, false, mstrProtectionPassword, false); // 2301
                //currentActiveDocument.ActiveWindow.View.ShowRevisionsAndComments = blnRestoreRevView; // 2301
                CommonService.DocumentProtection_Off(currentActiveDocument, mstrProtectionPassword);

                currentActiveDocument.AcceptAllRevisions();
                if (mblnDocTrackChgsReq == true)
                {
                    if (currentActiveDocument.TrackRevisions == false)
                    {
                        currentActiveDocument.TrackRevisions = true;
                    }
                }
                CommonService.DocumentProtection_On(mdocMBR, mstrProtectionPassword);
                this.Close();
                currentActiveDocument.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
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

        private void Setup_Document()
        {
            const string MyProcName = "Setup_Document";
            int intSectionToProtect;
            int intBookmarksDeleted = 0;
            //List<bool> ablnSectionProtected = null;
            List<bool> ablnSectionProtected = new List<bool>(); //220125

            Table tbl;

            try
            {
                Selection selection = currentActiveDocument.Application.Selection;
                if (currentActiveDocument.ProtectionType != WdProtectionType.wdNoProtection)
                {
                    currentActiveDocument.Unprotect(mstrProtectionPassword);
                }
                currentActiveDocument.TrackRevisions = false;
                CommonService.PageSetup(selection.Range);
                Setup_DocumentProperties();

                CommonService.InsertHeaderMBR(1, "Lot Number: ", currentActiveDocument);
                CommonService.InsertFooterMBR(1, true, false, currentActiveDocument);
                CommonService.ViewPage_Main(currentActiveDocument.Application);

                if (mstrDocHasMaterial == "R")
                {
                    RunEditMaterial(qRequest.qRequest_Write, QWriteWhat.qWriteMBRTable);

                }
                if (mstrDocHasModule != "N")
                {
                    RunEditModule(qRequest.qRequest_Write, QWriteWhat.qWriteMBRTable);

                }
                if (mstrDocHasProperty != "N")
                {
                    //RunEditProperty(qRequest.qRequest_Write, QWriteWhat.qWriteMBRTable); //220125

                }

                // intSectionToProtect = selection.Sections[1].Index; // 2301
                selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                CommonService.InsertFooterMBR(2, false, true, currentActiveDocument);
                CommonService.ViewPage_Main(currentActiveDocument.Application);


                selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                CommonService.Bookmark_IntroEnd_Add(selection.Range);
                CommonService.SectionArray_Build(currentActiveDocument, ref ablnSectionProtected); //220125 // 2301
                ablnSectionProtected[0] = true; // 2301
                CommonService.DocumentProtection_Create(ablnSectionProtected, mstrProtectionPassword, intBookmarksDeleted, currentActiveDocument);
            }
            catch (Exception ex)
            {
                if (ex.Data["errNum"] != null) // NewCode2
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, CommonService.ErrorProcLatest(MyProcName));
                }
            }
        }

        private void Define_Variables()
        {
            CommonService.StepVarClear();
        }

        private void Setup_DocumentProperties()
        {
            //RunEditProperty(qRequest.qRequest_Write, QWriteWhat.qWriteMBRProperty); 220125
        }

        private void Setup_Database()
        {
            int SupercedesDoc_ListIndex;
            int SupercedesDocID;
            if (mstrDocHasMaterial != "N")
            {
                RunEditMaterial(qRequest.qRequest_Write, QWriteWhat.qWriteDBUpdate);
            }

            if (mstrDocHasModule != "N")
            {
                RunEditModule(qRequest.qRequest_Write, QWriteWhat.qWriteDBUpdate);
            }
            //Define_DocumentProperties(); //Need to check 080125
            if (mstrDocHasProperty != "N")
            {
                //RunEditProperty(qRequest.qRequest_Write, QWriteWhat.qWriteDBUpdate); //Handled on FormEditProperty 220125
            }

            if (DatabaseHelper.Transaction_Active() == mstrDocName)
            {
                DatabaseHelper.Transaction_Commit(mstrDocName);
            }
        }

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
            //added for GTIN PANNEJ02
            string sVialLabelGlobalTradeIDNumber;
            string sDocTitle;
            //Dim arsDocPropertyRequired As ADODB.Recordset;
            const string DO_NOT_USE = "DO NOT USE";
            try
            {
                DocumentInfo documentInfo = new DocumentInfo();
                var docInfoResult = _documentManager.Document_SelectQPSObject(mstrDocName, documentInfo);

                if (mstrDocIsModule == "Y")
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

                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cDocumentTitle, PropertyValue: txtDocumentTitle.Text);
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cProductName, PropertyValue: ProductName);
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cProductID, PropertyValue: ProductID);
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cRecordType, PropertyValue: RecordType);
                if (mstrDocIsModule == "Y")
                {

                }
                else
                {
                    ProductMfgProductCode = docInfoResult.GetProductDetail.Manufacturing_Product_Code;
                    sNDC_Code = docInfoResult.GetProductDetail.NDC_Code;
                    sProduct_Identification_Code = docInfoResult.GetProductDetail.Product_Identification_Code;
                    sRCS_Sales_Code = docInfoResult.GetProductDetail.RCS_Sales_Code;
                    sGlobal_Trade_ID_Number = docInfoResult.GetProductDetail.Global_Trade_ID_Number;
                    //added for GTIN PANNEJ02
                    sVialLabelGlobalTradeIDNumber = docInfoResult.GetProductDetail.VialLabelGlobalTradeIDNumber;
                    RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cManufacturing_Product_Code, PropertyValue: ProductMfgProductCode);
                    RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cNDC_Code, PropertyValue: sNDC_Code);
                    RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cProduct_Identification_Code, PropertyValue: sProduct_Identification_Code);
                    RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cRCS_Sales_Code, PropertyValue: sRCS_Sales_Code);
                    RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cGlobal_Trade_ID_Number, PropertyValue: sGlobal_Trade_ID_Number);
                    //added for GTIN - PANNEj02
                    RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cVialLabelGlobalTradeIDNumber, PropertyValue: sVialLabelGlobalTradeIDNumber);
                }


                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cOrderNum, PropertyValue: PropertyBracketed(CommonService.cOrderNum));
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cLotNum, PropertyValue: PropertyBracketed(CommonService.cLotNum));
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cPrintDate, PropertyValue: PropertyBracketed(CommonService.cPrintDate));
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cPrintTime, PropertyValue: PropertyBracketed(CommonService.cPrintTime));
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cPrintPages, PropertyValue: CommonService.cPrintPages_ValueDefault);
                RunEditProperty(qRequest_Validate: qRequest.qRequest_Add, PropertyName: CommonService.cChangeNum, PropertyValue: CommonService.cChangeNum_ValueDefault);

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

        private bool Valid_Entries()
        {
            mintParmsValidated = 0;
            if (mstrDocHasMaterial != "N")
            {
                if (RunEditMaterial(qRequest.qRequest_Validate) != qReturn.qReturn_Valid)
                {
                    //Page_JumpTo_SetFocus(Me, cmdEditMaterial)
                }
            }
            if (mstrDocHasModule != "N")
            {
                if (RunEditModule(qRequest.qRequest_Validate) != qReturn.qReturn_Valid)
                {

                }

            }
            if (mstrDocHasProperty != "N")
            {
                /*if (RunEditProperty(qRequest.qRequest_Validate) != qReturn.qReturn_Valid) //220125 commented bcoz this logic is done on editproperty itself
                {

                }*/
            }
            mintParmsValidated = cintParmsValidated;
            return true;
        }

        private qReturn RunEditMaterial(qRequest pqrqRequest, QWriteWhat pqwwWriteWhat = 0)
        {
            qReturn runMaterial = qReturn.qReturn_Cancel;
            FormEditMaterials editMaterial = new FormEditMaterials();

            StepVarManager.Form_Push(this);
            editMaterial.QLet(this.Top, this.Left + mpgWizardPage.Left, mqpsEditMaterialParms);
            switch (pqrqRequest)
            {
                case qRequest.qRequest_Display:
                    runMaterial = editMaterial.QDisplay(QDisplay_FirstLast.QDisplay_First);
                    break;
                case qRequest.qRequest_Validate:
                    runMaterial = editMaterial.QValidate();
                    break;
                case qRequest.qRequest_Write:
                    runMaterial = editMaterial.QWrite(pqwwWriteWhat);
                    break;
            }

            if (CommonService.QReturn_QGet(MyCompName, runMaterial))
            {
                editMaterial.QGet(mlngMyTop, mlngMyLeft, mqpsEditMaterialParms);
                mlngMyLeft = mlngMyLeft - mpgWizardPage.Left;
                PositionForm();
            }
            if (pqrqRequest == qRequest.qRequest_Write)
            {
                if (pqwwWriteWhat == qWriteFinal)
                {
                    editMaterial.QClose();
                }

            }
            else
            {
                editMaterial.QClose();
            }
            StepVarManager.Form_Pop(this);
            return runMaterial;
        }

        private qReturn RunEditModule(qRequest qRequest_Validate, QWriteWhat qWriteDBUpdate = 0)
        {
            qReturn runModule = qReturn.qReturn_Cancel;
            FormEditModule editModule = new FormEditModule();

            StepVarManager.Form_Push(this);
            editModule.QLet(this.Top, this.Left + mpgWizardPage.Left, mqpsEditModuleParms);

            switch (qRequest_Validate)
            {
                case qRequest.qRequest_Display:
                    runModule = editModule.QDisplay(QDisplay_FirstLast.QDisplay_First);
                    break;
                case qRequest.qRequest_Validate:
                    runModule = editModule.QValidate();
                    break;
                case qRequest.qRequest_Write:
                    runModule = editModule.QWrite(qWriteDBUpdate);
                    break;
            }

            if (CommonService.QReturn_QGet(MyCompName, runModule))
            {
                editModule.QGet(mlngMyTop, mlngMyLeft, mqpsEditModuleParms);
                mlngMyLeft = mlngMyLeft - mpgWizardPage.Left;
                PositionForm();

            }
            if (qRequest_Validate == qRequest.qRequest_Write)
            {
                if (qWriteDBUpdate == qWriteFinal)
                {
                    editModule.QClose();
                }
            }
            else
            {
                editModule.QClose();
            }

            StepVarManager.Form_Pop(this);
            return runModule;
        }


        private qReturn RunEditProperty(qRequest qRequest_Validate, QWriteWhat qWriteDBUpdate = 0, string PropertyName = "", string PropertyValue = "", int PropertyIndex = 0, string PropertyUOM = "")
        {
            qReturn runProperty = qReturn.qReturn_Cancel;
            FormEditProperty editProperty = new FormEditProperty();
            editProperty.SetTargetDocumentName(currentActiveDocument?.Name, currentActiveDocument?.FullName);
            editProperty.mstrDocumentTitle = txtDocumentTitle.Text; //270125 passed the value to set DocumentTitle

            StepVarManager.Form_Push(this);
            editProperty.QLet(this.Top, this.Left + mpgWizardPage.Left, mqpsEditPropertyParms);

            switch (qRequest_Validate)
            {
                case qRequest.qRequest_Display:
                    runProperty = editProperty.QDisplay(QDisplay_FirstLast.QDisplay_First);
                    break;
                case qRequest.qRequest_Validate:
                    runProperty = editProperty.QValidate();
                    break;
                case qRequest.qRequest_Write:
                    runProperty = editProperty.QWrite(qWriteDBUpdate);
                    break;
                case qRequest.qRequest_Add:
                    runProperty = editProperty.QAdd(PropertyName, PropertyValue, PropertyIndex, PropertyUOM);
                    break;
                case qRequest.qRequest_Delete:
                    runProperty = editProperty.QDelete(PropertyName);
                    break;

            }

            if (CommonService.QReturn_QGet(MyCompName, runProperty))
            {
                editProperty.QGet(mlngMyTop, mlngMyLeft, mqpsEditPropertyParms);
                mlngMyLeft = mlngMyLeft - mpgWizardPage.Left;
                PositionForm();

            }
            if (qRequest_Validate == qRequest.qRequest_Add)
            {
                if (qWriteDBUpdate == qWriteFinal)
                {
                    //editProperty.QClose(); //070125
                }
            }
            else
            {
                //editProperty.QClose(); //0701
            }

            StepVarManager.Form_Pop(this);
            return runProperty;
        }

        private void cmdNext_Click(object sender, EventArgs e)
        {
            ShowNextTab();
        }
    }

}


