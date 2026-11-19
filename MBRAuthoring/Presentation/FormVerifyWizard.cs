using MBRAuthoring.Business;
using System;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MBRAuthoring.Presentation
{
    public partial class FormVerifyWizard : FormBase
    {
        protected override Panel progressBarPanel => panel2;
        public IDocumentManager _documentManager = null;// ServiceLocator.GetDocumentManager();
                                                        // public ICommandBarManager _commandBarManager;
        private Word.Document currentActiveDocument;

        private Bitmap formImage;

        const string MyCompName = "FormVerifyWizard";
        string mstrDocName;
        int mDocID;
        bool mDocIsModule;
        string mDocPasswordRead;
        bool mblnCalledWizard;
        string mstrProtectionPassword;
        DocumentOptions mDocumentOptions;
        qReturn mqrtQReturn;

        private bool m_blnSysPropError;
        private const string usp_DocReqPropCheck = "usp_DocReqPropCheck";
        const string cDocVerifyAuth = "DocVerifyAuth";
        const string cDocVerifyLock = "DocVerifyLock";
        const string cDocVerifyStgd = "DocVerifyStgd";
        const string cDocVerifyAppr = "DocVerifyAppr";
        const string cDocVerifyBlck = "DocVerifyBlck";
        const string cDocVerifyArch = "DocVerifyArch";
        private const int MAX_PATH = 255;
        private bool ErrorHit;
        private const string aReturnMessage = "@ReturnMessage";
        private DataTable dt;

        //Todo
        // Private Declare PtrSafe Function apiGetTempDir Lib "kernel32" Ali "GetTempPathA" (ByVal nBufferLength  Long, ByVal lpBuffer String)  Long
        private struct DocumentOptions
        {
            public string OptionList;
            public bool AcceptAllChanges;
            public bool SetTrackChangesON;
            public bool SetTrackChangesOFF;
            public bool AssignWritePassword;
            public bool RemoveWritePassword;
            public bool VerifyNoUnacceptedChanges;
            public bool VerifyStatistics;
            public bool VerifySystemProperties;
            public bool VerifyProperties;
            public bool VerifyRestartNumbering;
            public bool VerifyDocumentView;
            public bool SetDocumentView;
            public bool VerifyInEDMS;
            public bool VerifyNotInEDMS;
            public bool VerifyRequiredProperty;
            public bool VerifyTrackChangesOff;
            public bool WordOpenIsRequired;
            public bool WordSaveIsRequired;

        }

        public FormVerifyWizard()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();
            dt = new DataTable();
            lstStatus.DataSource = dt;
            dt.Columns.Add("MasterDocID");
            dt.Columns.Add("MasterDocumentName");
            dt.Columns.Add("MasterLocationName");
            dt.Columns.Add("ModuleDocumentName");
            dt.Columns.Add("ModuleLocationName");
            dt.Columns.Add("PassFail");
            //lstStatus.RowTemplate.Height = 15;
            lstStatus.AutoGenerateColumns = false;
            lstStatus.Columns["MasterDocID"].Visible = false;
            lstStatus.Columns[1].Width = 225;
            lstStatus.Columns[2].Width = 45;
            lstStatus.Columns[3].Width = 125;
            lstStatus.Columns[4].Width = 47;
            lstStatus.Columns[5].Width = 90;
        }

        private void FormVerifyWizard_Load(object sender, EventArgs e)
        {

            Cursor = Cursors.WaitCursor;
            currentActiveDocument = Globals.ThisDocument.ActiveDocument;
            string strDocComparePath;
            strDocComparePath = @"C:WINDOWS\Pfizer\DocCompare*";

            SetToolTip(lable0, "To skip to this step, click here.");  // NewCode
            SetToolTip(shp0, "To skip to this step, click here.");
            SetToolTip(lable1, "To skip to this step, click here.");
            SetToolTip(label3, "To skip to this step, click here.");
            SetToolTip(cmdPrint, "Print this Wizard Form");  // NewCode

            //To Do
            //Call Form_TabIndexing(Me)

            mstrDocName = currentActiveDocument.Name;
            var result = _documentManager.Document_SelectAttributes(mstrDocName);
            mDocID = result.DocumentID;
            mstrProtectionPassword = result.DocumentProtectionPassword;

            mblnCalledWizard = false;
            lblTitle.Text = "Document Verify";
            lblThisDocVerify.Text = lblThisDocVerify.Text + "  '" + mstrDocName + "'  as of " + DateTime.Now.ToString("dddd, dd-MMM-yyyy");
            SetToolTip(lblThisDocVerify, "Related Documents");  // NewCode

            //To do
            //Call Page_JumpTo(Me, 0)
            cmdFinish.Enabled = false;
            DocumentAuthVerify(mDocID, mDocIsModule, mstrDocName, _formManager.GetDocumentVerifyPath(), mDocPasswordRead, mstrProtectionPassword, mstrProtectionPassword);

            if (File.Exists(strDocComparePath))
            {
                cmdRunDocCompare.Visible = true;
                cmdRunDocCompare.Enabled = true;
                cmdRunDocCompare.Text = "Run DocCompare on '" + mstrDocName + "'";

            }

            else
            {

                cmdRunDocCompare.Visible = true;
                cmdRunDocCompare.Enabled = false;
                cmdRunDocCompare.Text = "DocCompare NOT installed on this computer";
                SetToolTip(cmdRunDocCompare, "To Install run 'DMS Add New Programs'");
            }


            chkMoveToLockedState.Text = "Move  '" + mstrDocName + "'  to a LOCKED State";
            Cursor = Cursors.Default;
            currentActiveDocument.Application.StatusBar = "";
        }

        //private void FormVerifyWizard_Activated(object sender, EventArgs e)
        //{
        //    // DocumentAuthVerify(mDocID, mDocIsModule, mstrDocName, _formManager.GetDocumentVerifyPath(), mDocPasswordRead, mstrProtectionPassword, mstrProtectionPassword);

        //}
        public void SetPathTrailingSlash(string PathName) //this is not in use
        {
            if (PathName.Substring(PathName.Length - 1) == "\\")
            {
                return;
            }

            PathName = PathName + "\\";
        }

        public void fReturnTempDir()  //this is not is use
        {
            string strTempDir;
            long lngx;
            strTempDir = new string('\0', MAX_PATH);
            //To DO
            //            lngx = apiGetTempDir(MAX_PATH, strTempDir);
            //If lngx<> 0 Then
            //fReturnTempDir = Left$(strTempDir, lngx)
            //Else
            //fReturnTempDir = ""

        }

        public bool DocumentAuthVerify(int DocID, bool DocIsModule, string DocumentName, string DocumentPath, string DocPasswordRead, string DocPasswordWrite, string DocPasswordProtect)
        {
            bool VerificationResult = false;
            bool VerifySuccess;
            bool ExecAllowed = false;
            string PasswordWrite_Ignore = "";

            mDocumentOptions.OptionList = _formManager.AppStatus_Select(cDocVerifyAuth);

            ExecAllowed = DocumentOptionsSetting(ref mDocumentOptions, mDocumentOptions.OptionList, ExecAllowed, PasswordWrite_Ignore, false);
            if (ExecAllowed)
            {
                VerifySuccess = FileProcessing(DocID, DocumentPath, DocumentName, DocPasswordRead, DocPasswordWrite, DocPasswordWrite, DocPasswordProtect, DocIsModule, mDocumentOptions, ref VerificationResult);

                if (!VerificationResult)
                {
                    VerificationResult = true;
                    ShowStatus("MS/Word processing is not required.");
                }
            }
            else
            {
                VerifySuccess = false;
            }

            ReportComplete("Verification for Authoring Document " + DocumentName, VerifySuccess);
            var docResult = _formManager.Fill_Document_Validation(DocID);
            SetToolTip(lstStatus, "lstStatus: DocID = " + DocID);  // NewCode
            if (docResult.Count == 0)
            {
                //lstStatus.Rows.Clear();
                DataRow dataRow = dt.NewRow();

                dataRow["MasterDocumentName"] = "No Related Documents Exist";
                dt.Rows.Add(dataRow);
            }
            else
            {
                lstStatus.DataSource = docResult;
            }

            SetToolTip(lblRelatedDocs, "Related Documents");  // NewCode

            if (lstStatus.Rows.Count < 2)
            {
                if (lstStatus.Rows[0].Cells["MasterDocumentName"].Value.ToString() == "No Related Documents Exist")
                {
                    lblRelatedDocs.Text = "Related Documents Count:  0";

                }
                else
                {
                    lblRelatedDocs.Text = "Related Documents Count:  " + lstStatus.Rows.Count;
                }

            }
            else
            {
                lblRelatedDocs.Text = "Related Documents Count:  " + lstStatus.Rows.Count;

            }
            if (VerifySuccess)//Merge 080125
            {
                txtStatus.BackColor = Color.FromArgb(0xC0, 0xFF, 0xC0); //green
            }
            else
            {
                txtStatus.BackColor = Color.FromArgb(0xFF, 0xAF, 0xA3); //dark pink
            }

            return VerifySuccess;
        }


        private bool DocumentOptionsSetting(ref DocumentOptions DocOptions, string OptionList, bool ExecAllowed, string passwordWriteAction, bool DocumentSaveAllowed)
        {
            const string MyProcName = "DocumentOptionsSetting";
            string[] OptionCodes;
            int OptionCount;
            int OptionScan;
            OptionList = OptionList.Trim();
            DocOptions.OptionList = OptionList;
            if (string.IsNullOrEmpty(OptionList))
            {
                ExecAllowed = false;
                passwordWriteAction = "=";
                return false;
            }

            else
            {
                ExecAllowed = true;
            }
            OptionCodes = OptionList.Split(',');
            OptionCount = OptionCodes.GetUpperBound(0); // Will return index of the last element.
            for (OptionScan = 0; OptionScan <= OptionCount; OptionScan++)
            {
                switch (OptionCodes[OptionScan].Trim().ToUpper())
                {
                    case "AC":
                        DocOptions.AcceptAllChanges = true;
                        DocOptions.WordOpenIsRequired = true;
                        DocOptions.WordSaveIsRequired = true;
                        break;
                    case "SV":
                        DocOptions.SetDocumentView = true;
                        DocOptions.WordOpenIsRequired = true;
                        DocOptions.WordSaveIsRequired = true;
                        break;
                    case "TN":
                        DocOptions.SetTrackChangesOFF = true;
                        DocOptions.WordOpenIsRequired = true;
                        DocOptions.WordSaveIsRequired = true;
                        break;
                    case "TY":
                        DocOptions.SetTrackChangesON = true;
                        DocOptions.WordOpenIsRequired = true;
                        DocOptions.WordSaveIsRequired = true;
                        break;
                    case "VM":
                        DocOptions.VerifyInEDMS = true;
                        break;
                    case "VN":
                        DocOptions.VerifyNotInEDMS = true;
                        break;
                    case "VP":
                        DocOptions.VerifyProperties = true;
                        DocOptions.WordOpenIsRequired = true;
                        break;
                    case "VQ":
                        DocOptions.VerifyRequiredProperty = true;
                        break;
                    case "VR":
                        DocOptions.VerifyRestartNumbering = true;
                        DocOptions.WordOpenIsRequired = true;
                        break;
                    case "VS":
                        DocOptions.VerifyStatistics = true;
                        DocOptions.WordOpenIsRequired = true;
                        break;
                    case "VI": // "VI" category is not present on DB (table: tblAppStatus).
                        DocOptions.VerifySystemProperties = true;
                        break;
                    case "VT":
                        DocOptions.VerifyTrackChangesOff = true;
                        DocOptions.WordOpenIsRequired = true;
                        break;
                    case "VU":
                        DocOptions.VerifyNoUnacceptedChanges = true;
                        DocOptions.WordOpenIsRequired = true;
                        break;
                    case "VV":
                        DocOptions.VerifyDocumentView = true;
                        DocOptions.WordOpenIsRequired = true;
                        break;
                    case "WA":
                        DocOptions.AssignWritePassword = true;
                        DocOptions.WordOpenIsRequired = true;
                        DocOptions.WordSaveIsRequired = true;
                        break;
                    case "WR":
                        DocOptions.RemoveWritePassword = true;
                        DocOptions.WordOpenIsRequired = true;
                        DocOptions.WordSaveIsRequired = true;
                        break;
                    default:
                        var ex = new Exception();
                        ex.Data["errNum"] = "9023";
                        ex.Data["errDesc"] = "ERROR: OptionCode '" + OptionCodes[OptionScan] + "' not recognized";
                        throw ex;
                        break;
                }

            }

            if (DocOptions.AssignWritePassword)
            {
                passwordWriteAction = "A";
            }
            else if (DocOptions.RemoveWritePassword)
            {
                passwordWriteAction = "R";

            }
            else
            {
                passwordWriteAction = "=";
            }

            if (DocOptions.WordSaveIsRequired && !DocumentSaveAllowed)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "9024";
                ex.Data["errDesc"] = "ERROR: OptionCodes '" + OptionList + "' imply Document Update, which is not allowed in this context";
                throw ex;
            }
            return ExecAllowed;
        }

        private bool FileProcessing(int DocID, string DocumentPath, string DocumentName, string DocPasswordRead, string DocPasswordWriteCurrent, string DocPasswordWriteNew, string DocPasswordProtect, bool DocIsModule, DocumentOptions DocOptions, ref bool VerificationResult)
        {
            bool fileprocessing;
            string DocumentFileName;
            bool DocumentFound;
            bool ErrorDetected;
            string strMessage = "";
            bool blnSysPropError;
            ErrorDetected = false;
            fileprocessing = false;
            DocumentFileName = "";
            DocumentFound = false;
            //string gstrAppVersionNumber = "1.6.0"; //To do , Hard coded

            ShowStatus("MBR_Authoring Version=" + ModChangeHistory.gstrAppVersionNumber);

            ShowStatus("OptionList=" + DocOptions.OptionList);
            if (DocOptions.VerifyRequiredProperty)
            {
                strMessage = VerifyReqProperty(DocID);
                if (string.IsNullOrEmpty(strMessage))
                {
                    ShowStatus("All required properties are present");
                }
                else
                {
                    ShowStatus(strMessage);
                    ErrorDetected = true;
                }

            }

            if (DocOptions.VerifySystemProperties)
            {
                strMessage = VerifySystemProperties(DocID); //This method is not in use
                if (string.IsNullOrEmpty(strMessage))
                {
                    ShowStatus("All system properties are in sync");

                }
                else
                {
                    ShowStatus(strMessage);
                    m_blnSysPropError = true;
                }

            }

            fileprocessing = true;
            VerificationResult = true;
            ShowStatus("MS/Word processing is not required.");
            if (ErrorDetected)
            {
                fileprocessing = false;

            }

            return fileprocessing;
        }

        public void ReportComplete(string strEvent, bool Success)
        {
            try
            {
                if (Success)
                {
                    txtStatus.ReadOnly = false;
                    txtStatus.Text = txtStatus.Text + DateTime.Now.ToString("HH:mm:ss") + ": **** Successful " + strEvent + " ****";

                }
                else
                {
                    txtStatus.Text = txtStatus.Text + DateTime.Now.ToString("HH:mm:ss") + ": !!!! FAILED " + strEvent + " !!!!";

                }
                cmdFinish.Enabled = true;
            }
            catch (Exception e)
            {
                CommonService.Error_Handler(e.HResult.ToString(), e.Message, MyCompName, "ReportComplete");
            }
        }

        public void ShowStatus(string strEvent)
        {
            const string MyProcName = "ShowStatus";
            try
            {
                if (cmdFinish.Enabled)
                {
                    CommonService.Error_Handler("0", "Called After 'ReportComplete'", MyCompName, MyProcName);
                    ErrorHit = true;

                }
                txtStatus.Text = txtStatus.Text + DateTime.Now.ToString("HH:mm:ss tt") + ": " + strEvent + "\r\n";
                txtStatus.Multiline = true;
                if (!this.Visible)
                {
                    this.Show();
                }
            }
            catch (Exception e)
            {
                CommonService.Error_Handler(e.HResult.ToString(), e.Message, MyCompName, MyProcName);

            }
        }

        public string VerifyReqProperty(int p_DocID)
        {
            const string MyProcName = "VerifyReqProperty";
            string resultVerifyReqProperty = "";

            try
            {
                resultVerifyReqProperty = _formManager.VerifyProperty(p_DocID);

            }
            catch (Exception ex)//(ApplicationException ex) //Need to check
            {
                MessageBox.Show("Error (" + ex.HResult.ToString() + "): " + ex.Message, "VerifyReqProperty", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            return resultVerifyReqProperty;
        }

        public string VerifySystemProperties(int p_DocID)
        {
            string result = "";
            // [usp_DocSysPropCheck] is using to get message . But this SP is not present on DB

            const string MyProcName = "VerifySystemProperties";
            //            On Error GoTo VerifySystemPropertiesError
            //VerifySystemProperties = ""
            //Dim objCmd As New ADODB.Command
            //With objCmd
            //.ActiveConnection = GetConnection
            //.CommandText = "usp_DocSysPropCheck"
            //.CommandType = adCmdStoredProc
            //.Parameters.Refresh
            //.Parameters(aDocID).Value = p_DocID
            //.Execute
            //End With
            //If objCmd.Parameters(aReturnMessage).Value <> "" Then
            //VerifySystemProperties = "ERROR in System Property-Verification: " & objCmd.Parameters(aReturnMessage).Value
            //End If
            //Exit Function
            //VerifySystemPropertiesError:
            //            MsgBox "ERROR (" & Err.Number & "): " & Err.Description, vbCritical, MyProcName
            //            VerifySystemProperties = " "
            //End Function
            return result;
        }

        public void ModuleValidate(int p_DocID)
        {
            //method is not calling anywhere. Need to check .
            const string MyProcName = "ModuleValidate";
            try
            {

            }
            catch (Exception ex)
            {

            }
            //Public Function ModuleValidate(p_DocID As Long) As String
            //Const MyProcName As String = "ModuleValidate"
            //On Error GoTo ModuleValidateError
            //ModuleValidate = ""
            //Dim objCmd As New ADODB.Command
            //With objCmd
            //.ActiveConnection = GetConnection
            //.CommandText = "usp_ModuleValidate"
            //.CommandType = adCmdStoredProc
            //.Parameters.Refresh
            //.Parameters(aDocID).Value = p_DocID
            //.Execute
            //End With
            //If objCmd.Parameters(aReturnMessage).Value <> "" Then
            //ModuleValidate = "ERROR in System Property-Verification: " & objCmd.Parameters(aReturnMessage).Value
            //End If
            //Exit Function
            //ModuleValidateError:
            //            MsgBox "ERROR (" & Err.Number & "): " & Err.Description, vbCritical, MyProcName
            //            ModuleValidate = " "
            //End Function

        }

        private void cmdPrint_Click(object sender, EventArgs e)
        {
            cmdRunDocCompare.Visible = false;
            cmdPrint.Visible = false;
            cmdCancel.Visible = false;
            cmdBack.Visible = false;
            cmdNext.Visible = false;
            cmdFinish.Visible = false;
            chkMoveToLockedState.Visible = false;
            lstStatus.Height = 180;

            //Capture the form as Image
            formImage = new Bitmap(this.Width, this.Height);
            this.DrawToBitmap(formImage, new Rectangle(0, 0, this.Width, this.Height));

            PrintDocument printDocument = new PrintDocument();
            printDocument.PrintPage += PrintDocument_PrintPage;

            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = printDocument;

            printDialog.Document = printDocument;
            printDocument.Print();

            cmdRunDocCompare.Visible = true;
            cmdPrint.Visible = true;
            cmdCancel.Visible = true;
            cmdBack.Visible = true;
            cmdNext.Visible = true;
            cmdFinish.Visible = true;
            chkMoveToLockedState.Visible = true;
            lstStatus.Height = 148;
        }


        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            FormBorderStyle originalBorderStyle = this.FormBorderStyle;
            this.FormBorderStyle = FormBorderStyle.None;

            Application.DoEvents();
            //Capture the form as Image
            formImage = new Bitmap(this.Width, this.Height);
            this.DrawToBitmap(formImage, new Rectangle(0, 0, formImage.Width, formImage.Height));
            e.Graphics.DrawImage(formImage, 0, 0);

            this.FormBorderStyle = originalBorderStyle;
            Application.DoEvents();

        }

        private void cmdRunDocCompare_Click(object sender, EventArgs e) //Not aware of the functionality, Even user also never used this .
        {
                double dblDocCompare;
                string strPath_DocCompareExe;
                string strDocPath;
                string strCurrDir;
                //To do
                //strCurrDir = currentActiveDocument.AttachedTemplate.Path;
                // SetCurDir(strCurrDir);
                strDocPath = _formManager.GetDocumentVerifyPath();
                strPath_DocCompareExe = _formManager.AppStatus_Select("DocCompare");
                //To do 
                //  dblDocCompare = Shell( $"{strPath_DocCompareExe} {strDocPath } {mstrDocName}", WindowStyle:= vbNormalFocus);
           
            
        }


        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }

        private void cmdFinish_Click(object sender, EventArgs e)
        {
            bool blnProtect;
            Cursor = Cursors.WaitCursor;
            try
            {
                if (mblnCalledWizard)
                {
                    UserForm__SHUTDOWN(qReturn.qReturn_Finish);
                    return;
                }

                if (currentActiveDocument.ProtectionType == Word.WdProtectionType.wdNoProtection)
                {
                    blnProtect = false;

                }
                else
                {
                    blnProtect = true;
                    Document_UnProtect();
                }

                if (blnProtect)
                {
                    Document_Protect();
                }

                if (chkMoveToLockedState.Checked)
                {
                    DialogResult result = MessageBox.Show("Are you sure you want '" + mstrDocName + "' moved to a LOCKED state?\r\n\r\nIf you...\r\n\r\n Click [Yes]: This document will be LOCKED in the database, and will no\r\nlonger be editable within Authoring once closed here.\r\n\r\n    Click [No]: This document will NOT be locked in the database and will \r\n  remain editable within Authoring.\r\n\r\n    Click [Cancel]: You'll remain here within the 'Verify Document Wizard'.", "Confirm Move To A Locked State", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3);

                    switch (result)
                    {
                        case DialogResult.Yes:
                            int recordsAffected = _formManager.SetDocumentToLockedState(mDocID);
                            if (recordsAffected != 1)
                            {
                                MessageBox.Show("Unexpected RecordsAffected=" + recordsAffected, "SetDocumentToLockedState", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }

                            break;
                        case DialogResult.No:
                            chkMoveToLockedState.Checked = false;
                            break;
                        case DialogResult.Cancel:
                            return;
                    }
                }

                this.Close();
                CommonService.DocumentProtection_Assure(currentActiveDocument, mstrProtectionPassword);
                Cursor = Cursors.Default;

            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "cmdFinish_Click");

            }
            finally
            {
                CommonService.CleanUpBackgroundWordProcesses();
            }

        }
        private void UserForm__SHUTDOWN(qReturn pqrtQReturn)
        {
            mqrtQReturn = pqrtQReturn;
            if (!mblnCalledWizard)
            {
                MessageBox.Show("ERROR: Called UserForm__SHUTDOWN(" + pqrtQReturn + "), but this is not a CalledWizard.", MyCompName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }


            if (this.Visible)
            {
                this.Hide();
            }

        }

        private void Document_Protect()
        {
            int intBookmarksDeleted;
            if (mblnCalledWizard)
            {
                return;
            }
            CommonService.DocumentProtection_On(currentActiveDocument, mstrProtectionPassword);
        }

        private void Document_UnProtect()
        {
            if (mblnCalledWizard)
            {
                return;
            }

            CommonService.DocumentProtection_Off(currentActiveDocument, mstrProtectionPassword);
        }

        private void lable1_Click(object sender, EventArgs e)
        {
            cmdFinish_Click(sender, e);
        }

        private void label3_Click(object sender, EventArgs e)
        {
            cmdFinish_Click(sender, e);
        }
    }

}
