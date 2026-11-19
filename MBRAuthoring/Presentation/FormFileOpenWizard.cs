using MBRAuthoring.Business;
using MBRAuthoring.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MBRAuthoring.Presentation
{
    public partial class FormFileOpenWizard : FormBase
    {
        public IDocumentManager _documentManager;
        //public IFormManager _formManager;
        protected override Panel progressBarPanel => panel1;

        public const string MyCompName = "FormFileOpenWizard";
        string docTitle_Original;
        int docTypeID;
        string docTypeAbb;
        bool userSelection;
        int opUnitID;
        string opUnitType;
        bool mblnCalledWizard;
        string mstrProtectionPassword = "";
        Microsoft.Office.Interop.Word.Application wordApp;
        public FormFileOpenWizard()
        {
            InitializeComponent();
            _formManager = new FormManager();
            _documentManager = new DocumentManager();
        }

        private void FormFileOpenWizard_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor; // merge
            const string MyProcName = "FormFileOpenWizard_Load";
            //directory clean up
            //string folderPath = @"\\kzoamrfil05\brs\Kalamazoo\MBR\Documents\DP\Production//";
            string folderPath = ConfigurationManager.AppSettings["TemplatePath"];


            try
            {
                System.Collections.Generic.List<string> DeletedtempFiles = new System.Collections.Generic.List<string>();
                System.Collections.Generic.List<string> NewtempFiles = new System.Collections.Generic.List<string>();
                System.Collections.Generic.List<string> ActivetempFiles = new System.Collections.Generic.List<string>();
                string[] tempFiles = Directory.GetFiles(folderPath, "~$*.doc*", SearchOption.TopDirectoryOnly);
                

                foreach (string tempFile in tempFiles)
                {
                    try
                    {
                        FileInfo tempInfo = new FileInfo(tempFile);

                        // Skip recent files
                        if (tempInfo.LastWriteTime > DateTime.Now.AddDays(-1))
                        {
                            NewtempFiles.Add(tempFile);
                            continue;
                        }

                        //if (IsFileLocked(tempFile))
                        //{ 
                        //ActivetempFiles.Add(tempFile);
                        //continue;
                        // }


                        DeletedtempFiles.Add(tempFile);
                        File.Delete(tempFile);
                        
                        //Console.WriteLine($"Deleted: {tempFile}");
                    }
                    catch (IOException)
                    {
                        //Console.WriteLine("Error");
                        // Ignore files that cannot be deleted

                    }
                    catch (UnauthorizedAccessException)
                    {
                        //Console.WriteLine("Error");
                        // Ignore files without sufficient permissions}
                    }
                }

                //directory clean up
                //Todo:
                //Call Form_Accelerators(Me)
                //Call Form_TabIndexing(Me)
                //cmdOpUnitSetDefault.TabStop = false;
                //cmdOpUnitSetDefault.Visible = false;
                lblWizardInfo.Text = "";

                InitializeForm();

                InitializeTabs(tabControl1, cmdBack, cmdNext, tabPage1, tabPage2);
                ShowTab(0);

                lblProject.Visible = false;
                cboProject.Visible = false;
                userSelection = true;

                cmdUpdateDocTitle.Enabled = true;
                cmdUpdateDocTitle.Visible = true;

                SetToolTip(cboOpUnit, "cboOpUnit: OpUnits");  // NewCode
                SetToolTip(cmdCancel, "Terminate Function with NO Output");  // NewCode
                SetToolTip(cmdBack, "Show Previous Wizard Page");  // NewCode
                SetToolTip(cmdNext, "Show Next Wizard Page");  // NewCode
                SetToolTip(cmdFinish, "Open Document, then Terminate this Function");  // NewCode

                SetToolTip(cmdCopyDocTitle, "Copy the DOC TITLE (ALT-D) to the PASTE (Ctrl-V) buffer");
                SetToolTip(cmdCopyFileName, "Copy FILE NAME (ALT-F) to the PASTE(Ctrl-V) buffer");
                SetToolTip(cmdCopyProductID, "Copy the PRODUCT ID (ALT-P) to the PASTE (Ctrl-V) buffer");
                SetToolTip(cmdCopyProductName, "Copy the PRODUCT NAME (ALT-N) to the PASTE (Ctrl-V) buffer");
                SetToolTip(cmdUpdateDocTitle, "Update 'Title' of this Document in Database");
                Cursor = Cursors.Default;
                wordApp = Globals.ThisDocument.ActiveDocument.Application;
                wordApp.StatusBar = "";
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
            finally // merge
            {
                Cursor = Cursors.Default;
            }
        }

        private void FormFileOpenWizard_Activate(object sender, EventArgs e)
        {
            if (cboOpUnit.Items.Count < 1)
            {
                this.Close();
                CommonService.CleanUpBackgroundWordProcesses();
            }
            //SendKeys.SendWait("%");
        }

        private void InitializeForm()
        {
            LoadcboOpUnit();

            foreach (Control ctrl in progressBarPanel.Controls)
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

        private void LoadcboOpUnit() // 1st dropdown
        {
            var opUnits = _formManager.Fill_OpUnit();
            cboOpUnit.DataSource = opUnits;
            cboOpUnit.DisplayMember = "OpUnit";
            cboOpUnit.ValueMember = "OpUnitID";

            cboOpUnit.SelectedIndexChanged += cboOpUnit_SelectedIndexChanged;

            if (cboOpUnit.Items.Count == 0)
            {
                MessageBox.Show("NO OPUNIT LIST FOR THIS USER" + "\r\n" + "\r\n" + "USERNAME=" + CommonService.gstrUserName + "   USERID=" + CommonService.gintUserID + "\r\n" + "\r\n" + "FUNCTION TERMINATING", "FileOpen - No OpUnit List", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (CommonService.gblnUserType_ADMIN)
            {
                lblFileOpen.Text = lblFileOpen.Text + " - ADMIN";
            }
            LoadcboDocType();
            LoadcboSelectFile();

        }

        private void LoadcboSelectFile()
        {
            bool IncludeLocked;
            bool IncludeAuthoring;
            bool IncludeApproved;
            bool IncludeArchived;
            bool IncludeBlocked;
            try
            {

                if (CommonService.gblnUserType_ADMIN)
                {
                    IncludeLocked = true;
                }
                else
                {
                    IncludeLocked = false;
                }

                IncludeAuthoring = true;
                IncludeApproved = false;
                IncludeArchived = false;
                IncludeBlocked = false;

                TblOpUnit opUnit = (TblOpUnit)cboOpUnit.SelectedItem;
                OpUnitDocType opUnitDocType = (OpUnitDocType)cboDocType.SelectedItem;

                SaveManager.SetOpUnit(opUnit, opUnitDocType);//added // 2301

                var files = _formManager.Fill_Document(opUnit.OpUnitID, opUnitDocType.DocTypeID, IncludeLocked, IncludeAuthoring, IncludeApproved, IncludeArchived, IncludeBlocked);
                if (files.Count > 0)
                {

                    cboSelectFile.DisplayMember = "DisplayDocumentName";
                    cboSelectFile.ValueMember = "DocumentName";
                    cboSelectFile.DataSource = files;

                    cboSelectFile.DrawMode = DrawMode.OwnerDrawFixed;
                    cboSelectFile.DrawItem += cboSelectFile_DrawItem;
                    cboSelectFile.Format += cboSelectFile_Format;
                    SetToolTip(cboSelectFile, "cboSelectFile: OpUnitID=" + opUnit.OpUnitID + " OpUnitType=" + opUnit.OpUnitType + " DocTypeID=" + opUnitDocType.DocTypeID + " InclLocked=" + IncludeLocked);  // NewCode
                    //cboSelectFile.Format += (cboSelectFileSender, args) => cboBox_Format<DocumentDetails>(cboSelectFileSender, args, new List<string> { "DocumentName", "DocumentTitle" }, new List<int> { 50, 100 });
                }

            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "LoadcboSelectFile");
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "LoadcboSelectFile");
                }
            }
        }

        private void cboSelectFile_Format(object sender, ListControlConvertEventArgs e)
        {
            var documentDetails = (DocumentDetails)e.ListItem;
            e.Value = documentDetails.DisplayDocumentName;
        }

        private void cboSelectFile_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            if (e.Index >= 0)
            {
                var documentDetails = (DocumentDetails)cboSelectFile.Items[e.Index];

                System.Drawing.Font drawFont = e.Font;
                int padding = 180;  // Space between DocumentName and DocumentTitle columns

                e.Graphics.DrawString(documentDetails.DisplayDocumentName, drawFont, Brushes.Black, e.Bounds.X, e.Bounds.Y);
                e.Graphics.DrawString(documentDetails.DocumentTitle, drawFont, Brushes.Black, e.Bounds.X + padding, e.Bounds.Y);
            }

            e.DrawFocusRectangle();
        }

        private void LoadcboDocType() // 2nd dropdown
        {
            TblOpUnit opUnit = (TblOpUnit)cboOpUnit.SelectedItem;
            var opUnitDocTypes = _formManager.Fill_OpUnit_DocType(opUnit.OpUnitID, opUnit.OpUnitType);

            cboDocType.DisplayMember = "DocType";
            cboDocType.ValueMember = "DocTypeID";
            cboDocType.DataSource = opUnitDocTypes;
            SetToolTip(cboDocType, "cboDocType: OpUnitID=" + opUnit.OpUnitID + " OpUnitType=" + opUnit.OpUnitType + " OpUnitDocTypes");  // NewCode
        }

        private void cboOpUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string MyProcName = "cboOpUnit_SelectedIndexChanged";

            try
            {
                TblOpUnit opUnit = (TblOpUnit)cboOpUnit.SelectedItem;

                if (cboOpUnit.Items.Count == 0)
                {
                    return;
                }

                opUnitID = opUnit.OpUnitID;
                opUnitType = opUnit.OpUnitType;
                if (userSelection)
                {
                    userSelection = false;
                }
                LoadcboDocType();
                LoadcboSelectFile();
                userSelection = true;

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

        private void cboDocType_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string MyProcName = "cboDocType_SelectedIndexChanged";
            string strStepName;

            try
            {

                OpUnitDocType opUnitDocType = (OpUnitDocType)cboDocType.SelectedItem;
                if (cboDocType.Items.Count == 0)
                {
                    return;
                }
                docTypeID = opUnitDocType.DocTypeID;
                docTypeAbb = opUnitDocType.DocTypeAbb;
                if (docTypeID == 11)
                {
                    txtDocTitle.ReadOnly = false;
                    cmdUpdateDocTitle.Enabled = true;
                    cmdUpdateDocTitle.Visible = true;
                }
                else
                {
                    txtDocTitle.ReadOnly = true;
                    cmdUpdateDocTitle.Enabled = false;
                    cmdUpdateDocTitle.Visible = false;
                }
                if (userSelection == true)
                {
                    LoadcboSelectFile();
                }
                //Todo: But strStepName value is never used.
                //strStepName = Boomark_IntroEnd_StepTextName(pblnCleaningRecord:= (mstrDocTypeAbb = "CLR"))

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

        private async void cmdFinish_Click(object sender, EventArgs e)
        {
            const string MyProcName = "cmdFinish_Click";  // NewCode
            Cursor = Cursors.WaitCursor; // merge
            Document openedDocument = null;
            Template temp = null;
            Template wtpTemplate = null;
            Find fndProcessingSteps = null;
            bool blnFound = false;
            bool documentClosed = false;
            string documentName = null;
            Microsoft.Office.Interop.Word.Application wordApp = Globals.ThisDocument.ActiveDocument.Application;
            try
            {
                if (wordApp.ActiveWindow.View.SplitSpecial == WdSpecialPane.wdPaneNone)
                {
                    wordApp.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
                    wordApp.ActiveWindow.ActivePane.View.Zoom.PageFit = WdPageFit.wdPageFitBestFit;
                }
                else
                {
                    wordApp.ActiveWindow.View.Type = WdViewType.wdNormalView;
                    wordApp.ActiveWindow.View.Zoom.PageFit = WdPageFit.wdPageFitBestFit;
                }

                wordApp.ActiveWindow.View.ShowBookmarks = true;

                if (cboSelectFile.Items.Count == 0)
                {
                    MessageBox.Show("No document to open.", "Warning - File Open Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    wordApp.StatusBar = ""; // merge
                    return;
                }
                else
                {
                    string strExpectedTemplateFullName = CommonService.templateFullName;
                    DocumentInfo documentInfo = new DocumentInfo();

                    var document = (DocumentDetails)cboSelectFile.SelectedItem;
                    string docPath = document.DocumentPath;
                    int docID = document.DocID;
                    documentName = document.DocumentName;

                    if (!docPath.EndsWith("\\"))
                    {
                        docPath += "\\";
                    }

                    string directory = docPath;
                    string fileNameWithPath = directory + document.DocumentName;
                    //performancetuning - Fix for unavailable documents and closing dummy word instance
                    if (File.Exists(fileNameWithPath))
                    {
                        wordApp.StatusBar = "Opening File " + document.DocumentName;

                        openedDocument = _documentManager.OpenDocument(fileNameWithPath, document.DocPasswordRead, document.DocPasswordWrite, document.DocumentName);
                    }
                    else
                    {
                        MessageBox.Show("Selected document is not available.", "Warning - File Open Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        wordApp.StatusBar = ""; // merge
                        Cursor = Cursors.Default;
                        Close();
                        return;
                    }
                    //end


                    if (openedDocument != null) //added to avoid null exception
                    {
                        openedDocument.Application.StatusBar = "Please wait...";
                        openedDocument.Application.System.Cursor = WdCursorType.wdCursorWait;
                        Globals.ThisDocument.ActiveDocument = openedDocument;
                        temp = openedDocument.get_AttachedTemplate();
                        string strDocumentTemplateFullName = temp.FullName;

                        if (strDocumentTemplateFullName != CommonService.templateFullName)
                        {
                            // merge
                            // openedDocument.Application.Visible = true;
                            //lblWizardInfo.Text = "From the open document, choose \"Open a Read Only copy\" and proceed...";
                            lblWizardInfo.Text = "Chg Tpt from '" + strDocumentTemplateFullName + "' to '" + CommonService.templateFullName + "'";  // NewCode
                            openedDocument.set_AttachedTemplate(CommonService.templateFullName);

                            wtpTemplate = openedDocument.get_AttachedTemplate();
                            if (wtpTemplate.FullName != CommonService.templateFullName)
                            {
                                CommonService.Error_Handler("0", "Failed to attach Template '" + CommonService.templateFullName + "'", MyCompName, MyProcName);
                            }

                            CommonService.Style_Copy_All(openedDocument, _documentManager);
                            openedDocument.Saved = true;
                            lblWizardInfo.Text = "";
                        }

                        string docName = openedDocument.Application.ActiveDocument.Name;
                        documentInfo = _documentManager.Document_SelectQPSObject(docName, documentInfo, true);
                        string strPasswordProtect = documentInfo.GetClsDocument.DocPasswordProtect;
                        bool blnDocTrackChgsRequired = documentInfo.GetClsDocument.mblnTrackChangesReq;
                        bool blnReTrackChanges = blnDocTrackChgsRequired;
                        bool blnReProtect = true;
                        mstrProtectionPassword = strPasswordProtect;
                        if (!Globals.ThisDocument.docDetails.ContainsKey(docName)) // 2301
                        {
                            Globals.ThisDocument.docDetails.Add(docName, strPasswordProtect);
                        }
                        CommonService.DocumentProtection_Off(openedDocument, strPasswordProtect);
                        openedDocument.TrackRevisions = false;
                        openedDocument.Application.StatusBar = "Add-ins Loading..."; // merge
                        openedDocument.Application.System.Cursor = WdCursorType.wdCursorWait; // merge
                        

                        if (_documentManager.DocEqualsLastSave(docName) == false)
                        {
                            openedDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
                            documentClosed = true;
                            ReleaseComObject(ref openedDocument);
                            CommonService.CleanUpBackgroundWordProcesses();
                            wordApp.StatusBar = "";
                            wordApp.System.Cursor = WdCursorType.wdCursorNormal;
                            Close();
                            return;
                        }
                        openedDocument.Application.Activate();  // NewCode
                        TblOpUnit opunit = (TblOpUnit)cboOpUnit.SelectedItem;
                        OpUnitDocType opUnitDocType = (OpUnitDocType)cboDocType.SelectedItem;
                        Globals.ThisDocument.FileOpenDocumentName(openedDocument.Name);
                        Globals.ThisDocument.CommandBar_Build(opunit.OpUnitType, opUnitDocType.DocTypeAbb, opunit.OpUnitID, opUnitDocType.DocTypeID);
                        
                        if (chkFindProcessingSteps.Checked)
                        {

                            fndProcessingSteps = openedDocument.Application.Selection.Find;
                            QPS_Find qpfFindSaveArea = CommonService.Find_Save(fndProcessingSteps);
                            CommonService.Find_Setup(fndProcessingSteps);
                            fndProcessingSteps.Text = txtFindString.Text;
                            blnFound = fndProcessingSteps.Execute();
                            if (blnFound)
                            {
                                openedDocument.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                            }
                            CommonService.Find_Restore(fndProcessingSteps, qpfFindSaveArea);
                        }
                        _formManager.NotifyUser_UnIncludedMaterials(docID);
                        if (blnReTrackChanges)
                        {
                            openedDocument.TrackRevisions = true;
                        }
                        if (blnReProtect)
                        {
                            CommonService.DocumentProtection_On(openedDocument, "");
                        }

                        //openedDocument.Application.Activate(); //Recheck
                    }
                }
                // merge
                Close();
                wordApp.StatusBar = "";
                wordApp.System.Cursor = WdCursorType.wdCursorNormal;
                if (openedDocument != null)
                {
                    openedDocument.Application.StatusBar = "";
                    openedDocument.Application.System.Cursor = WdCursorType.wdCursorWait;
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                Close();
                if (openedDocument != null)
                {
                    openedDocument.Application.StatusBar = "";
                    openedDocument.Application.System.Cursor = WdCursorType.wdCursorWait;
                }
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    wordApp.StatusBar = "";  // NewCode
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                }
            }
            finally
            {
                ReleaseComObject(ref fndProcessingSteps);
                ReleaseComObject(ref wtpTemplate);
                ReleaseComObject(ref temp);

                if (documentClosed)
                {
                    ReleaseComObject(ref openedDocument);
                }

                CommonService.CleanUpBackgroundWordProcesses();
            }
        }

        private void cmdBack_Click(object sender, EventArgs e)
        {
            ShowPrevTab();
        }

        private void cmdNext_Click(object sender, EventArgs e)
        {
            ShowNextTab();
            if (docTypeID == 11)
            {
                cmdUpdateDocTitle.Enabled = true;
                cmdUpdateDocTitle.Visible = true;
            }
            else
            {
                cmdUpdateDocTitle.Enabled = false;
                cmdUpdateDocTitle.Visible = false;
            }
            lblUpdated.Text = "";

            tabControl_Selected();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }

        private void cmdCopyFileName_Click(object sender, EventArgs e)
        {
            txtDocFileName.Focus();
            txtDocFileName.SelectAll();
            SendKeys.Send("^c");
        }

        private void cmdCopyDocTitle_Click(object sender, EventArgs e)
        {
            txtDocTitle.Focus();
            txtDocTitle.SelectAll();
            SendKeys.Send("^c");
        }

        private void cmdCopyProductID_Click(object sender, EventArgs e)
        {
            txtDocProductID.Focus();
            txtDocProductID.SelectAll();
            //Clipboard.SetText(txtDocProductID.Text);
            SendKeys.Send("^c");
        }

        private void cmdCopyProductName_Click(object sender, EventArgs e)
        {
            txtDocProductName.Focus();
            txtDocProductName.SelectAll();
            //Clipboard.SetText(txtDocProductName.Text);
            SendKeys.Send("^c");
        }

        private void cmdUpdateDocTitle_Click(object sender, EventArgs e)
        {
            string myCompName = "cmdUpdateDocTitle_Click";

            var value = (DocumentDetails)cboSelectFile.SelectedItem;
            if (docTitle_Original.Trim() == txtDocTitle.Text.Trim())
            {
                MessageBox.Show("Update rejected - no change was made to Document Title", "Update Document Title", MessageBoxButtons.OK);
                return;
            }
            int rowAffected = _formManager.Document_UpdateTitle(value.DocID, txtDocTitle.Text);
            if (rowAffected != 1)
            {
                MessageBox.Show("Unexpected RecordsAffected=" + rowAffected, myCompName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                lblUpdated.Text = "DocTitle UPDATED in database";
            }

            _formManager.DocProperty_Apply(value.DocID, 0, "DocumentTitle", txtDocTitle.Text, "", -1, -1);

        }

        private void tabControl_Selected()
        {
            if (cboSelectFile.SelectedIndex < 0)
            {
                txtDocFileName.Text = "";
                txtDocTitle.Text = "";
                txtDocProductID.Text = "";
                txtDocProductName.Text = "";
            }
            else
            {
                var value = (DocumentDetails)cboSelectFile.SelectedItem;
                var documentDetails = _formManager.FillDocument_Details(value.DocID);

                txtDocFileName.Text = documentDetails.DocumentName;
                txtDocTitle.Text = documentDetails.DocumentTitle;
                txtDocProductID.Text = documentDetails.ItemCode;
                txtDocProductName.Text = documentDetails.MaterialName;
            }
            docTitle_Original = txtDocTitle.Text;
        }

        private void cboProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cboProject.Items.Count > 0)
                {
                    if (userSelection)
                    {
                        LoadcboSelectFile();
                    }
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "cboProject_Change");
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "cboProject_Change");
                }
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {
            cmdFinish_Click(sender, e);
        }

        private void label3_Click(object sender, EventArgs e)
        {
            cmdFinish_Click(sender, e);
        }

        private void shp1_Click(object sender, EventArgs e)
        {
            cmdNext_Click(sender, e);
        }

        private void label1_Click(object sender, EventArgs e)
        {
            cmdNext_Click(sender, e);
        }

        private void FormFileOpenWizard_FormClosing(object sender, FormClosingEventArgs e) // merge
        {
            mblnCalledWizard = true;
            if (wordApp != null && wordApp.ActiveDocument != null)
            {
                _formManager.Wizard_Terminate(wordApp.ActiveDocument, mblnCalledWizard, mstrProtectionPassword);
            }
            CommonService.CleanUpBackgroundWordProcesses();

        }

        private static void ReleaseComObject<T>(ref T comObject) where T : class
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch
            {
            }
            finally
            {
                comObject = null;
            }
        }

        //Not Req
        //private void FormFileOpenWizard_FormClosed(object sender, FormClosedEventArgs e)
        //{
        //    wordApp.Application.System.Cursor = WdCursorType.wdCursorWait;
        //    Dispose();
        //    wordApp.Application.System.Cursor = WdCursorType.wdCursorNormal;
        //}
        private static bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(
                    filePath,
                    FileMode.Open,
                    FileAccess.ReadWrite,
                    FileShare.None))
                {
                    // Successfully obtained exclusive access
                }

                return false;
            }
            catch (IOException)
            {
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return true;
            }
        }
    }
}
