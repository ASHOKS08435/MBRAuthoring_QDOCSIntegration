using MBRAuthoring.Business;
using MBRAuthoring.DataAccess;
using MBRAuthoring.Models;
using MBRAuthoring.Presentation;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static MBRAuthoring.Presentation.FormEditProperty;
using Application = Microsoft.Office.Interop.Word.Application;
using Forms = System.Windows.Forms;
using Timer = System.Timers.Timer;
using Word = Microsoft.Office.Interop.Word;

namespace MBRAuthoring
{
    public partial class ThisDocument
    {
        public Dictionary<string, string> docDetails = new Dictionary<string, string>();

        public Timer timer;
        IDocumentManager _documentManager = new DocumentManager();
        FormFileOpenWizard formFileOpenWizard; // merge
        FormEditMaterials formEditMaterials;
        FormEditModule formEditModule;
        FormEditProperty formEditProperty;
        FormSections sections;
        FormHistory formhistory;
        FormPageSetup formpageSetup;
        FormProtection formprotection;
        FormVerifyWizard formverifyWizard;
        FormBlankStepMBR formBlankStepMBR;
        FormInsertTable formInsertTable;
        FormInsertSignDateCrossRef formInsertSignDateCrossRef;
        FormInsertRespBox formInsertRespBox;
        FormInsertField formInsertField;
        private const string QPSNT_COMMANDBARNAME = "QPSNT";//"MBR"-->Facing condition failed issue while changing to MBR . 
        public Document ActiveDocument;
        public Application GlobalApplication;
        public const string BUILDFOR_TEMPLATE = "TEMPLATE";
        public Dictionary<string, CommandBarButton> _buttonMap;
        public CommandBarButton cbnCommandBarButton;
        public CommandBarPopup cpoCommandBarPopup;
        public CommandBar cbrMain;
        public CommandBar cbrSub;
        public CommandBar cbrSubSub;
        public CommandBar cbrScan = null;

        public bool buildCommandBar = true;

        private List<CommandBarButton> cmb = new List<CommandBarButton>(); // new

        public string strCommandBarName;
        public string strCommandBarScan;
        public bool blnBuildCommandBar;

        public CommandBarButton cbnFileOpen;
        public CommandBarButton cbnEditMaterials;
        public CommandBarButton cbnEditModuleList;
        public CommandBarButton cbnEditProperty;
        public CommandBarButton cbnRefresh;

        private const string MENUTYPE_WIZARD = "W";
        private const string MENUTYPE_WIZARD_SPECCHAR = "WSC";
        private const string MENUTYPE_UTILITY = "U";
        private const string MENUTYPE_FILENEW = "N";
        private const string MENUTYPE_FILEOPEN = "O";
        private const string MENUTYPE_WORDPERFECT = "P";
        private const string Update_All_Fields = "Update All Fields";
        private const string Insert_Cross_Ref = "Insert Step Cross Reference";
        private const string Insert_Response_Box = "Insert Response Box";

        private const string AppAcronym = "MBR";
        private static string mstrFileOpenDocumentName = string.Empty;

        string mstrDocName;
        string mstrDocFullName;
        int mDocID;
        string mstrProtectionPassword;
        string mstrOpUnit;
        bool mblnCalledWizard;
        int mintOpUnitID;
        string mstrOpUnitType;
        int mintDocTypeID;
        MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
        Table tabProperty;
        Table tabModule;
        System.Data.DataTable dtDocProperty = new System.Data.DataTable();
        System.Data.DataTable dtDocModule = new System.Data.DataTable();
        System.Data.DataTable dtDocMaterial = new System.Data.DataTable();
        Range mrngHold;
        Range mrngYield;

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }


        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            try
            {

                ActiveDocument = Application.ActiveDocument;
                ActiveDocument.Application.System.Cursor = WdCursorType.wdCursorWait;

                DocumentVariableAssignments(true, true);
                CommandBar_DeleteAll();
                if (buildCommandBar)
                {
                    CommandBar_Build(BUILDFOR_TEMPLATE, BUILDFOR_TEMPLATE);
                }

                Template temp = (Template)AttachedTemplate;
                temp.Saved = true;
            }
            catch (Exception ex)
            {
                //err line
            }
        }

        private void DocumentVariableAssignments(bool pblnAuthInfo, bool MarkAsSaved, string AppVersionNumber = "", string AppAcronym = "", string AppDate = "")
        {
            Table tblMain;
            Cell celMessage;
            try
            {
                tblMain = ActiveDocument.Tables[1];
                if (AppVersionNumber == "")
                {
                    ModChangeHistory.gstrAppVersionNumber = tblMain.Cell(2, 3).Range.Text;
                    ModChangeHistory.gstrAppVersionNumber = ModChangeHistory.gstrAppVersionNumber.Substring(0, ModChangeHistory.gstrAppVersionNumber.Length - 2);
                }
                else
                {
                    ModChangeHistory.gstrAppVersionNumber = AppVersionNumber;
                }

                if (AppAcronym == "")
                {
                    ModChangeHistory.gstrAppAcronym = tblMain.Cell(1, 2).Range.Text;
                    ModChangeHistory.gstrAppAcronym = ModChangeHistory.gstrAppAcronym.Substring(0, ModChangeHistory.gstrAppAcronym.IndexOf(" "));
                }
                else
                {
                    ModChangeHistory.gstrAppAcronym = AppAcronym;
                }

                if (AppDate == "")
                {
                    ModChangeHistory.gstrAppDate = tblMain.Cell(3, 3).Range.Text;
                    ModChangeHistory.gstrAppDate = ModChangeHistory.gstrAppDate.Substring(0, ModChangeHistory.gstrAppDate.Length - 2);

                }
                else
                {
                    ModChangeHistory.gstrAppDate = AppDate;
                }

                celMessage = tblMain.Cell(2, 3);
                celMessage.Range.Text = ModChangeHistory.gstrAppVersionNumber;
                ModChangeHistory.gstrAppVersionMajor = ModChangeHistory.gstrAppVersionNumber.Substring(0, 1);

                celMessage = tblMain.Cell(1, 2);
                celMessage.Range.Text = "";
                celMessage.Range.Text = ModChangeHistory.gstrAppAcronym + " Authoring";

                celMessage = tblMain.Cell(3, 3);
                celMessage.Range.Text = ModChangeHistory.gstrAppDate;

                celMessage = tblMain.Cell(tblMain.Rows.Count, 1);
                celMessage.Range.Text = " ";
                celMessage.Select(); // Select the cell

                Selection selection = ThisApplication.Selection;
                selection.Collapse(WdCollapseDirection.wdCollapseStart);

                if (pblnAuthInfo)
                {
                    CommonService.TypeStepText(StepTextType.qtImmediate, "AuthorInfoMessage");
                }

                if (MarkAsSaved)
                {
                    ActiveDocument.Saved = true;
                }

            }
            catch (Exception ex)
            {
                ActiveDocument.Application.StatusBar = "";
            }
            finally
            {
                ActiveDocument.Application.System.Cursor = WdCursorType.wdCursorNormal;
                ActiveDocument.Application.StatusBar = "";
            }

        }

        public void CommandBar_DeleteAll()
        {
            int commandBarNameLength = QPSNT_COMMANDBARNAME.Length;

            foreach (CommandBar commandBar in ActiveDocument.Application.CommandBars)
            {
                string commandBarName = commandBar.Name;

                if (commandBarName.Length >= commandBarNameLength)
                {
                    if (commandBarName.StartsWith(QPSNT_COMMANDBARNAME))
                    {
                        commandBar.Delete();
                    }
                }
            }
        }

        public void CommandBar_Build(string pstrOpUnitType, string pstrDocTypeAbb, int OpUnitID = -1, int DocTypeID = -1)
        {
            //const string MyProcName = "CommandBar_Build";
            try
            {

                switch (pstrOpUnitType)
                {
                    case "TEMPLATE":
                        strCommandBarName = QPSNT_COMMANDBARNAME;
                        break;
                    case "DPR":
                    case "PKG":
                    default:
                        strCommandBarName = QPSNT_COMMANDBARNAME + pstrOpUnitType + pstrDocTypeAbb;
                        break;
                }
                blnBuildCommandBar = true;

                foreach (CommandBar cbMain in ActiveDocument.Application.CommandBars)
                {
                    if (cbMain.Name == strCommandBarName)
                    {
                        blnBuildCommandBar = false;
                        break;
                    }
                }
                if (blnBuildCommandBar)
                {
                    cbrMain = ActiveDocument.Application.CommandBars.Add(strCommandBarName);
                    CommandBar_Button_Add(cbrMain, "File &New", FileNew_MBR);
                    cbnFileOpen = (CommandBarButton)CommandBar_Button_Add(cbrMain, "File O&pen", FileOpen_MBR, true);
                    cbrSub = CommandBar_Button_Add(cbrMain, "MBR He&lp", null, true);
                    CommandBar_Button_Add(cbrSub, "About MBR", AboutApp, false);
                    cbrSubSub = CommandBar_Button_Add(cbrSub, "MBR Features", null);
                    CommandBar_Button_Add(cbrSubSub, "MBR Help", wzd_MBRHelp);

                    if (pstrOpUnitType != "TEMPLATE")
                    {
                        CommandBar_Build_Wizards(cbrMain, pstrOpUnitType, pstrDocTypeAbb, OpUnitID, DocTypeID);
                    }
                }

                CommandBar_Visible(strCommandBarName);
                //if (strCommandBarName != QPSNT_COMMANDBARNAME)
                //{
                //    foreach (CommandBar cbScan in ActiveDocument.Application.CommandBars)
                //    {
                //        strCommandBarScan = cbScan.Name;
                //        if (strCommandBarScan == strCommandBarPattern && strCommandBarScan != strCommandBarName)
                //        {
                //            cbScan.Visible = false;
                //        }
                //    }
                //}

                //Template attachedTemplate = ActiveDocument.get_AttachedTemplate();
                //attachedTemplate.Saved = true; recheck
            }
            catch (Exception ex)
            {
                Forms.MessageBox.Show("ERROR (" + ex.HResult + "): " + ex.Message, "CommandBar_Build", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);
            }
        }

        public void CommandBar_Visible(string strCommandBarName, long position = 1, MsoBarRow RowIndex = MsoBarRow.msoBarRowLast, long Left = 160, bool Visible = true)
        {
            bool commandBarExists = false;

            foreach (CommandBar bar in ActiveDocument.CommandBars)
            {
                if (bar.Name == strCommandBarName)
                {
                    commandBarExists = true;
                }
            }

            if (commandBarExists)
            {
                cbrScan = GetCommandBar(ActiveDocument.Application, strCommandBarName);
                if (cbrScan != null)
                {
                    cbrScan.Position = (MsoBarPosition)position;
                    cbrScan.RowIndex = (int)RowIndex;
                    cbrScan.Left = 0; // 160 or 0 Recheck
                    cbrScan.Visible = Visible;
                }
            }
        }

        public CommandBar GetCommandBar(Application app, string CommandBarName)
        {
            foreach (CommandBar bar in app.CommandBars)
            {
                if (bar.Name == CommandBarName)
                {
                    return bar;
                }
            }
            return null;
        }

        private void wzd_MBRHelp(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            FormMBRHelp formMBRHelp = new FormMBRHelp();
            formMBRHelp.ShowDialog();
        }

        private CommandBar CommandBar_Button_Add(CommandBar pcbrMain, string pstrCaption, _CommandBarButtonEvents_ClickEventHandler pstrOnAction = null, bool pblnGroup = false)
        {

            if (pstrOnAction == null)
            {
                cpoCommandBarPopup = (CommandBarPopup)pcbrMain.Controls.Add(MsoControlType.msoControlPopup);
                cpoCommandBarPopup.Caption = pstrCaption;
                cpoCommandBarPopup.BeginGroup = pblnGroup;
                return cpoCommandBarPopup.CommandBar;
            }
            else
            {
                CommandBarButton newButton = (CommandBarButton)pcbrMain.Controls.Add(MsoControlType.msoControlButton);
                //cbnCommandBarButton = (CommandBarButton)pcbrMain.Controls.Add(MsoControlType.msoControlButton);
                newButton.Caption = pstrCaption;
                newButton.BeginGroup = pblnGroup;
                //cbnCommandBarButton.OnAction = pstrOnAction;
                newButton.Click += pstrOnAction;
                newButton.Style = MsoButtonStyle.msoButtonCaption;
                cmb.Add(newButton);
                return null;
            }
        }

        private void CommandBar_Build_Wizards(CommandBar pcbrMain, string pstrOpUnitType, string pstrDocTypeAbb, int OpUnitID = 0, int DocTypeID = 0)
        {
            bool ProjectNameUsed;
            bool DocIsModule;
            bool DocHasModule;
            bool DocHasMaterial;
            bool DocHasWorkCenter;
            bool DocHasProperty;
            bool DocHasWorkCenterAlt;
            bool DocHasProduct;

            cbrSub = CommandBar_Button_Add(pcbrMain, "Wi&zards", null, true);
            CommandBar_Button_Add(cbrSub, "Blank Step", wzd_BlankStep);
            cbrSubSub = CommandBar_Button_Add(cbrSub, "Fields (Properties)", null);
            CommandBar_Button_Add(cbrSubSub, "Field (Property)", wzd_InsertField);
            CommandBar_Button_Add(cbrSubSub, Update_All_Fields, wzd_UpdateField);
            CommandBar_Button_Add(cbrSub, Insert_Response_Box, wzd_ResponseBox);
            CommandBar_Button_Add(cbrSub, Insert_Cross_Ref, wzd_InsertSignDateCrossRef);
            CommandBar_Button_Add(cbrSub, "Insert Table", wzd_InsertTable);

            OpUnitDocType opUnitDocType = DatabaseHelper.OpUnitDocType_SelectOne(OpUnitID, DocTypeID);

            ProjectNameUsed = opUnitDocType.ProjectNameUsed != "N";
            DocIsModule = opUnitDocType.DocIsModule == "Y";
            DocHasModule = opUnitDocType.DocHasModule != "N";
            DocHasMaterial = opUnitDocType.DocHasMaterial != "N";
            DocHasWorkCenter = opUnitDocType.DocHasWorkCenter != "N";
            DocHasWorkCenterAlt = opUnitDocType.DocHasWorkCenterAlt != "N";
            DocHasProduct = opUnitDocType.DocHasProduct != "N";
            DocHasProperty = opUnitDocType.DocHasProperty != "N";


            cbrSub = CommandBar_Button_Add(pcbrMain, "&Utilities", null, true);
            if (CommonService.gblnUserType_ADMIN)
            {
                CommandBar_Button_Add(cbrSub, "Manage &Protection/TrackChanges", utl_ManageProtection);
                CommandBar_Button_Add(cbrSub, "Manage Se&ctions", utl_ManageSections);
            }
            else
            {
                CommandBar_Button_Add(cbrSub, "Manage Se&ctions", utl_ManageSections);
                CommandBar_Button_Add(cbrSub, "Highlight &Changes", utl_ManageProtection);
            }

            if (DocHasMaterial)
            {
                cbnEditMaterials = (CommandBarButton)CommandBar_Button_Add(cbrSub, "Edit &Materials", utl_EditMaterials);
            }

            //Not req
            //if (DocHasWorkCenter)
            //{
            //    CommandBar_Button_Add(cbrSub, "Edit &Work Centers", utl_EditWC);
            //}

            if (DocHasModule)
            {
                cbnEditModuleList = (CommandBarButton)CommandBar_Button_Add(cbrSub, "Edit &ModuleList", utl_EditModule);
            }

            if (DocHasProperty)
            {
                cbnEditProperty = (CommandBarButton)CommandBar_Button_Add(cbrSub, "Edit &Property", utl_EditProperty);
            }
            if (DocHasMaterial || DocHasModule || DocHasProperty)
            {
                cbrSubSub = CommandBar_Button_Add(cbrSub, "Refresh &Document", null);
                if (DocHasMaterial)
                    CommandBar_Button_Add(cbrSubSub, "Refresh &Materials", utl_RefreshMaterials);
                if (DocHasModule)
                    CommandBar_Button_Add(cbrSubSub, "Refresh &ModuleList", utl_RefreshModules);
                if (DocHasProperty)
                    CommandBar_Button_Add(cbrSubSub, "Refresh &Property", utl_RefreshProperties);
            }
            CommandBar_Button_Add(cbrSub, "Document &Verify Utility", utl_VerifyDocument);
            CommandBar_Button_Add(cbrSub, "Show 'Save' &History", utl_DocSaveHistory);
            CommandBar_Button_Add(cbrSub, "Pa&ge Setup", utl_PageSetup);
        }

        private void utl_ManageProtection(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
            {
            commandBarButton.Application.StatusBar = "Form Loading...";
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (!MenuItemAvailable(MENUTYPE_UTILITY)) { return; }

            if (formprotection == null || formprotection.IsDisposed)
                formprotection = new FormProtection();
            Set_Parent_AND_ShowDialog(formprotection, ActiveDocument);
            //formprotection.ShowDialog();
            formprotection.Dispose();
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                formprotection = null;
                throw;
            }

        }


        private void wzd_InsertTable(CommandBarButton commandBarButton, ref bool CancelDefault)
        {
            try
            {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            // if (!MenuItemAvailable(MENUTYPE_WIZARD)) { return; }
            if (ActiveDocument.Application.Selection.Information[WdInformation.wdWithInTable])
            {
                MessageBox.Show("ERROR: Cursor-Selection must NOT be within a TABLE.\n\nPlease re-position cursor and try again.", "Menu Validation - Wizard", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            if (formInsertTable == null || formInsertTable.IsDisposed)
                formInsertTable = new FormInsertTable();
            //Set_Parent_AND_ShowDialog(formInsertTable, ActiveDocument);
            formInsertTable.ShowDialog();
            formInsertTable.Dispose();
            }
            catch(Exception ex)
            {
                formInsertTable = null;
                throw;
            }


        }

        private void wzd_InsertSignDateCrossRef(CommandBarButton commandBarButton, ref bool CancelDefault)
        {
            try
            {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            // if (!MenuItemAvailable(MENUTYPE_WIZARD)) { return; }
            if (formInsertSignDateCrossRef == null || formInsertSignDateCrossRef.IsDisposed)
                formInsertSignDateCrossRef = new FormInsertSignDateCrossRef();
            Set_Parent_AND_ShowDialog(formInsertSignDateCrossRef, ActiveDocument);
            //formInsertSignDateCrossRef.ShowDialog();
            formInsertSignDateCrossRef.Dispose();
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                formInsertSignDateCrossRef = null;
                throw;
            }

        }

        private void wzd_ResponseBox(CommandBarButton commandBarButton, ref bool CancelDefault)
        {
            try
            {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            // if (!MenuItemAvailable(MENUTYPE_WIZARD)) { return; }
            if (ActiveDocument.Application.Selection.Information[WdInformation.wdWithInTable])
            {
                MessageBox.Show("ERROR: Cursor-Selection must NOT be within a TABLE.\n\nPlease re-position cursor and try again.", "Menu Validation - Wizard", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            if (formInsertRespBox == null || formInsertRespBox.IsDisposed)
                formInsertRespBox = new FormInsertRespBox();
            Set_Parent_AND_ShowDialog(formInsertRespBox, ActiveDocument);
            //formInsertRespBox.ShowDialog();
            formInsertRespBox.Dispose();
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                formInsertRespBox = null;
                throw;
            }

        }

        private void wzd_UpdateField(CommandBarButton commandBarButton, ref bool CancelDefault)
        {
            try
            {
            int ErrorCount = 0;
            const string FieldError = "Error! Unknown Document Property Name.";
            string ErrorText;
            if (!MenuItemAvailable(MENUTYPE_UTILITY))
            {
                return;
            }
            ActiveDocument.Fields.Update();
            foreach (Field fld in ActiveDocument.Fields)
            {
                if (fld.Result.Text != null)
                {

                    if (fld.Result.Text.ToUpper() == FieldError.ToUpper())
                    {
                        ErrorCount = ErrorCount + 1;
                    }
                }
            }

            if (ErrorCount == 0)
            {
                ErrorText = "NO";
            }
            else
            {
                ErrorText = ErrorCount.ToString();
            }
            System.Windows.Forms.MessageBox.Show("All Fields have been Updated" + "\r\n" + "\r\n" +
                "Document " + ActiveDocument.Name + "\r\n" + "\r\n" + "has " + ActiveDocument.Fields.Count +
                " Fields" + "\r\n" + "\r\n" +
                "and " + ActiveDocument.CustomDocumentProperties.Count +
                " Custom Properties" + "\r\n" + "\r\n" + ErrorText + " Fields show '" + FieldError + "'", Update_All_Fields, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information);
        }
            catch(Exception ex)
            {
                throw;
            }
        }




        private void wzd_InsertField(CommandBarButton commandBarButton, ref bool CancelDefault)
        {
            try
            {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (formInsertField == null || formInsertField.IsDisposed)
                formInsertField = new FormInsertField();
            Set_Parent_AND_ShowDialog(formInsertField, ActiveDocument);
            //formInsertField.ShowDialog();
            formInsertField.Dispose();
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                formInsertField = null;
                throw;
            }
        }

        private void wzd_BlankStep(CommandBarButton commandBarButton, ref bool CancelDefault)
        {
            try
            {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (formBlankStepMBR == null || formBlankStepMBR.IsDisposed)
                formBlankStepMBR = new FormBlankStepMBR();
            Set_Parent_AND_ShowDialog(formBlankStepMBR, ActiveDocument);
            //formBlankStepMBR.ShowDialog();
            formBlankStepMBR.Dispose();
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                formBlankStepMBR = null;
                throw;
            }

        }


        public bool MenuItemAvailable(string strSourceCode, bool pblnWithinTableOK = false)
        {
            const string MyProcName = "MenuItemAvailable";
            string strSource;
            string strMenuErrorMessage = "";
            bool blnActiveRequired = false;
            bool blnPendingRefused = false;
            string strDocumentPendingUpdate = "";
            bool blnProtectionRequired = false;
            bool blnAfterIntroEnd = false;
            bool blnCheckCursor = false;
            Range rngSelection = null;
            Paragraph par = null;
            string strStepNumber = "";
            int intListLevelNumber = 0;

            switch (strSourceCode)
            {
                case MENUTYPE_WIZARD:
                    strSource = "Wizard";
                    blnActiveRequired = true;
                    blnPendingRefused = true;
                    blnProtectionRequired = true;
                    blnAfterIntroEnd = true;
                    blnCheckCursor = true;
                    break;
                case MENUTYPE_UTILITY:
                    strSource = "Utility";
                    blnActiveRequired = true;
                    blnPendingRefused = true;
                    blnProtectionRequired = true;
                    blnAfterIntroEnd = false;
                    blnCheckCursor = false;
                    break;
                case MENUTYPE_WORDPERFECT:
                    strSource = "WP Conversion";
                    blnActiveRequired = true;
                    blnPendingRefused = true;
                    blnProtectionRequired = true;
                    blnAfterIntroEnd = false;
                    blnCheckCursor = false;
                    break;
                case MENUTYPE_FILENEW:
                    strSource = "FileNew";
                    blnActiveRequired = false;
                    blnPendingRefused = true;
                    blnProtectionRequired = false;
                    blnAfterIntroEnd = false;
                    blnCheckCursor = false;
                    break;
                case MENUTYPE_FILEOPEN:
                    strSource = "FileOpen";
                    blnActiveRequired = false;
                    blnPendingRefused = true;
                    blnProtectionRequired = false;
                    blnAfterIntroEnd = false;
                    blnCheckCursor = false;
                    break;
                case MENUTYPE_WIZARD_SPECCHAR:
                    strSource = "SpecialCharacter";
                    blnActiveRequired = true;
                    blnPendingRefused = true;
                    blnProtectionRequired = true;
                    blnAfterIntroEnd = false;
                    blnCheckCursor = false;
                    break;
                default:
                    strSource = "*UNKNOWN*";
                    CommonService.Error_Handler("0", $"Invalid Source '{strSource}'", "MenuValidation", MyProcName);
                    goto MenuItemUnAvailable;
                    break;
            }

            if (blnActiveRequired)
            {

                if (!_documentManager.ValidDocument(ActiveDocument.Name))
                {
                    strMenuErrorMessage = $"This function is NOT AVAILABLE without a valid {AppAcronym} Document.\n";
                    if (!ActiveDocument.Name.StartsWith("Document"))
                    {
                        strMenuErrorMessage += $"\n('Save' may be needed to identify '{ActiveDocument.Name}' as Valid {AppAcronym} Document)"; ;
                    }
                    goto MenuItemUnAvailable;
                }
            }

            if (blnPendingRefused)
            {
                strDocumentPendingUpdate = _documentManager.PendingUpdateAny(ActiveDocument.Name);
                if (!string.IsNullOrEmpty(strDocumentPendingUpdate))
                {
                    strMenuErrorMessage = $"PENDING UPDATE on Document:\n\n\t{strDocumentPendingUpdate}\n\n" +
                                       $"Document requires:\n\n\t'Save' to commit the changes, -or-\n\t'Close' to lose the changes.\n\n" +
                                       $"Action required before using any {AppAcronym} functions.";
                    goto MenuItemUnAvailable;
                }

                strDocumentPendingUpdate = DatabaseHelper.Transaction_Active();
                if (!string.IsNullOrEmpty(strDocumentPendingUpdate))
                {
                    strMenuErrorMessage = $"ACTIVE TRANSACTION for Document:\n\n\t{strDocumentPendingUpdate}\n\n" +
                                       $"Document requires:\n\n\t'Save' to commit the changes, -or-\n\t'Close' to lose the changes.\n\n" +
                                       $"Action required before using any {AppAcronym} functions.";
                    goto MenuItemUnAvailable;
                }
            }


            if (blnProtectionRequired)
            {
                if (ActiveDocument.ProtectionType == WdProtectionType.wdAllowOnlyFormFields)
                {
                    // Do nothing, allowed
                }
                else if (ActiveDocument.ProtectionType == WdProtectionType.wdNoProtection)
                {
                    var result = _documentManager.Document_SelectQPSObject(ActiveDocument.Name, null, true);
                    CommonService.DocumentProtection_On(ActiveDocument,
                        result.GetClsDocument.DocPasswordProtect); // Todo less params
                }
                else
                {
                    strMenuErrorMessage = $"Document:\n\n\t{ActiveDocument.Name}\n\n" +
                                       $"has unknown ProtectionType '{ActiveDocument.ProtectionType}'\n\n" +
                                       $"Document is not Available for Authoring.\n\n" +
                                       $"Please notify {AppAcronym} Support Staff.";
                    if (CommonService.gblnUserType_ADMIN)
                    {
                        System.Windows.Forms.MessageBox.Show($"ADMIN WARNING: {strMenuErrorMessage}\nPROCESSING CONTINUES", $"{AppAcronym} ADMIN OVERRIDE", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    }
                    else
                    {
                        goto MenuItemUnAvailable;
                    }
                }
            }

            if (blnAfterIntroEnd)
            {
                // Req? // Todo
            }

            if (blnCheckCursor)
            {
                if (ActiveDocument.Application.Selection.Information[WdInformation.wdWithInTable])
                {
                    if (!pblnWithinTableOK)
                    {
                        strMenuErrorMessage = "Cursor-Selection must NOT be within a TABLE.\n\nPlease re-position cursor and try again.";
                        goto MenuItemUnAvailable;
                    }
                }

                if (ActiveDocument.Application.Selection.Text.Length > 1)
                {
                    if (ActiveDocument.Application.Selection.Information[WdInformation.wdWithInTable] && pblnWithinTableOK)
                    {
                        // Do nothing
                    }
                    else
                    {
                        strMenuErrorMessage = "Cursor-Selection must be a single point.\n\nPlease re-select and try again.";
                        goto MenuItemUnAvailable;
                    }
                }
                rngSelection = ActiveDocument.Application.Selection.Range;

                if (rngSelection.End + 1 != rngSelection.Paragraphs[1].Range.End)
                {
                    if (!(pblnWithinTableOK && ActiveDocument.Application.Selection.Information[WdInformation.wdWithInTable]))
                    {
                        rngSelection.Paragraphs[1].Range.Select();
                        strMenuErrorMessage = "Cursor-Selection must be at the END of a PARAGRAPH.\n\nPlease re-position cursor and try again.";
                        goto MenuItemUnAvailable;
                    }
                }

                par = ActiveDocument.Application.Selection.Paragraphs[1];
                try
                {
                    par = par.Next();
                    while (string.IsNullOrWhiteSpace(par.Range.Text) || par.Range.Text.Length < 2)
                    {
                        par = par.Next();
                    }
                }
                catch
                {
                    // End of the paragraphs - handle it below
                }

                if (pblnWithinTableOK && ActiveDocument.Application.Selection.Information[WdInformation.wdWithInTable])
                {
                    // Do nothing
                }
                else if (par.get_Style().NameLocal == CommonService.gstrStyle_MBRSteps)
                {
                    strStepNumber = StepNumber_Find_Extraction(par.Range, false, intListLevelNumber);
                    if (intListLevelNumber != 1)
                    {
                        goto MenuItemUnAvailable;
                    }
                }
                else if (par.Range.Sections[1].Index != ActiveDocument.Application.Selection.Range.Sections[1].Index)
                {
                    // Do nothing
                }
                else if (string.Equals(par.Range.Text.TrimEnd('\r', '\n'), "END OF RECORD", StringComparison.OrdinalIgnoreCase))
                {
                    // Do nothing
                }
                else
                {
                    goto MenuItemUnAvailable;
                }
            }

            return true;

        MenuItemUnAvailable:
            if (!string.IsNullOrEmpty(strMenuErrorMessage))
            {
                System.Windows.Forms.MessageBox.Show($"ERROR: {strMenuErrorMessage}", $"{AppAcronym} Menu Validation - {strSource}", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            }
            return false;
        }

        public string StepNumber_Find_Extraction(Range prngCandidate, bool pblnReturnLevelOneOnly, int pintListLevelNumber = -1)
        {
            ListFormat lftListFormat = prngCandidate.ListFormat;
            string strStepNumber = lftListFormat.ListString;
            pintListLevelNumber = lftListFormat.ListLevelNumber;
            if (strStepNumber != "")
            {
                while (strStepNumber != "" && (strStepNumber.StartsWith("_") || strStepNumber.StartsWith(" ")))
                {
                    strStepNumber = strStepNumber.Substring(1);
                }

                while (strStepNumber != "" && strStepNumber.EndsWith("."))
                {
                    strStepNumber = strStepNumber.Substring(0, strStepNumber.Length - 1);
                }
                if (pblnReturnLevelOneOnly)
                {
                    if (lftListFormat.ListLevelNumber == 1)
                    {
                    }
                    else
                    {
                        strStepNumber = "";
                    }
                }
                else { }
            }
            return strStepNumber;
        }
        public void FileOpenDocumentName(string pstrDocumentName)
        {
            mstrFileOpenDocumentName = pstrDocumentName;
        }
        private void ShowError(string message, string source)
        {
            System.Windows.Forms.MessageBox.Show($"ERROR: {message}", $"{AppAcronym} Menu Validation - {source}", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
        }
        private void FileNew_MBR(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
            {
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (!MenuItemAvailable(MENUTYPE_FILENEW)) { return; }
            FormFileNewWizard formFileNewWizard = new FormFileNewWizard();
            Set_Parent_AND_ShowDialog(formFileNewWizard, ActiveDocument);
            //formFileNewWizard.ShowDialog();
            SetFocusToFile();
            }
            catch(Exception ex)
            {
                throw;
            }

        }

        private void MBRHelp_Click(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            FormMBRHelp formMBRHelp = new FormMBRHelp();
            Set_Parent_AND_ShowDialog(formMBRHelp, ActiveDocument);
            //formMBRHelp.ShowDialog();
        }

        public void SetFocusToFile(string documentName = "")
        {
            string targetDocumentName = string.IsNullOrEmpty(documentName) ? mstrFileOpenDocumentName : documentName;

            // Exit if no document name is specified
            if (string.IsNullOrEmpty(targetDocumentName))
            {
                return;
            }

            foreach (Window window in ActiveDocument.Application.Windows)
            {
                if (window.Document.Name == targetDocumentName)
                {
                    window.SetFocus();
                    // Switch to Print View if not already in Print View
                    if (window.View.Type != WdViewType.wdPrintView)
                    {
                        window.View.Type = WdViewType.wdPrintView;
                    }
                    // Exit once the document is found and focused
                    break;
                }
            }

            // Clear the document name after focus is set
            mstrFileOpenDocumentName = string.Empty;
        }

        private void AboutApp(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
            {
            FormAbout formAbout = new FormAbout();
            Set_Parent_AND_ShowDialog(formAbout, ActiveDocument);
            }
            catch(Exception ex)
            {
                throw;
            }
        }

        private void utl_PageSetup(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
        {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (!MenuItemAvailable(MENUTYPE_UTILITY)) { return; }

            if (formpageSetup == null || formpageSetup.IsDisposed)
                formpageSetup = new FormPageSetup(); // merge
            Set_Parent_AND_ShowDialog(formpageSetup, ActiveDocument);
            formpageSetup.Dispose();
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                formpageSetup = null;
                throw;
            }
        }
        private void utl_EditMaterials(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
        {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (!MenuItemAvailable(MENUTYPE_UTILITY)) { return; }
            if (formEditMaterials == null || formEditMaterials.IsDisposed)
            {
                formEditMaterials = new FormEditMaterials();
            }

            formEditMaterials.ShowDialog();
            formEditMaterials.Dispose();
            }
            catch(Exception ex)
            {
                formEditMaterials = null;
                throw;
            }
        }

        private void utl_EditModule(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
            {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (!MenuItemAvailable(MENUTYPE_UTILITY)) { return; }
            if (formEditModule == null || formEditModule.IsDisposed)
                formEditModule = new FormEditModule();
           // Set_Parent_AND_ShowDialog(formEditModule, ActiveDocument);
            formEditModule.ShowDialog();
            formEditModule.Dispose();
            // commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal; To suppress RPC
            }
            catch(Exception ex)
            {
                formEditModule = null;
                throw;
            }
        }

        private void utl_EditProperty(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
            {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            //commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (!MenuItemAvailable(MENUTYPE_UTILITY)) { return; }
            if (formEditProperty == null || formEditProperty.IsDisposed)
                formEditProperty = new FormEditProperty();
            formEditProperty.SetTargetDocumentName(ActiveDocument?.Name, ActiveDocument?.FullName);
            
            // Set_Parent_AND_ShowDialog(formEditProperty, ActiveDocument);

            formEditProperty.ShowDialog();
            formEditProperty.Dispose();
            //commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                formEditProperty = null;
                throw;
            }
        }

        private void utl_RefreshProperties(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
            {
                Word.Application App = commandBarButton.Application;
                Word.Document doc = App.ActiveDocument;
                if(doc == null)
                {
                    MessageBox.Show("No Active Document Found");
                    return;
                }
                ActiveDocument= doc;
                ActiveDocument.Application.System.Cursor = Word.WdCursorType.wdCursorWait;
                //commandBarButton.Application.System.Cursor = Word.WdCursorType.wdCursorWait;
                //ActiveDocument = commandBarButton.Application.ActiveDocument;
                GetActiveDocValues();
                if (ExecuteWordComWithRetry(() => ActiveDocument.ProtectionType != global::Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection))
                {
                    Document_UnProtect();
                    
                }              

                Setup_Document_Property();
               tabProperty  = CommonService.Table_Find(ExecuteWordComWithRetry(() => ActiveDocument), "Property List");
                if (tabProperty == null)
                {
                    MessageBox.Show("Property TABLE NOT FOUND - NO TABLE REPLACEMENT");
                }
                else
                {
                    Setup_DocumentProperty_Text(pblnReplaceTable: true);
                }

                if (ExecuteWordComWithRetry(() => ActiveDocument.ProtectionType != global::Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)) 
                {
                    Document_UnProtect();
                }
                if (ExecuteWordComWithRetry(() => ActiveDocument.TrackRevisions == true))
                {
                    ExecuteWordComWithRetry(() => ActiveDocument.TrackRevisions = false);
                }
                UpdateWordDocumentFields(); //220125
                //if (mblnDocTrackChgsReq == true)
                //{
                    //if (ActiveDocument.TrackRevisions == false)
                    //{
                    //    ActiveDocument.TrackRevisions = true;
                    //}
                //}
               
                SaveManager.Document_Save_FileAndDatabase(pdocCurrent: ExecuteWordComWithRetry(() => ActiveDocument), docpasswordProtect: mstrProtectionPassword, pblnRestoreCommandBar: false);
                
                CommonService.DocumentProtection_Assure(pdocProtectedDocument: ExecuteWordComWithRetry(() => ActiveDocument), pstrProtectionPassword: mstrProtectionPassword);
                 ExecuteWordComWithRetry(() => ActiveDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal);
                //ExecuteWordComWithRetry(() => commandBarButton.Application.System.Cursor = Word.WdCursorType.wdCursorNormal);
            }
            catch (Exception ex)
            {
                
                throw;
            }
        }

        public void GetActiveDocValues()
        {


            var result = _documentManager.Document_SelectAttributes(ExecuteWordComWithRetry(() => ActiveDocument.Name));
            mstrDocName = ActiveDocument.Name;
            mstrDocFullName = ActiveDocument.FullName;
            mDocID = result.DocumentID;
            mstrProtectionPassword = result.DocumentProtectionPassword;
            mintOpUnitID = result.DocumentOpUnitID;
            mintDocTypeID = result.DocumentDocTypeID;
            mstrOpUnitType = result.DocumentOpUnitType;

        }
        private void Setup_Document_Property()
        {
           
            int intListBoxPropertyCount;
            DocProperty[] adpyListBoxProperty;
            dynamic dpcDocProperties;
            int intListScan;
            
            try
            {
                dpcDocProperties = ExecuteWordComWithRetry(() => ActiveDocument.CustomDocumentProperties);
                dtDocProperty = objMBRAuthDAL.GetDocumentProperty(mDocID);

            dtDocProperty.Columns.Remove("CountOfMasters");
            dtDocProperty.Columns.Remove("CountOfModules");

                while (dpcDocProperties.Count > 0)
                {
                    dpcDocProperties.Item(1).Delete();
                }
                adpyListBoxProperty = new DocProperty[dtDocProperty.Rows.Count]; //COmp
                intListBoxPropertyCount = 0;

                for (intListScan = 0; intListScan < dtDocProperty.Rows.Count; intListScan++)
                {
                    Array_Add(parrWork: ref adpyListBoxProperty, pintArrayMax: ref intListBoxPropertyCount, pstrPropertyName: dtDocProperty.Rows[intListScan]["PropertyInstance"].ToString(), pstrPropertyValue: dtDocProperty.Rows[intListScan]["PropertyValueDisplay"].ToString());
                }


                foreach (DocProperty property in adpyListBoxProperty)
                    {
                    if (property.PropertyName.ToUpper() == "CHANGENUM")
                    {
                        CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, CommonService.cChangeNum_ValueDefault, ExecuteWordComWithRetry(() => ActiveDocument));
                        continue;
                    }
                    CommonService.DocumentProperty_PropertyAdd(property.PropertyName, false, property.PropertyValue, ExecuteWordComWithRetry(() => ActiveDocument));
                    }
                
                UpdateWordDocumentFields(); 
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
        public void UpdateWordDocumentFields()
        {
            

            foreach (Word.Section sec in ExecuteWordComWithRetry(() => ActiveDocument.Sections))
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

        private void Setup_DocumentProperty_Text(bool pblnReplaceTable)
        {
            
            int intTableRow = 0;
            int DocProperty_ListCount;
            int DocProperty_ListIndex;
            int TableRow_Count;
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
                    tabProperty = CommonService.Table_Find(ExecuteWordComWithRetry(() => ActiveDocument), "Property List");
                }

                if (pblnReplaceTable)
                {
                    Document_UnProtect();
                    ReApplyProtection = true;
                    CommonService.Table_Delete(ptblToDelete: tabProperty, pdocCurrent: ExecuteWordComWithRetry(() => ActiveDocument), pblnSelect: true);
                }
                else
                {
                    ReApplyProtection = false;
                }
                DocProperty_ListCount = dtDocProperty.Rows.Count;
                if (dtDocProperty.Rows.Count > 0)
                {
                    TableRow_Count = DocProperty_ListCount;
                }
                else
                {
                    TableRow_Count = 1;
                }
                
                CommonService.AutoText_Insert("PropertyList", ref intTableRow, TableRow_Count, tableAutoText: tabProperty, wordApp: ExecuteWordComWithRetry(() => ActiveDocument.Application));
                {
                    tabProperty = CommonService.Table_Find_PropertyMBR(ExecuteWordComWithRetry(() => ActiveDocument));
                    tabProperty.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                }
               
                for (DocProperty_ListIndex = 0; DocProperty_ListIndex <= DocProperty_ListCount - 1; DocProperty_ListIndex++)
                {
                    
                    PropertyName = dtDocProperty.Rows[DocProperty_ListIndex]["PropertyInstance"].ToString();//qdp100125
                    tabProperty.Cell(TableRow_Index, cintColPropertyName).Range.Text = PropertyName;
                    tabProperty.Cell(TableRow_Index, cintColPropertyValue).Range.Text = " ";
                    CommonService.Field_InsertProperty(where: tabProperty.Cell(TableRow_Index, cintColPropertyValue).Range, propertyName: PropertyName);
                    TableRow_Index = TableRow_Index + 1;
                }
               

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

        private void Array_Add(ref DocProperty[] parrWork, ref int pintArrayMax, string pstrPropertyName, string pstrPropertyValue)
        {
            
            if (pintArrayMax <= parrWork.Count())
            {
                parrWork[pintArrayMax].PropertyName = pstrPropertyName;
                parrWork[pintArrayMax].PropertyValue = pstrPropertyValue;
                pintArrayMax = pintArrayMax + 1;
            }

        }
        private void Document_Protect()
        {
           
            CommonService.DocumentProtection_On(pdocProtectedDocument: ActiveDocument, pstrProtectionPassword: mstrProtectionPassword);
        }
        private void Document_UnProtect()
        {

            CommonService.DocumentProtection_Off(pdocProtectedDocument: ActiveDocument, pstrProtectionPassword: mstrProtectionPassword);
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

        private void utl_RefreshModules(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
            {
                Word.Application App = commandBarButton.Application;
                Word.Document doc = App.ActiveDocument;
                bool blnProtect = false;
                if (doc == null)
                {
                    MessageBox.Show("No Active Document Found");
                    return;
                }
                ActiveDocument = doc;
                ActiveDocument.Application.System.Cursor = Word.WdCursorType.wdCursorWait;
                //commandBarButton.Application.System.Cursor = Word.WdCursorType.wdCursorWait;
                GetActiveDocValues();
                int DocID = mDocID;

                dtDocModule = objMBRAuthDAL.GetDocumentModule(DocID);
                //dtDocModule.Columns.RemoveAt(6);
                //dtDocModule.Columns.RemoveAt(7);
                //dtDocModule.Columns.RemoveAt(8);
                //dtDocModule.Columns.RemoveAt(9);

                if (ExecuteWordComWithRetry(() => ActiveDocument.ProtectionType != global::Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection))
                {
                    blnProtect = true;
                    Document_UnProtect();
                }
                else
                {
                    blnProtect = false;
                }
                tabModule = CommonService.Table_Find(ExecuteWordComWithRetry(() => ActiveDocument), "Module List");


                if (tabModule == null)
                {
                    MessageBox.Show("Existing Module-Table NOT FOUND - Skipping UPDATE", "utl_RefreshModules", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    int moduleTableStart = ExecuteWordComWithRetry(() => tabModule.Range.Start);
                    CommonService.Table_Delete(ptblToDelete: tabModule, pdocCurrent: ExecuteWordComWithRetry(() => ActiveDocument), pstrProtectionPassword: mstrProtectionPassword);
                    ExecuteWordComWithRetry(() => ActiveDocument.Range(moduleTableStart, moduleTableStart).Select());
                    Setup_DocumentModules();
                }

                if (blnProtect)
                {
                    Document_Protect();
                }
                SaveManager.Document_Save_FileAndDatabase(pdocCurrent: ExecuteWordComWithRetry(() => ActiveDocument), docpasswordProtect: mstrProtectionPassword, pblnRestoreCommandBar: false);
                
                CommonService.DocumentProtection_Assure(pdocProtectedDocument: ExecuteWordComWithRetry(() => ActiveDocument), pstrProtectionPassword: mstrProtectionPassword);
                
                ExecuteWordComWithRetry(() => ActiveDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal);
               
            }
            catch {
                throw;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void Setup_DocumentModules()
        {
            int intModule;
            int intTableRow = 0;
            
            int intModuleCount;
            if (dtDocModule.Rows.Count > 0)
            {
                intModuleCount = dtDocModule.Rows.Count;
            }
            else
            {
                intModuleCount = 1;
            }
            
            CommonService.AutoText_Insert("ModuleList", ref intTableRow, intModuleCount, tableAutoText: tabModule, wordApp: ActiveDocument.Application);
            {
                tabModule = CommonService.Table_Find_ModuleMBR(ActiveDocument);
                tabModule.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            }

            for (intModule = 0; intModule <= dtDocModule.Rows.Count - 1; intModule++)
            {
              
                tabModule.Cell(intTableRow, 1).Range.Text = dtDocModule.Rows[intModule]["ModuleBaseName"].ToString();
                if (dtDocModule.Rows[intModule]["ModuleUsage"].ToString() == "B")
                {
                    tabModule.Cell(intTableRow, 2).Range.Text = "Batch";
                }

                else if (dtDocModule.Rows[intModule]["ModuleUsage"].ToString() == "I")
                {
                    tabModule.Cell(intTableRow, 2).Range.Text = "Info";
                }
                string strText;
                strText = dtDocModule.Rows[intModule]["DocumentTitle"].ToString();
                if (dtDocModule.Rows[intModule]["ModuleTitle"].ToString() != "N/A")
                {
                    strText = strText + "\r\n" + "Module Title: " + dtDocModule.Rows[intModule]["ModuleTitle"].ToString();
                }

                if (dtDocModule.Rows[intModule]["ModuleTableRows"].ToString() != "N /A" & dtDocModule.Rows[intModule]["ModuleTableRows"].ToString() != "")
                    strText = strText + "\r\n" + "Module Table Rows: " + dtDocModule.Rows[intModule]["ModuleTableRows"].ToString();
                tabModule.Cell(intTableRow, 3).Range.Text = strText;
                tabModule.Cell(intTableRow, 4).Range.Text = dtDocModule.Rows[intModule]["Packet"].ToString();
                tabModule.Cell(intTableRow, 5).Range.Text = dtDocModule.Rows[intModule]["ModuleRepetitions"].ToString();
                intTableRow = intTableRow + 1;


            }
        }


        private void utl_RefreshMaterials(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
            {
                
                Word.Application App = commandBarButton.Application;
                Word.Document doc = App.ActiveDocument;
                //bool blnProtect = false;
                if (doc == null)
                {
                    MessageBox.Show("No Active Document Found");
                    return;
                }
                ActiveDocument = doc;
                ExecuteWordComWithRetry(() => ActiveDocument.Application.System.Cursor = Word.WdCursorType.wdCursorWait);
                GetActiveDocValues();
                Load_lstMaterialsWithAmounts();                
                Word.Selection selection = ExecuteWordComWithRetry(() => ActiveDocument.Application.Selection);
                mrngHold = selection.Range;
                Document_UnProtect();
                Setup_DocumentMaterial(true);
                CommonService.SwapExecute(ExecuteWordComWithRetry(() => ActiveDocument));
                mrngHold.Select();
                SaveManager.Document_Save_FileAndDatabase(ExecuteWordComWithRetry(() => ActiveDocument), docpasswordProtect: mstrProtectionPassword, pblnRestoreCommandBar: false);
                              
                ExecuteWordComWithRetry(() => ActiveDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal);
                

            }
            finally {
                Cursor.Current = Cursors.Default;
            }
        }
        public static IFormManager _formManager;
        private void Load_lstMaterialsWithAmounts()
        {
             _formManager =  new FormManager();

            List<DocMaterial_SelectDetails> docMaterials = _formManager.Fill_Document_Material(mDocID);
            dtDocMaterial = new System.Data.DataTable();
            dtDocMaterial.Columns.Add("MaterialDisplay");
            dtDocMaterial.Columns.Add("Amount");
            dtDocMaterial.Columns.Add("Units");
            dtDocMaterial.Columns.Add("Usage");
            dtDocMaterial.Columns.Add("MaterialDBID");
            dtDocMaterial.Columns.Add("MaterialItemCode");
            dtDocMaterial.Columns.Add("MaterialQS");
            dtDocMaterial.Columns.Add("MaterialTitle");
            dtDocMaterial.Columns.Add("MaterialCategory");
            dtDocMaterial.Columns.Add("MaterialImprint");
            dtDocMaterial.Columns.Add("MaterialPerUnitData");
            dtDocMaterial.Columns.Add("MaterialBOMText");
            dtDocMaterial.Columns.Add("MaterialUseCount");
            foreach (var material in docMaterials)
            {
                DataRow dataRow = dtDocMaterial.NewRow();

                dataRow["MaterialDisplay"] = material.MaterialDisplay;
                dataRow["MaterialTitle"] = material.MaterialTitle;
                dataRow["MaterialDBID"] = material.MaterialID;
                dataRow["MaterialItemCode"] = material.ItemCode;

                if (System.Convert.ToDouble(material.MaterialAmountDisplay) != 0)
                {
                    dataRow["Amount"] = material.MaterialAmountDisplay;
                    dataRow["Units"] = material.MaterialUOM;

                }
                else
                {
                    dataRow["Amount"] = "";
                    dataRow["Units"] = "";
                }
                if (material.QuantitySufficient == true)
                {
                    dataRow["MaterialQS"] = CommonService.qdlMaterialAmount_qs_No_Amount;
                }
                else
                {
                    dataRow["MaterialQS"] = "";

                }
                dataRow["Usage"] = material.MaterialUsage;
                dataRow["MaterialCategory"] = material.MaterialCategory;
                dataRow["MaterialImprint"] = material.MaterialImprint;
                dataRow["MaterialPerUnitData"] = material.MaterialPerUnitData;
                dataRow["MaterialBOMText"] = material.MaterialBOMText;
                dataRow["MaterialUseCount"] = 0;

                dtDocMaterial.Rows.Add(dataRow);
            }

        }

        private void Setup_DocumentMaterial(bool pblnReplaceTable)
        {
            Word.Table tblMaterial = null;
            int intListScan;
            int intRowScan = 0;
            string MaterialUsage;
            int MaterialUsage_Index;
            bool blnAllowAmount;
            Word.Range rngWork;
            string MaterialImprint;
            string MaterialPerUnitData;
            string MaterialBOMText;
            if (pblnReplaceTable)
            {
                tblMaterial = CommonService.Table_Find_MaterialMBR(ExecuteWordComWithRetry(() => ActiveDocument));

                if (tblMaterial is null)
                {
                    MessageBox.Show("COULD NOT FIND MATERIAL TABLE - TERMINATING", "Setup_DocumentMaterial", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                CommonService.Table_Delete(tblMaterial, ExecuteWordComWithRetry(() => ActiveDocument), pstrProtectionPassword: mstrProtectionPassword);
            }
            ExecuteWordComWithRetry(() => ActiveDocument.Application.Selection.Range.Select());
            ExecuteWordComWithRetry(() => ActiveDocument.Application.Selection.Font.Size = 4);

            CommonService.AutoText_Insert("MaterialsMBR", ref intRowScan, dtDocMaterial.Rows.Count, tableAutoText: tblMaterial, wordApp: ExecuteWordComWithRetry(() => ActiveDocument.Application));
            tblMaterial = CommonService.Table_Find_MaterialMBR(ExecuteWordComWithRetry(() => ActiveDocument));
            tblMaterial.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            var usageAllowAmountMap = new FormManager()
                   .Fill_MaterialUsage()
                   .Where(x => x.MaterialUsage != null)
                   .GroupBy(x => x.MaterialUsage)
                   .ToDictionary(g => g.Key, g => g.First().MaterialUsageAllowAmount);
            for (intListScan = 0; intListScan < dtDocMaterial.Rows.Count; intListScan++)
            {
                //if (dtDocMaterial.Rows[intListScan].IsNewRow) continue; // added :there are no empty rows (optional check)

                int tableMaterialRowIndex = intRowScan + intListScan;

                tblMaterial.Cell(tableMaterialRowIndex, 1).Range.Text = dtDocMaterial.Rows[intListScan]["MaterialItemCode"].ToString();
                MaterialUsage = dtDocMaterial.Rows[intListScan]["Usage"].ToString();

                if (MaterialUsage.ToUpper() != "STANDARD")
                {
                    rngWork = tblMaterial.Cell(tableMaterialRowIndex, 1).Range;
                    rngWork.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    rngWork.Text = MaterialUsage + " ";
                    rngWork.Font.Italic = 1;
                }

                MaterialPerUnitData = dtDocMaterial.Rows[intListScan]["MaterialPerUnitData"].ToString();
                MaterialImprint = dtDocMaterial.Rows[intListScan]["MaterialImprint"].ToString();
                MaterialBOMText = dtDocMaterial.Rows[intListScan]["MaterialBOMText"].ToString();

                tblMaterial.Cell(tableMaterialRowIndex, 2).Range.Text = dtDocMaterial.Rows[intListScan]["MaterialTitle"].ToString();

                if (!string.IsNullOrEmpty(MaterialPerUnitData))
                {
                    tblMaterial.Cell(tableMaterialRowIndex, 2).Range.Text += "Per Unit Data: " + MaterialPerUnitData;
                }

                if (!string.IsNullOrEmpty(MaterialImprint))
                {
                    tblMaterial.Cell(tableMaterialRowIndex, 2).Range.Text += "Imprint: " + MaterialImprint;
                }

                if (!string.IsNullOrEmpty(MaterialBOMText))
                {
                    tblMaterial.Cell(tableMaterialRowIndex, 2).Range.Text += "B.O.M. Text: " + MaterialBOMText;
                }

                MaterialUsage = dtDocMaterial.Rows[intListScan]["Usage"].ToString();

                // Replaced cboMaterialUsage lookup
                blnAllowAmount = usageAllowAmountMap.TryGetValue(MaterialUsage, out var allowAmount) && allowAmount;

                //blnAllowAmount = materialUsages.FirstOrDefault(x => x.MaterialUsage == dtDocMaterial.Rows[intListScan]["Usage"].ToString())?.MaterialUsageAllowAmount ?? false;
               // blnAllowAmount = cboMaterialUsage.Items.OfType<MaterialUsageDetails>().FirstOrDefault(x => x.MaterialUsage == MaterialUsage).MaterialUsageAllowAmount; // Fix for user-handle error: Clicking Save instead of Update 

                if (blnAllowAmount)
                {
                    if (dtDocMaterial.Rows[intListScan]["MaterialQS"].ToString() == CommonService.qdlMaterialAmount_qs_No_Amount)
                    {
                        if (dtDocMaterial.Rows[intListScan]["Amount"].ToString() == "")
                        {
                            tblMaterial.Cell(tableMaterialRowIndex, 3).Range.Text = CommonService.qdlMaterialAmount_qs_No_Amount;
                        }
                        else
                        {
                            tblMaterial.Cell(tableMaterialRowIndex, 3).Range.Text = CommonService.qdlMaterialAmount_qs_With_Amount + " " + dtDocMaterial.Rows[intListScan]["Amount"].ToString() + " " + dtDocMaterial.Rows[intListScan]["Units"].ToString();
                        }
                    }
                    else
                    {
                        tblMaterial.Cell(tableMaterialRowIndex, 3).Range.Text = dtDocMaterial.Rows[intListScan]["Amount"].ToString() + " " + dtDocMaterial.Rows[intListScan]["Units"].ToString();
                    }
                }
            }

        }

        private void FileOpen_MBR(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
            {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            //Application app = commandBarButton.Application;
            //app.WindowActivate += new ApplicationEvents4_WindowActivateEventHandler(App_WindowActivate);
            if (!MenuItemAvailable(MENUTYPE_FILEOPEN)) { return; }
            mstrFileOpenDocumentName = "";

            if (formFileOpenWizard == null || formFileOpenWizard.IsDisposed)
                formFileOpenWizard = new FormFileOpenWizard(); // merge
            Set_Parent_AND_ShowDialog(formFileOpenWizard, ActiveDocument);
            //formFileOpenWizard.ShowDialog();
            formFileOpenWizard.Dispose();
            SetFocusToFile();
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                formFileOpenWizard = null;
                throw;
            }
            //GC.Collect();
        }

        
        private void utl_DocSaveHistory(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
            {
            commandBarButton.Application.StatusBar = "Form Loading...";
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (!MenuItemAvailable(MENUTYPE_UTILITY)) { return; }
            if (formhistory == null || formhistory.IsDisposed)
                formhistory = new FormHistory();
            Set_Parent_AND_ShowDialog(formhistory, ActiveDocument);
            //formhistory.ShowDialog();
            formhistory.Dispose();
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                formhistory = null;
                throw;
            }
            //GC.Collect();
        }

        private void utl_ManageSections(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
        {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (!MenuItemAvailable(MENUTYPE_UTILITY)) { return; }
            if (sections == null || sections.IsDisposed)
                sections = new FormSections();
            Set_Parent_AND_ShowDialog(sections, ActiveDocument);
            //sections.ShowDialog(); // merge
            sections.Dispose();
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                sections = null;
                throw;
            }
        }

        private void utl_VerifyDocument(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            try
        {
            commandBarButton.Application.StatusBar = "Form Loading..."; // merge
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorWait;
            ActiveDocument = commandBarButton.Application.ActiveDocument;
            if (!MenuItemAvailable(MENUTYPE_UTILITY)) { return; }
            if (formverifyWizard == null || formverifyWizard.IsDisposed)
                formverifyWizard = new FormVerifyWizard();
            Set_Parent_AND_ShowDialog(formverifyWizard, ActiveDocument);
            //formverifyWizard.ShowDialog();
            formverifyWizard.Dispose();
            commandBarButton.Application.System.Cursor = WdCursorType.wdCursorNormal;
            }
            catch(Exception ex)
            {
                formverifyWizard = null;
                throw;
            }
            //GC.Collect();
        }

       
        private void Set_Parent_AND_ShowDialog(Form form, Document document)
        {
            //Get the active Word application
            if (document != null)
            {
                // Get the active window
                Window activeWindow = document.ActiveWindow;

                if (activeWindow != null)
                {
                    IntPtr parentHandle = new IntPtr(activeWindow.Hwnd);
                    RECT rect;
                    GetWindowRect(parentHandle, out rect);

                    // Calculate the center position
                    int centerX = rect.Left + (rect.Right - rect.Left - form.Width) / 2;
                    int centerY = rect.Top + (rect.Bottom - rect.Top - form.Height) / 2;

                    // Set the form's location
                    form.StartPosition = FormStartPosition.Manual;
                    form.Location = new System.Drawing.Point(centerX, centerY);
                    //Set the owner of the form to the Word window
                    if (parentHandle != IntPtr.Zero)
                    {
                        form.ShowDialog(new WindowWrapper(parentHandle)); // Show the form with Word as the owner
                    }
                    else
                    {
                        MessageBox.Show("Could not find the Word window.");
                    }
                }
            }
        }

        public class WindowWrapper : IWin32Window // Newcode
        {
            public IntPtr _hwnd;

            public WindowWrapper(IntPtr hwnd)
            {
                _hwnd = hwnd;
            }

            public IntPtr Handle => _hwnd;
        }


        private void ThisDocument_Shutdown(object sender, System.EventArgs e)  // NewCode
        {
            try
            {
                if (ActiveDocument.FullName != this.FullName)
                {
                    CommandBar_DeleteAll();
                    Marshal.ReleaseComObject(ActiveDocument.Application);
                }
            }
            catch(Exception ex)
            {
                // ignore the exception raised by closing the Template when there is no active business doc 
            }
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);

            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
