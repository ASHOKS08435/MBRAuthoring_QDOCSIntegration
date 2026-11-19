using MBRAuthoring.Business;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MBRAuthoring.Presentation
{
    public partial class FormProtection : FormBase
    {

        public Word.Document currentActiveDocument;
        Word.Template tptActive;

        public IDocumentManager _documentManager = null;// ServiceLocator.GetDocumentManager();

        public string mstrDocName;
        string mstrProtectionPassword;
        bool mblnInternalClick;

        const string MyCompName = "frmProtection";

        const string cstrChgDisp_Final = "Final";
        const string cstrChgDisp_Final_Markup = "Final Showing Markup";
        const string cstrChgDisp_Original = "Original";
        const string cstrChgDisp_Original_Markup = "Original Showing Markup";


        public FormProtection()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();
        }

        private void FormProtection_Load(object sender, EventArgs e)
        {
            const string myMethodName = "FormProtection_Load";
            Cursor = Cursors.WaitCursor;

            currentActiveDocument = Globals.ThisDocument.ActiveDocument;//currentDocument;
            mstrDocName = currentActiveDocument.Name;

            var selectAttribute = _documentManager.Document_SelectAttributes(mstrDocName);//, mDocID, mblnCleaningRecord, mblnMasterRecord, mstrOpUnit, mintOpUnitID, mstrOpUnitType, mstrDocType, mstrDocTypeAbb, mintDocTypeID, mstrProtectionPassword, mblnDocTrackChgsReq)

            mstrProtectionPassword = selectAttribute.DocumentProtectionPassword;
            cboChangesDisplay.Items.Add(cstrChgDisp_Final_Markup);
            cboChangesDisplay.Items.Add(cstrChgDisp_Final);
            cboChangesDisplay.Items.Add(cstrChgDisp_Original_Markup);
            cboChangesDisplay.Items.Add(cstrChgDisp_Original);

            //hide 2 checkboxes : In existing app also these 2 checkboxes are not in use.
            chkTrackHighlightPrint.Visible = false;
            chkTrackHighlightScreen.Visible = false;
            lblStatus__LOAD();
            Cursor = Cursors.Default;
            currentActiveDocument.Application.StatusBar = "";
        }


        private void FormProtection_Activated(object sender, EventArgs e)
        {
            string strTemplateName;
            bool blnReview;
            tptActive = currentActiveDocument.get_AttachedTemplate();
            strTemplateName = tptActive.Name.ToUpper();

            int position = strTemplateName.IndexOf("REVIEW");
            if (position > 1)
            {
                blnReview = true;
            }
            else
            {
                blnReview = false;
            }


            if (CommonService.gblnUserType_ADMIN)
            {
                this.Text = ModChangeHistory.gstrAppAcronym + " ADMIN: Document Protection and Track Changes";
            }
            else
            {
                if (blnReview && CommonService.gblnUserType_REVIEWER)
                {
                    this.Text = ModChangeHistory.gstrAppAcronym + " REVIEW: Document Protection";
                    chkTrackChanges.Visible = false;
                    cmdPasswordDisplay.Visible = false;

                }
                else
                {
                    this.Text = ModChangeHistory.gstrAppAcronym + " Author: Track Changes";
                    chkProtectDocument.Visible = false;
                    chkTrackChanges.Enabled = true;
                    cmdPasswordDisplay.Visible = false;
                }

            }
        }

        private void cmdClose_Click(object sender, EventArgs e)
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }

        private void cmdPasswordDisplay_Click(object sender, EventArgs e)
        {

            string prompt = "Password for " + mstrDocName;
            string MyCompName = "frmProtection";

            InputBox.ShowDialog(prompt, MyCompName, mstrProtectionPassword, mstrDocName);
            cmdClose.Focus();

        }

        private void chkProtectDocument_CheckedChanged(object sender, EventArgs e)
        {
            if (mblnInternalClick)
                return;

            if (chkProtectDocument.Checked)
            {
                CommonService.DocumentProtection_On(currentActiveDocument, mstrProtectionPassword);
            }
            else
            {
                CommonService.DocumentProtection_Off(currentActiveDocument, mstrProtectionPassword, mstrDocName);
            }

            lblStatus__LOAD();
        }

        private void chkTrackChanges_CheckedChanged(object sender, EventArgs e)
        {

            bool blnRestoreProtection;
            if (mblnInternalClick)
            {
                return;
            }
            if (currentActiveDocument.ProtectionType != Word.WdProtectionType.wdNoProtection)
            {
                CommonService.DocumentProtection_Off(currentActiveDocument, mstrProtectionPassword, mstrDocName);
                blnRestoreProtection = true;

            }
            else
            {
                blnRestoreProtection = false;
            }
            currentActiveDocument.TrackRevisions = chkTrackChanges.Checked;
            if (blnRestoreProtection)
            {
                CommonService.DocumentProtection_On(currentActiveDocument, mstrProtectionPassword);

            }
            lblStatus__LOAD();
        }

        private void cboChangesDisplay_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                if (mblnInternalClick)
                {
                    return;
                }
                var view = currentActiveDocument.ActiveWindow.View;

                switch (cboChangesDisplay.Text)
                {
                    case cstrChgDisp_Final_Markup:
                        view.RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal;
                        view.ShowRevisionsAndComments = true;
                        break;
                    case cstrChgDisp_Final:
                        view.RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal;
                        view.ShowRevisionsAndComments = false;
                        break;
                    case cstrChgDisp_Original_Markup:
                        view.RevisionsView = Word.WdRevisionsView.wdRevisionsViewOriginal;
                        view.ShowRevisionsAndComments = true;
                        break;
                    case cstrChgDisp_Original:
                        view.RevisionsView = Word.WdRevisionsView.wdRevisionsViewOriginal;
                        view.ShowRevisionsAndComments = false;
                        break;
                    default:
                        MessageBox.Show("Unexpected value '" + cboChangesDisplay.Text + "'");
                        break;
                }
                lblStatus__LOAD();
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show("This method or property is not available because the document is a protected document.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //Added
            }
        }

        private void lblStatus__LOAD()
        {

            string strStatusProtection = "";
            string strStatusTrackChanges = "";
            string strStatusDisplayChanges = "";
            bool blnRestoreProtection = false;
            //int intProtectionType;
            int intSectionIndex;
            string strSectionList;
            string strReviewing;
            mblnInternalClick = true;

            Word.WdProtectionType intProtectionType = currentActiveDocument.ProtectionType;

            var view = currentActiveDocument.ActiveWindow.View;
            switch (intProtectionType)
            {
                case Word.WdProtectionType.wdNoProtection:
                    strStatusProtection = "Document Protection is OFF";
                    blnRestoreProtection = false;
                    chkProtectDocument.Checked = false;
                    break;

                case Word.WdProtectionType.wdAllowOnlyFormFields:
                    strStatusProtection = "Document Protection is ON (Form Fields).";
                    blnRestoreProtection = true;
                    chkProtectDocument.Checked = true;
                    break;

                case Word.WdProtectionType.wdAllowOnlyRevisions:
                    strStatusProtection = "Document Protection is 'Only Revisions' - NOT " + ModChangeHistory.gstrAppAcronym + " PROTECTION!";
                    blnRestoreProtection = true;
                    chkProtectDocument.Checked = true;
                    break;

                case Word.WdProtectionType.wdAllowOnlyComments:
                    strStatusProtection = "Document Protection is 'Only Comments' - NOT " + ModChangeHistory.gstrAppAcronym + " PROTECTION!";
                    blnRestoreProtection = true;
                    chkProtectDocument.Checked = true;
                    break;

                default:
                    lblStatus.Text = "Unrecognized Document ProtectionType=" + currentActiveDocument.ProtectionType;
                    return;
            }

            if (blnRestoreProtection)
            {
                CommonService.DocumentProtection_Off(currentActiveDocument, mstrProtectionPassword);

            }

            chkTrackChanges.Checked = currentActiveDocument.TrackRevisions;

            if (view.RevisionsView == Word.WdRevisionsView.wdRevisionsViewOriginal)
            {
                if (view.ShowRevisionsAndComments)
                {
                    strReviewing = cstrChgDisp_Original_Markup;
                }
                else
                {
                    strReviewing = cstrChgDisp_Original;
                }
            }
            else if (view.RevisionsView == Word.WdRevisionsView.wdRevisionsViewFinal)
            {
                if (view.ShowRevisionsAndComments)
                {
                    strReviewing = cstrChgDisp_Final_Markup;

                }
                else
                {
                    strReviewing = cstrChgDisp_Final;
                }
            }
            else
            {
                if (view.ShowRevisionsAndComments)
                {
                    strReviewing = view.RevisionsView + " Showing Markup";

                }
                else
                {
                    strReviewing = view.RevisionsView.ToString();
                }
            }
            strStatusDisplayChanges = "Display for Review: " + strReviewing;
            cboChangesDisplay.Text = strReviewing;
            if (currentActiveDocument.TrackRevisions)
            {
                strStatusTrackChanges = "TrackChanges is ON.";
            }
            else
            {
                strStatusTrackChanges = "TrackChanges is OFF.";

            }
            if (blnRestoreProtection)
            {
                CommonService.DocumentProtection_On(currentActiveDocument, mstrProtectionPassword);
            }
            strSectionList = "";
            for (intSectionIndex = 1; intSectionIndex <= currentActiveDocument.Sections.Count; intSectionIndex++)
            {
                Word.Section section = currentActiveDocument.Sections[intSectionIndex];
                if (section.ProtectedForForms)
                {
                    strSectionList = strSectionList + intSectionIndex + ",";
                }

            }
            if (strSectionList == "")
            {
                strSectionList = "NONE";

            }
            else
            {
                strSectionList = strSectionList.Substring(0, strSectionList.Length - 1);
            }
            strStatusProtection = strStatusProtection + ";    Section(s)=" + strSectionList;
            lblStatus.Text = "For Document " + mstrDocName + "\r\n" + strStatusProtection + "\r\n" + strStatusTrackChanges + "\r\n" + strStatusDisplayChanges;
            cmdClose.Focus();
            mblnInternalClick = false;

        }

    }

}
