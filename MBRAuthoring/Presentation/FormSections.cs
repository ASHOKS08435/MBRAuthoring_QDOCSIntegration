using MBRAuthoring.Business;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MBRAuthoring.Presentation
{
    public partial class FormSections : FormBase
    {
        string mstrProtectionPassword;

        public IDocumentManager _documentManager = null;// ServiceLocator.GetDocumentManager();
                                                        // public ICommandBarManager _commandBarManager;
        private Word.Document currentActiveDocument;

        Word.Section sec;
        Word.Table table;
        public FormSections()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();
        }

        private void FormSections_Load(object sender, EventArgs e)
        {
            const string MyProcName = "UserForm_Initialize";
            Cursor = Cursors.WaitCursor;

            currentActiveDocument = Globals.ThisDocument.ActiveDocument;//currentDocument;
            var result = _documentManager.Document_SelectAttributes(currentActiveDocument.Name);
            mstrProtectionPassword = result.DocumentProtectionPassword;

            grdSections.Columns.Add("SectionIndex", "Section");
            grdSections.Columns[0].Width = 84;

            grdSections.Columns.Add("SectionChars", "End-Page");
            grdSections.Columns[1].Width = 84;

            grdSections.Columns.Add("Protected", "Protect");
            grdSections.Columns[2].Width = 84;

            grdSections.Columns.Add("HeaderLink", "Hdr-LinkPrev");
            grdSections.Columns[3].Width = 84;

            grdSections.Columns.Add("FooterLink", "Ftr-LinkPrev");
            grdSections.Columns[4].Width = 84;

            grdSections.Columns.Add("PageRestart", "RestartPage");
            grdSections.Columns[5].Width = 84;

            grdSections.Columns.Add("PageStartAt", "PageNmbr");
            grdSections.Columns[6].Width = 84;

            grdSections.Rows.Clear();
            Cursor = Cursors.Default;
            currentActiveDocument.Application.StatusBar = "";
        }

        private void FormSections_Activate(object sender, EventArgs e)
        {
            CommonService.DocumentProtection_Off(currentActiveDocument, mstrProtectionPassword);
            grdSections.Rows.Clear();
            try
            {
                Cursor = Cursors.WaitCursor;
                ListData();
                txtRestartPageNumbering.Text = grdSections.Rows[grdSections.CurrentCell.RowIndex].Cells["PageStartAt"].Value.ToString();
                cmdApply__AVAILABLE();
                Cursor = Cursors.Default;
            }
            catch
            {
                Cursor = Cursors.Default;
                CommonService.DocumentProtection_On(currentActiveDocument, mstrProtectionPassword);
            }
        }


        private void cmdApply_Click(object sender, EventArgs e) //merge
        {
            Cursor = Cursors.WaitCursor;
            int intListIndex;
            intListIndex = grdSections.SelectedRows[0].Index + 1;
            try
            {
                sec = currentActiveDocument.Sections[intListIndex];
                CommonService.DocumentProtection_Off(currentActiveDocument, mstrProtectionPassword);
                if ((sec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious) != chkHeaderLinkPrevious.Checked)
                {
                    sec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = chkHeaderLinkPrevious.Checked;
                }

                if (sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious != chkFooterLinkPrevious.Checked)
                {
                    sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = chkFooterLinkPrevious.Checked;

                    if (sec.Index == 2 && chkFooterLinkPrevious.Checked == false)
                    {
                        table = sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Tables[1];
                        CommonService.InsertFooterMBR_PageNumber(sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Tables[1], false);
                        CommonService.ViewPage_Main(currentActiveDocument.Application);
                    }
                }

                if (sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection != chkRestartPageNumbering.Checked)
                {
                    sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = chkRestartPageNumbering.Checked;
                }

                if (sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber != int.Parse(txtRestartPageNumbering.Text))
                {
                    sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = int.Parse(txtRestartPageNumbering.Text);
                }
                CommonService.DocumentProtection_On(currentActiveDocument, mstrProtectionPassword);
                this.Close();
                Cursor = Cursors.Default;
            }
            finally
            {
                CommonService.CleanUpBackgroundWordProcesses();
            }

        }

        private void cmdClose_Click(object sender, EventArgs e)
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }

        private void grdSections_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int intListIndex;
            intListIndex = grdSections.SelectedRows[0].Index + 1;

            if (grdSections.SelectedRows.Count > 0)
            {
                chkHeaderLinkPrevious.Visible = true;
                chkFooterLinkPrevious.Visible = true;
                chkRestartPageNumbering.Visible = true;
                txtRestartPageNumbering.Visible = true;

                //int rowIndex = grdSections.CurrentCell.RowIndex;
                chkHeaderLinkPrevious.Checked = (bool)grdSections.Rows[grdSections.CurrentCell.RowIndex].Cells["HeaderLink"].Value;
                chkFooterLinkPrevious.Checked = (bool)grdSections.Rows[grdSections.CurrentCell.RowIndex].Cells["FooterLink"].Value;
                chkRestartPageNumbering.Checked = (bool)grdSections.Rows[grdSections.CurrentCell.RowIndex].Cells["PageRestart"].Value;
                txtRestartPageNumbering.Text = grdSections.Rows[grdSections.CurrentCell.RowIndex].Cells["PageStartAt"].Value.ToString();
            }
            cmdApply__AVAILABLE();
        }

        private void ListData()
        {
            int sectionIndex = 1;

            // Ensure pagination is complete
            //currentActiveDocument.Application.Options.Pagination = true;
            currentActiveDocument.Repaginate();

            foreach (Word.Section sec in currentActiveDocument.Sections)
            {
                
                int sectionEndPage = sec.Range.Information[Word.WdInformation.wdActiveEndPageNumber];
                string strSectionChars = sec.Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber] + "/" + sectionEndPage;

                bool sectionProtected = sec.ProtectedForForms;
                bool headerLinedToPrevious = sec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious;
                bool footerLinedToPrevious = sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious;
                bool restartPageNumbers = sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection;
                int pageStartAt = sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber;

                grdSections.Rows.Add(sectionIndex, strSectionChars, sectionProtected, headerLinedToPrevious, footerLinedToPrevious, restartPageNumbers, pageStartAt);
                sectionIndex++;
            }
        }

        private void chkHeaderLinkPrevious_CheckedChanged(object sender, EventArgs e)
        {
            cmdApply__AVAILABLE();
        }

        private void chkFooterLinkPrevious_CheckedChanged(object sender, EventArgs e)
        {
            cmdApply__AVAILABLE();
        }

        private void chkRestartPageNumbering_CheckedChanged(object sender, EventArgs e)
        {
            cmdApply__AVAILABLE();
            // The "Start at Page Number" TextBox is available only if the "Restart Numbering" checkbox is on
            txtRestartPageNumbering.Enabled = chkRestartPageNumbering.Checked;
        }

        private void cmdApply__AVAILABLE()
        {

            int intListIndex;
            bool blnHdrLink;
            bool blnFtrLink;
            bool blnFtrRest;
            string strFtrPNbr;
            // After any change to a data-entry control, check if there is a net change:
            // If there is NO change, then disable the "APPLY" button
            // If there IS a change,  then  enable the "APPLY" button

            intListIndex = grdSections.SelectedRows[0].Index + 1;
            if (intListIndex > 0)
            {
                blnHdrLink = (bool)grdSections.Rows[grdSections.CurrentCell.RowIndex].Cells["HeaderLink"].Value;
                blnFtrLink = (bool)grdSections.Rows[grdSections.CurrentCell.RowIndex].Cells["FooterLink"].Value;
                blnFtrRest = (bool)grdSections.Rows[grdSections.CurrentCell.RowIndex].Cells["PageRestart"].Value;
                strFtrPNbr = grdSections.Rows[grdSections.CurrentCell.RowIndex].Cells["PageStartAt"].Value.ToString();
                if (chkHeaderLinkPrevious.Checked == blnHdrLink && chkFooterLinkPrevious.Checked == blnFtrLink && chkRestartPageNumbering.Checked == blnFtrRest && txtRestartPageNumbering.Text == strFtrPNbr)
                {
                    cmdApply.Enabled = false;
                }
                else
                {
                    cmdApply.Enabled = true;
                }

            }

            else
            {
                cmdApply.Enabled = false;
            }
            if (cmdApply.Enabled)
            {
                cmdApply.Text = "Apply to Section " + intListIndex + " and Close";
            }
            else
            {
                cmdApply.Text = "Apply";

            }

        }
    }
}
