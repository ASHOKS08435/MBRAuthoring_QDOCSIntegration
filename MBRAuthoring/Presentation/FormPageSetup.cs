using MBRAuthoring.Business;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MBRAuthoring.Presentation
{
    public partial class FormPageSetup : FormBase
    {
        protected override Panel progressBarPanel => panel1;
        public IDocumentManager _documentManager = null;// ServiceLocator.GetDocumentManager();
        private Word.Document currentActiveDocument;

        private Word.WdOrientation mintDocumentNewOrientation;
        Word.Selection selection;

        const string MyCompName = "frmPageSetup";

        public string mstrDocName;
        string mstrProtectionPassword;
        int mintDocumentSectionNumber;

        public FormPageSetup()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();
        }

        private void FormPageSetup_Load(object sender, EventArgs e)
        {
            var a = currentDocName;
            Cursor = Cursors.WaitCursor;

            try
            {
                //Todo
                //Form_Accelerators(this);
                //Form_TabIndexing(this);

                tabControl1.TabPages.Remove(tabPage2);
                tabControl1.TabPages.Remove(tabPage3);
                InitializeTabs(tabControl1, cmdBack, cmdNext, tabPage1);
                ShowTab(0);

                SetToolTip(cmdCancel, "Terminate Function with NO Output");  // NewCode
                SetToolTip(cmdFinish, "Verify Input, Generate Result, then Terminate Function");

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

                //_commandBarManager = null;// ServiceLocator.GetCommandBarManager();
                currentActiveDocument = Globals.ThisDocument.ActiveDocument;
                mstrDocName = currentActiveDocument.Name;

                var selectAttribute = _documentManager.Document_SelectAttributes(mstrDocName);
                mstrProtectionPassword = selectAttribute.DocumentProtectionPassword;

                // tabControl1.st = fmTabStyleNone; //Not req (In VBA - this line hides the tabs in a MultiPage control.

                selection = currentActiveDocument.Application.Selection;
                mintDocumentSectionNumber = selection.Sections[1].Index;
                lblWizardHeading.Text = "Enter/Verify options for Section " + mintDocumentSectionNumber + ":";
                Cursor = Cursors.Default;
                currentActiveDocument.Application.StatusBar = "";
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "FormPageSetup_Load");
            }

        }
        private void frmProtection_Activated(object sender, EventArgs e)
        {
            if (selection.Sections.Count > 1)
            {
                MessageBox.Show("Cannot run this function with  " + selection.Sections.Count + "SECTIONS Selected", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.Close();
                CommonService.CleanUpBackgroundWordProcesses();
                return;
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
                Cursor = Cursors.WaitCursor;
                if (!ValidEntries())
                {
                    return;
                }

                currentActiveDocument.Application.System.Cursor = Word.WdCursorType.wdCursorWait;
                this.Left = Application.OpenForms[0].Width - tabControl1.Left;
                Setup_Document();
                Cursor = Cursors.Default;
                this.Close();
                currentActiveDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;

            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                currentActiveDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;

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
            try
            {
                CommonService.DocumentProtection_Off(currentActiveDocument, mstrProtectionPassword);
                selection.PageSetup.Orientation = mintDocumentNewOrientation;

            }
            catch (Exception ex)
            {
                CommonService.DocumentProtection_On(currentActiveDocument, mstrProtectionPassword);
            }

        }

        private bool ValidEntries()
        {

            if (optOrientLandscape.Checked)
            {
                mintDocumentNewOrientation = Word.WdOrientation.wdOrientLandscape;
            }
            else if (optOrientPortrait.Checked)
            {
                mintDocumentNewOrientation = Word.WdOrientation.wdOrientPortrait;
            }
            else
            {
                MessageBox.Show("NO ORIENTATION SELECTED", "MBR Page Setup", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;

        }

        private void cmdNext_Click(object sender, EventArgs e)
        {
            ShowNextTab();
        }

        private void cmdBack_Click(object sender, EventArgs e)
        {
            ShowPrevTab();
        }

        private void lable3_Click(object sender, EventArgs e)
        {
            cmdFinish_Click(sender, e);
        }

        private void label3_Click(object sender, EventArgs e)
        {
            cmdFinish_Click(sender, e);
        }
    }

}
