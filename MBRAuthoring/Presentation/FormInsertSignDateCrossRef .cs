using MBRAuthoring.Business;
using Microsoft.Office.Interop.Word;
using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MBRAuthoring.Presentation
{
    public partial class FormInsertSignDateCrossRef : FormBase
    {
        private Word.Document activeDocument = Globals.ThisDocument.ActiveDocument;
        const string MyCompName = "FormInsertSignDateCrossRef";
        const string MsgBoxTitle = "Insert Field";
        protected override Panel progressBarPanel => panel1;
        string mstrDocName;
        int mDocID;
        string mstrProtectionPassword;
        public IDocumentManager _documentManager = null;

        bool mblnCalledWizard;
        long mlngMyTop;
        long mlngMyLeft;
        qReturn mqrtQReturn;
        public FormInsertSignDateCrossRef()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();
        }
        private void FormInsertSignDateCrossRef_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            const string MyProcName = "FormInsertSignDateCrossRef_Load";

            try
            {
                foreach (Control ctrl in progressBarPanel.Controls)//pb merge
                {
                    if (ctrl is Label label)
                    {
                        if (label.Tag == null || label.Tag.ToString() != "Exclude")
                        {
                            label.Click += ProgressLabel_Click;
                        }
                    }
                }
                InitializeTabs(tabStepCrossRef, btnBack, btnNext, tbStepCrossRef); ////pb merge
                ShowTab(0);

                var result = _documentManager.Document_SelectAttributes(activeDocument.Name);
                mstrDocName = activeDocument.Name;
                mDocID = result.DocumentID;
                mstrProtectionPassword = result.DocumentProtectionPassword;
                mblnCalledWizard = false;
                lblTitle.Text = "Select Sign && Date Numbered Step Cross Reference";
                LoadNumberedItems();
                optList123.Enabled = false;
                optList1Thru3.Enabled = false;
                optList123.Visible = false;
                optList1Thru3.Visible = false;
                btnFinish.Enabled = false;
                btnResetSelection.Enabled = false;
                btnResetSelection.Visible = false;
                btnBack.Enabled = false;
                btnNext.Enabled = false;
                Cursor = Cursors.Default;
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
        private void FormInsertSignDateCrossRef_Activate(object sender, EventArgs e)
        {
                string MyProcName = "UserForm_Activate";
                UserForm__Positioning();
                lstNumberedItems.Focus();
                //lstNumberedItems.SelectedIndex = 0;
                //lstNumberedItems.DrawItem += lstNumberedItems_DrawItem;
                ////Todo Form_Accelerators_Finish(cmdFinish)
            
        }

        private void lstNumberedItems_DrawItem(object sender, DrawItemEventArgs e)
        {

            if (e.Index == 0)
            {

                using (Pen dashedPen = new Pen(Color.Black, 1) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dot })
                {
                    e.Graphics.DrawRectangle(dashedPen, e.Bounds.Left, e.Bounds.Top, e.Bounds.Width, e.Bounds.Height);
                }
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
        private void LoadNumberedItems()
        {
            string MyProcName = "LoadNumberedItems";
            try
            {
                object NumberedItems = activeDocument.GetCrossReferenceItems(WdReferenceType.wdRefTypeNumberedItem);
                if (NumberedItems != null)
                {
                    Array a = (Array)NumberedItems;
                    int b = a.Length;
                    for (int i = 1; i < b; i++)
                    {
                        lstNumberedItems.Items.Add(a.GetValue(i));
                        //string item = (string)aa.GetValue(1);

                    }
                }

            }
            catch (Exception ex)
            {
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

        private void btnResetSelection_Click(object sender, EventArgs e)
        {
            int i;
            int SelectedCount = -1;
            btnFinish.Enabled = false;

           
                if (lstNumberedItems.SelectedIndex != 0)
                {
                    lstNumberedItems.ClearSelected();

                }
                
                if (SelectedCount == 2)
                {
                    optList123.Enabled = true;
                    optList1Thru3.Enabled = true;
                    lstNumberedItems.Enabled = false;
                }
                else
                {
                    optList123.Enabled = false;
                    optList1Thru3.Enabled = false;
                    lstNumberedItems.Enabled = true;
                }

        }

        private void lstNumberedItems_KeyDown(object sender, KeyEventArgs e)
        {
            btnFinish.Enabled = true;
            btnResetSelection.Enabled = true;
        }
        private void lstNumberedItems_KeyUp(object sender, KeyEventArgs e)
        {
            int i;
            int SelectedCount = 0;
            btnFinish.Enabled = true;
            btnResetSelection.Enabled = true;

            for (i = 0; i <= lstNumberedItems.Items.Count - 1; i++)
            {
                if (lstNumberedItems.SelectedIndex != 0)
                {
                    SelectedCount++;
                }
            }
            if (SelectedCount == 2)
            {
                optList123.Enabled = true;
                optList1Thru3.Enabled = true;
                lstNumberedItems.Enabled = false;
            }
            else
            {
                optList123.Enabled = false;
                optList1Thru3.Enabled = false;
                lstNumberedItems.Enabled = true;
            }
        }


        private void lstNumberedItems_DoubleClick(object sender, EventArgs e)
        {
            btnFinish_Click(sender, e);
        }

        private void lstNumberedItems_Click(object sender, EventArgs e)
        {
            btnFinish.Enabled = true;
            btnResetSelection.Enabled = true;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            string MyProcName = "btnCancel_Click";
            int i;
            try
            {

                //for (i = 0; i <= lstNumberedItems.Items.Count-1; i++)
                //{
                if (Convert.ToInt32(lstNumberedItems.SelectedIndex) != -1)
                {

                    DialogResult dialogResult = MessageBox.Show("Items selected not placed." + "\r\n" + "Do you wish to cancel placement of selected item(s)?", "Cancel Placement of Selected", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.Yes)
                    {
                        this.Close();
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }
                else
                {
                    this.Close();
                }

                //}


            }
            catch (Exception ex)
            {
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
                CommonService.CleanUpBackgroundWordProcesses();
            }
        }

        private void btnFinish_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            const string MyProcName = "btnFinish_Click";
            Microsoft.Office.Interop.Word.Selection currentSelection = activeDocument.Application.Selection;
            int i;
            int index = lstNumberedItems.SelectedIndex;
            try
            {
                for (i = 0; i <= lstNumberedItems.Items.Count - 1; i++)
                {
                    if (i == index)
                    {
                        if (activeDocument.ProtectionType != WdProtectionType.wdNoProtection)
                        {
                            CommonService.DocumentProtection_Off(pdocProtectedDocument: activeDocument, pstrProtectionPassword: mstrProtectionPassword, pstrDocumentName: mstrDocName);
                        }
                        if (currentSelection.Information[WdInformation.wdWithInTable])
                        {
                            currentSelection.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        }
                        currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                        currentSelection.Font.Bold = 1;
                        currentSelection.InsertCrossReference(ReferenceType: WdReferenceType.wdRefTypeNumberedItem, ReferenceKind: WdReferenceKind.wdNumberFullContext, ReferenceItem: i + 1, InsertAsHyperlink: true, IncludePosition: false, SeparateNumbers: false, SeparatorString: " ");

                    }
                }
                currentSelection.Font.Bold = 0;
                if (activeDocument.ProtectionType != WdProtectionType.wdNoProtection)
                {
                    CommonService.DocumentProtection_On(pdocProtectedDocument: activeDocument, pstrProtectionPassword: mstrProtectionPassword, pstrDocumentName: mstrDocName); ;
                }
                Cursor = Cursors.Default;
                this.Close();
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
                if (activeDocument.ProtectionType != WdProtectionType.wdNoProtection)
                {
                    CommonService.DocumentProtection_On(pdocProtectedDocument: activeDocument, pstrProtectionPassword: mstrProtectionPassword, pstrDocumentName: mstrDocName); ;
                }

            }
            finally
            {
                CommonService.CleanUpBackgroundWordProcesses();
            }
        }
    }
}
