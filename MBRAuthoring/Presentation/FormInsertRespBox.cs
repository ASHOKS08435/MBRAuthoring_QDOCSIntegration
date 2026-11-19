using MBRAuthoring.Business;
using Microsoft.Office.Interop.Word;
using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MBRAuthoring.Presentation
{
    public partial class FormInsertRespBox : FormBase
    {

        private Word.Document activeDocument = Globals.ThisDocument.ActiveDocument;
        protected override Panel progressBarPanel => panel1;
        const string MyCompName = "FormInsertRespBox";
        bool mblnCalledWizard;
        long mlngMyTop;
        long mlngMyLeft;
        qReturn mqrtQReturn;
        string mstrRespBxSelected;
        public FormInsertRespBox()
        {
            InitializeComponent();
        }
        private void FormInsertRespBox_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            string MyProcName = "FormInsertRespBox_Load";

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
                InitializeTabs(tabResponseBox, btnBack, btnNext, tbResponseBox);

                ShowTab(0);
                mblnCalledWizard = false;
                lblTitle.Text = "Insert Response Box Wizard";
                btnRspBoxInsert.Enabled = false;
                btnFinish.Enabled = false;
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
        private void FormInsertRespBox_Activate(object sender, EventArgs e)
        {
           
                UserForm__Positioning();
                grpResponseBox.Focus();
                rdoYesNo.Focus();
                rdoYesNo.Paint += rdoYesNo_Paint;
                //Todo Form_Accelerators_Finish(cmdFinish)
            
        }

        private void rdoYesNo_Paint(object sender, PaintEventArgs e)
        {
            using (Pen dashedPen = new Pen(Color.Black, 2) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dot })
            {
                e.Graphics.DrawRectangle(dashedPen, 0, 0, rdoYesNo.Width - 1, rdoYesNo.Height - 1);
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
        protected void rdoYesNo_CheckedChanged(object sender, EventArgs e)
        {
            
                btnRspBoxInsert.Enabled = true;
                mstrRespBxSelected = "RspBxYesNo";
                btnFinish.Enabled = true;
            

        }
        protected void rdoDonebyDate_CheckedChanged(object sender, EventArgs e)
        {
            
                btnRspBoxInsert.Enabled = true;
                mstrRespBxSelected = "RspBxDoneDate";
                btnFinish.Enabled = true;
            
        }
        protected void rdoDoneByCheckedBy_CheckedChanged(object sender, EventArgs e)
        {
           
                btnRspBoxInsert.Enabled = true;
                mstrRespBxSelected = "RspBxDoneChkDate";
                btnFinish.Enabled = true;
            
        }
        protected void rdoEntryBox_CheckedChanged(object sender, EventArgs e)
        {
            btnRspBoxInsert.Enabled = true;
            mstrRespBxSelected = "RspBxEntry";
            btnFinish.Enabled = true;
            
        }

        private void btnRspBoxInsert_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            const string MyProcName = "btnRspBoxInsert_Click";
            try
            {
                Word.Range rng = activeDocument.Application.Selection.Range;
                int intSubLoopLimiter;
                int tableDataRowFirst = -1;
                Microsoft.Office.Interop.Word.Selection currentSelection = activeDocument.Application.Selection;
                CommonService.AutoText_Insert(autotextEntryName: mstrRespBxSelected, tableDataRowFirst: ref tableDataRowFirst, rangeAutoText: rng, wordApp: activeDocument.Application);
                //rng.Select();
                //rng.Tables[1].Rows.Alignment = WdRowAlignment.wdAlignRowRight;
                intSubLoopLimiter = 1;


                intSubLoopLimiter = 2000;
                while (currentSelection.Information[WdInformation.wdWithInTable] & intSubLoopLimiter > 0)
                {
                    intSubLoopLimiter = intSubLoopLimiter - 1;
                    currentSelection.MoveDown(Unit: WdUnits.wdLine, Count: 1);
                }

                grpResponseBox.Focus();
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
            }
            finally
            {
                
                CommonService.CleanUpBackgroundWordProcesses();
            }
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }

        private void btnFinish_Click(object sender, EventArgs e)
        {

            const string MyProcName = "btnFinish_Click";
            try
            {
                btnRspBoxInsert_Click(sender, e);
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
    }
}
