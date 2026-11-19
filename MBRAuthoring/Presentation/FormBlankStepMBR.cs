using MBRAuthoring.Business;
using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
namespace MBRAuthoring.Presentation
{
    public partial class FormBlankStepMBR : FormBase
    {

        public IDocumentManager _documentManager = null;

        private Word.Document activeDocument = Globals.ThisDocument.ActiveDocument;
        protected override Panel progressBarPanel => panel1;

        object objMiss = System.Reflection.Missing.Value;
        const string MyCompName = "FormBlankStepMBR";
        string mstrDocName;
        int mDocID;

        public enum StepTextType
        {
            qtMainStep = 8904
        }
        public FormBlankStepMBR()
        {
            InitializeComponent();
        }


        private void FormBlankStepMBR_Load(object sender, EventArgs e)
        {
            
            const string MyProcName = "FormBlankStepMBR_Load";

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
                InitializeTabs(tbBlankStep, cmdBack, cmdNext, tbAPI, tbDP); ////pb merge
                ShowTab(0);
                cmdFinish.Enabled = false;

                mstrDocName = activeDocument.Name;
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

        private void FormBlankStepMBR_Activate(object sender, EventArgs e)
        {
            string MyProcName = "UserForm_Activate";
            tbAPI.Enabled = true;
            tbAPI.Visible = true;
            cmdInsertStep.Focus();
            cmdInsertStep.Paint += cmdInsertStep_Paint;
            cmdInsertHeader.Focus();
            cmdInsertHeader.Paint += cmdInsertHeader_Paint;
            tbDP.Enabled = true;
            tbDP.Visible = true;
        }
        private void cmdInsertStep_Paint(object sender, PaintEventArgs e)
        {
            using (Pen dashedPen = new Pen(Color.Black, 1) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dot })
            {
                e.Graphics.DrawRectangle(dashedPen, 0, 0, cmdInsertStep.Width - 5, cmdInsertStep.Height - 5);
            }
        }

        private void cmdInsertHeader_Paint(object sender, PaintEventArgs e)
        {
            using (Pen dashedPen = new Pen(Color.Black, 1) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dot })
            {
                e.Graphics.DrawRectangle(dashedPen, 0, 0, cmdInsertHeader.Width - 5, cmdInsertHeader.Height - 5);
            }
        }
        private void cmdInsertStep_Click(object sender, EventArgs e)
        {
            string MyProcName = "cmdInsertStep_Click";

            try
            {
                Word.Selection currentSelection = activeDocument.Application.Selection;
                if ((txtStepName.Text).Trim() == "")
                {
                    MessageBox.Show("STEPNAME is required for MAIN Step", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtStepName.Select();
                    return;
                }

                CommonService.TypeStepText(MBRAuthoring.Business.StepTextType.qtMainStep, "BLANK_ONLY", txtStepName.Text);
                currentSelection.MoveLeft();
                currentSelection.Delete();
                this.Close();
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
        private void cmdInsertSubstep_Click(object sender, EventArgs e)
        {
            string MyProcName = "cmdInsertSubstep_Click";
            try
            {
                Word.Selection currentSelection = activeDocument.Application.Selection;
                CommonService.TypeStepText(MBRAuthoring.Business.StepTextType.qtSubStep, "BLANK_ONLY");
                currentSelection.MoveLeft();
                currentSelection.Delete();
                this.Close();
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
        

        private void cmdInsertAutomationSubStep_Click(object sender, EventArgs e)
        {
            string MyProcName = "cmdInsertAutomationSubStep_Click";
            try
            {
                Word.Selection currentSelection = activeDocument.Application.Selection;
                Microsoft.Office.Interop.Word.Table objTab1;
                int MatTblrow = 1;
                int MatTblcolumn = 2;
                objTab1 = activeDocument.Tables.Add((Microsoft.Office.Interop.Word.Range)activeDocument.Application.Selection.Range, MatTblrow, MatTblcolumn, ref objMiss, ref objMiss);
                objTab1.Rows[1].SetLeftIndent(activeDocument.Application.InchesToPoints((float)0.75), Word.WdRulerStyle.wdAdjustSameWidth);
                int iRowMatTbl, iColMatTbl;
                for (iColMatTbl = 1; iColMatTbl <= MatTblcolumn; iColMatTbl++)
                {
                    for (iRowMatTbl = 1; iRowMatTbl <= MatTblrow; iRowMatTbl++)
                    {

                        if (iRowMatTbl == 1 && iColMatTbl == 1)
                        {
                            objTab1.Cell(1, 1).Width = activeDocument.Application.InchesToPoints((float)0.75);
                            objTab1.Cell(1, 1).Height = 20;
                            objTab1.Cell(1, 1).Range.Font.Italic = 1;
                            objTab1.Cell(1, 1).Range.Font.Name = "Times New Roman";
                            objTab1.Cell(1, 1).Range.Font.Size = 12;

                            Word.Paragraph para1 = activeDocument.Paragraphs.Add(currentSelection.Range);
                            para1.Range.Text = ("a.");
                            object styleName = "Normal";
                            para1.Range.set_Style(ref styleName);
                            objTab1.Cell(1, 1).Select();
                            currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            currentSelection.SetRange(para1.Range.End, para1.Range.End);
                        }






                        if (iRowMatTbl == 1 && iColMatTbl == 2)
                        {
                            objTab1.Cell(1, 2).Width = activeDocument.Application.InchesToPoints((float)5.33); ;
                            objTab1.Cell(1, 2).Height = 20;
                            objTab1.Cell(1, 2).Range.Font.Italic = 1;
                            objTab1.Cell(1, 2).Range.Font.Name = "Times New Roman";
                            objTab1.Cell(1, 2).Range.Font.Size = 12;
                            objTab1.Cell(1, 2).Select();
                            currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            Word.Range CellRange = objTab1.Cell(1, 2).Range;
                            CellRange.Select();

                        }

                        currentSelection.MoveRight(Unit: Word.WdUnits.wdCharacter, Count: 1);

                        //objTab1.Rows[iRowMatTbl].Cells[iColMatTbl].Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(Color.LightGray);
                        objTab1.Cell(iRowMatTbl, iColMatTbl).Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                        objTab1.Cell(iRowMatTbl, iColMatTbl).Range.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                        objTab1.Cell(iRowMatTbl, iColMatTbl).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                        objTab1.Cell(iRowMatTbl, iColMatTbl).Range.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                        //objTab1.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent);
                    }
                }
				this.Close();
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


        private void cmdInsertSubSubStep_Click(object sender, EventArgs e)
        {
            string MyProcName = "cmdInsertSubSubStep_Click";
            try
            {
                Word.Selection currentSelection = activeDocument.Application.Selection;
                CommonService.TypeStepText(MBRAuthoring.Business.StepTextType.qtSubSubStep, "BLANK_ONLY");
                currentSelection.MoveLeft();
                currentSelection.Delete();
                this.Close();
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


        private void cmdInsertHeader_Click(object sender, EventArgs e)
        {
            string MyProcName = "cmdInsertHeader_Click";
            try
            {
                CommonService.TypeStepText(MBRAuthoring.Business.StepTextType.qtMBRHeading, "blank_only");
                CommonService.Table_Exit();
                CommonService.TypeStepText(MBRAuthoring.Business.StepTextType.qtMBRStep, "blank_only");
                Insert_RecordingBlocks();
                Insert_SignatureBlocks();
                this.Close();
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
        private void Insert_RecordingBlocks()
        {
            int intBlockCounter;

            try
            {
                for (intBlockCounter = 1; intBlockCounter <= spnRecordingBlocks.Value; intBlockCounter++)
                {
                    CommonService.RecordingBlock(Caption: "<data" + intBlockCounter + ">");
                }
            }

            catch (Exception ex)
            {
                throw;
            }
        }

        private void Insert_SignatureBlocks()
        {
            int intBlockCounter;
            try
            {
                for (intBlockCounter = 1; intBlockCounter <= spnSignatureBlocks.Value; intBlockCounter++)
                {
                    CommonService.SignAndDate(CheckedBy: true);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void cmdInsertProcessStep_Click(object sender, EventArgs e)
        {
            string MyProcName = "cmdInsertProcessStep_Click";
            try
            {
                CommonService.TypeStepText(MBRAuthoring.Business.StepTextType.qtMBRStep, "blank_only");
                Insert_RecordingBlocks();
                Insert_SignatureBlocks();
                this.Close();
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

        private void cmdInsertProcessSubStep_Click(object sender, EventArgs e)
        {
            string MyProcName = "cmdInsertProcessSubStep_Click";
            try
            {
                CommonService.TypeStepText(MBRAuthoring.Business.StepTextType.qtMBRSub, "blank_only");
                Insert_RecordingBlocks();
                Insert_SignatureBlocks();
                this.Close();
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
        private void spnRecordingBlocks_ValueChanged(object sender, EventArgs e)
        {


            if (spnRecordingBlocks.Value > 5)
            {
                spnRecordingBlocks.Value = 0;
                MessageBox.Show("The maximum value allowed is 5,","", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }


        }
        private void spnSignatureBlocks_ValueChanged(object sender, EventArgs e)
        {
            if (spnSignatureBlocks.Value > 5)
            {
                spnSignatureBlocks.Value = 0;
                MessageBox.Show("The maximum value allowed is 5","", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        //private void spnRecordingBlocks_ValueChanged(object sender, EventArgs e)
        //{
        //    TextBox txtSpin = (TextBox)spnRecordingBlocks.Controls[1];//notice the textbox is the 2nd control in the numericupdown control

        //    int t = Convert.ToInt32(spnRecordingBlocks.Value);
        //    if (txtSpin.Text == "")
        //    {
        //        txtSpin.Text = Convert.ToString(spnRecordingBlocks.Value);
        //        return;
        //    }

        //    else if ((Convert.ToInt32(spnRecordingBlocks.Maximum) < t) | (Convert.ToInt32(spnRecordingBlocks.Minimum) > t))
        //    {
        //        txtSpin.Text = Convert.ToString(spnRecordingBlocks.Value);
        //        return;
        //    }
        //    else
        //    {
        //        spnRecordingBlocks.Value = Convert.ToInt32(txtSpin.Text);
        //    }
        //}

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }
        private void cmdNext_Click(object sender, EventArgs e)
        {
             ShowNextTab();
           
        }
        private void cmdBack_Click(object sender, EventArgs e)
        {
             ShowPrevTab();
            
        }

    }
}



