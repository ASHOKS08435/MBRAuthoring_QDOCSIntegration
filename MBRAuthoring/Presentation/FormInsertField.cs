using MBRAuthoring.Business;
using MBRAuthoring.DataAccess;
using Microsoft.Office.Interop.Word;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;



namespace MBRAuthoring.Presentation
{
    public partial class FormInsertField : FormBase
    {

        bool PropertExists = false;
        private Word.Document activeDocument = Globals.ThisDocument.ActiveDocument;
        protected override Panel progressBarPanel => panel1;
        string mstrDocName;
        int mDocID;
        string mstrProtectionPassword;
        public IDocumentManager _documentManager = null;

        CommonService objCommonService = new CommonService();
        Word.Selection currentSelection;
        Word.Range selectionRange;
        bool mblnCalledWizard;
        const string MyCompName = "FormInsertField";
        long mlngMyTop;
        long mlngMyLeft;
        qReturn mqrtQReturn;
        public FormInsertField()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();


        }

        private void FormInsertField_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            const string MyProcName = "FormInsertField_Load";

            try
            {
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
                InitializeTabs(tabProperty, btnBack, btnNext, tbProperty);
                ShowTab(0);
                currentSelection = activeDocument.Application.Selection;
                selectionRange = activeDocument.Application.Selection.Range;
                object objMiss = System.Reflection.Missing.Value;
                var result = _documentManager.Document_SelectAttributes(activeDocument.Name);
                mstrDocName = activeDocument.Name;
                mDocID = result.DocumentID;
                mstrProtectionPassword = result.DocumentProtectionPassword;
                SetToolTip(dgvProperty, "DocID:" + mDocID + "Property");
                lblTitle.Text = "Properties for Document " + mstrDocName;
                lblDocProperty.Text = "Properties Selected For " + mstrDocName;
                btnFinish.Enabled = false;
                mblnCalledWizard = false;
                GetProperty();
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

        private void FormInsertField_Activate(object sender, EventArgs e)
        {
            string MyProcName = "UserForm_Activate";
             UserForm__Positioning();
             dgvProperty.Focus();
                //Todo Form_Accelerators_Finish(cmdFinish)
            
        }
        private void UserForm__Positioning()
        {
            if (mlngMyTop > 0 & mlngMyLeft > 0)
                {
                    this.Top = (int)mlngMyTop;
                    this.Left = (int)mlngMyLeft;
                }
            
        }
        private void GetProperty()
        {
            const string MyProcName = "GetProperty";
            MBRAuthDAL objDALMBR = new MBRAuthDAL();
            System.Data.DataTable dtProperty = new System.Data.DataTable();

            try
            {
                dtProperty = objDALMBR.GetDocumentProperty(mDocID);
                if (dtProperty == null)
                {
                    MessageBox.Show("There are no properties available for this DocID");
                }
                else
                {
                    dgvProperty.AutoGenerateColumns = false;
                    dgvProperty.ColumnHeadersVisible = false;
                    dgvProperty.ColumnCount = 3;
                    dgvProperty.Columns[0].Name = "PropertyValueRequired";
                    dgvProperty.Columns[0].DataPropertyName = "PropertyValueRequired";
                    dgvProperty.Columns[0].Visible = false;
                    dgvProperty.Columns[1].Name = "PropertyInstance";
                    dgvProperty.Columns[1].DataPropertyName = "PropertyInstance";
                    dgvProperty.Columns[1].Width = 260;
                    dgvProperty.Columns[2].Name = "PropertyValueDisplay";
                    dgvProperty.Columns[2].DataPropertyName = "PropertyValueDisplay";
                    dgvProperty.Columns[2].Width = 250;

                    dgvProperty.DataSource = dtProperty;
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

            finally
            {
                if (objDALMBR != null)
                { objDALMBR = null; }

            }
        }



        private void cmdPropertyInsert_Click(object sender, System.EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            const string MyProcName = "cmdPropertyInsert_Click";
            int i = 0;
            MBRAuthDAL objDALMBR = new MBRAuthDAL();
            bool CheckProperty;
            Word.Field MyField;



            try
            {
                System.Data.DataTable dtProperty = new System.Data.DataTable();
                dtProperty = objDALMBR.GetDocumentProperty(mDocID);

                if (dgvProperty.SelectedRows.Count > 0)
                {

                    string strPropertyName = dgvProperty.SelectedRows[0].Cells[1].Value.ToString();
                    string strPropertyValue = dgvProperty.SelectedRows[0].Cells[2].Value.ToString();
                    string strPropertyNameUpper = (strPropertyName).ToUpper();


                    if (strPropertyName.Substring(0, 4) == "MOD_" || strPropertyName.Substring(0, 4) == "MTL_" || strPropertyNameUpper == "MODULETABLEROWS")
                    {
                        if (currentSelection.Information[WdInformation.wdWithInTable])
                        {


                            selectionRange.Collapse(WdCollapseDirection.wdCollapseStart);

                            switch (strPropertyNameUpper)
                            {
                                case "FILENAME":
                                    MyField = activeDocument.Fields.Add(selectionRange, Word.WdFieldType.wdFieldFileName, true);
                                    CheckProperty = false;
                                    break;

                                case "PAGE":
                                    MyField = activeDocument.Fields.Add(selectionRange, Word.WdFieldType.wdFieldPage, true);
                                    CheckProperty = false;
                                    break;

                                default:
                                    MyField = activeDocument.Fields.Add(selectionRange, Word.WdFieldType.wdFieldDocProperty, strPropertyName, true);
                                    CheckProperty = true;
                                    break;

                            }
                            CommonService.Field_ApplyShading(MyField);
                            if (CheckProperty)
                            {
                                foreach (var Property in activeDocument.CustomDocumentProperties)
                                {

                                    if (Property.Name.ToUpper() == strPropertyNameUpper)
                                    {

                                        PropertExists = true;
                                        break;
                                    }
                                }

                                if (!PropertExists)
                                {


                                    MessageBox.Show("WARNING: Field references Property  " + strPropertyName + "which DOES NOT EXIST in Document" + activeDocument.Name + "Property will be added to Document.", "Field added", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    CommonService.DocumentProperty_PropertyAdd(DocumentPropertyName: strPropertyName, Bracketed: true, document: activeDocument);


                                }

                            }

                        }


                        else
                        {
                            MessageBox.Show("Unable to place " + strPropertyName + " field, unless cursor is placed into a Word Document table first.", "Cannot Insert Field", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    else
                    {




                        selectionRange.Collapse(WdCollapseDirection.wdCollapseStart);

                        switch (strPropertyNameUpper)
                        {
                            case "FILENAME":
                                MyField = activeDocument.Fields.Add(selectionRange, Word.WdFieldType.wdFieldFileName, true);
                                CheckProperty = false;
                                break;
                            case "PAGE":
                                MyField = activeDocument.Fields.Add(selectionRange, Word.WdFieldType.wdFieldPage, true);
                                CheckProperty = false;
                                break;

                            default:
                                MyField = activeDocument.Fields.Add(selectionRange, Word.WdFieldType.wdFieldDocProperty, strPropertyName, true);
                                CheckProperty = true;
                                break;

                        }
                        CommonService.Field_ApplyShading(MyField);
                        if (CheckProperty)
                        {

                            foreach (var Property in activeDocument.CustomDocumentProperties)
                            {

                                if (Property.Name.ToUpper() == strPropertyNameUpper)
                                {

                                    PropertExists = true;
                                    break;
                                }
                            }

                            if (!PropertExists)
                            {


                                MessageBox.Show("WARNING: Field references Property  " + strPropertyName + " which DOES NOT EXIST in Document " + activeDocument.Name + " Property will be added to Document. ", "Field added", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                CommonService.DocumentProperty_PropertyAdd(DocumentPropertyName: strPropertyName, Bracketed: true, document: activeDocument);

                            }


                        }


                    }


                }



                else
                {
                    MessageBox.Show("Please select a Property to insert");
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
            }
            finally
            {
                if (objDALMBR != null)
                { objDALMBR = null; }
                CommonService.CleanUpBackgroundWordProcesses();
            }

        }



        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }


    }
}

