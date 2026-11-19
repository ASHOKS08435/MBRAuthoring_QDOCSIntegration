using MBRAuthoring.Business;
using MBRAuthoring.DataAccess;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MBRAuthoring.Presentation
{
    public partial class FormInsertTable : FormBase
    {
        const string MyCompName = "FormInsertTable";
        bool PropertExists = false;
        private Word.Document activeDocument = Globals.ThisDocument.ActiveDocument;
        protected override Panel progressBarPanel => panel1;
        string mstrDocName;
        int mDocID;
        string mstrProtectionPassword;
        public IDocumentManager _documentManager = null;

        CommonService objCommonService = new CommonService();
        bool mblnCalledWizard;
        long mlngMyTop;
        long mlngMyLeft;
        qReturn mqrtQReturn;
        string strTableToInsert;
        string strTableToInsertFormatted;
        System.Data.DataTable dtTablelist = new System.Data.DataTable();

        public FormInsertTable()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }
        private void FormInsertTable_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            
            string MyProcName = "FormInsertTable_Load";

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
                InitializeTabs(tabInsertTable, btnBack, btnNext, tbInsertTable); ////pb merge
                ShowTab(0);
                var result = _documentManager.Document_SelectAttributes(activeDocument.Name);
                mstrDocName = activeDocument.Name;
                mDocID = result.DocumentID;
                mstrProtectionPassword = result.DocumentProtectionPassword;
                btnBack.Enabled = false;
                btnNext.Enabled = false;
                if (activeDocument.ProtectionType == Word.WdProtectionType.wdNoProtection)
                { }
                else
                {
                    Document_UnProtect();
                }
                LoadTableListDropdown();
                mblnCalledWizard = false;
                lblTableListTitle.Text = "Insert Selected Table";
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
        private void FormInsertTable_Activate(object sender, EventArgs e)
        {
            bool blnIncludeLocked;
            bool blnIncludeAuthoring;
            bool blnIncludeApproved;
            bool blnIncludeArchived;
            blnIncludeLocked = true;
            blnIncludeAuthoring = true;
            blnIncludeApproved = true;
            blnIncludeArchived = true;
            const string MyProcName = "FormInsertTable_Activate";
            const string cstrNoProject = "";
            const string cstrNoComboSupercedes = "";
            string strActiveTransaction;

                UserForm__Positioning();


                strActiveTransaction = ""; // Todo: Transaction_Active()
                if (strActiveTransaction != "" && strActiveTransaction != mstrDocName)
                {
                    MessageBox.Show("CANNOT RUN THIS LOGIC, DUE TO ACTIVE DATABASE TRANSACTION FOR '" + strActiveTransaction, MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
                    CommonService.CleanUpBackgroundWordProcesses();
                    return;
                }
                if (mblnCalledWizard)
                {
                    btnFinish.Text = "Close";

                }
                else
                {
                    btnFinish.Text = "Finish";

                }
           
        }
        private void LoadTableListDropdown()
        {
            const string MyProcName = "LoadTableListDropdown";
            MBRAuthDAL objDALMBR = new MBRAuthDAL();
            List<string> lstTableList = new List<string>();
            try
            {
                dtTablelist = objDALMBR.GetTablelist();


                if (dtTablelist != null && dtTablelist.Rows.Count > 0)
                {
                    for (int intRow = 0; intRow <= dtTablelist.Rows.Count - 1; intRow++)
                    {
                        string strTableName = dtTablelist.Rows[intRow]["AppStatus"].ToString();
                        int i = strTableName.IndexOf("|");
                        strTableName = strTableName.Substring(0, i);
                        lstTableList.Add(strTableName);
                    }

                    ddlTableList.DisplayMember = "AppStatus";
                    ddlTableList.ValueMember = "AppStatus";
                    ddlTableList.DataSource = lstTableList;
                    ddlTableList.Enabled = true;


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
                if (activeDocument.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    CommonService.DocumentProtection_On(pdocProtectedDocument: activeDocument, pstrProtectionPassword: mstrProtectionPassword, pstrDocumentName: mstrDocName); ;
                }

            }
        }

        private void btnFinish_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            const string MyProcName = "cmdFinish_Click";
            //Todo clsHourGlass objHourGlass;
            Word.Table tbl;
            bool blnProtect;
            try
            {
                if (mblnCalledWizard)
                {
                    UserForm__SHUTDOWN(qReturn.qReturn_Finish);
                    return;
                }

                if (activeDocument.ProtectionType == Word.WdProtectionType.wdNoProtection)
                {
                    blnProtect = false;
                }
                else

                {
                    blnProtect = true;
                    Document_UnProtect();
                }


                strTableToInsert = dtTablelist.Rows[ddlTableList.SelectedIndex]["AppStatus"].ToString();
                int i = strTableToInsert.IndexOf("|");
                int tableDataRowFirst = -1;
                if (i != -1)
                {
                    strTableToInsertFormatted = strTableToInsert.Substring(i + 1);

                }
                CommonService.AutoText_Insert(autotextEntryName: strTableToInsertFormatted, tableDataRowFirst: ref tableDataRowFirst, IspropertyDefine: true, wordApp: activeDocument.Application, document: activeDocument);
                if (strTableToInsertFormatted == "MaterialPrecautions")
                {
                }

                if (blnProtect)
                    Document_Protect();
                SaveManager.Document_Save_FileAndDatabase(activeDocument);
                this.Close();
                CommonService.DocumentProtection_Assure(activeDocument, pstrProtectionPassword: mstrProtectionPassword);
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
                if (activeDocument.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    CommonService.DocumentProtection_On(pdocProtectedDocument: activeDocument, pstrProtectionPassword: mstrProtectionPassword, pstrDocumentName: mstrDocName); ;
                }
            }
            finally
            {
                CommonService.CleanUpBackgroundWordProcesses();
            }


        }




        private void ddlTableList_SelectedIndexChanged(object sender, EventArgs e)
        {
            strTableToInsert = ddlTableList.SelectedItem.ToString();

        }

        private void Document_Protect()
        {
            string MyProcName = "Document_Protect";
            try
            {

                if (mblnCalledWizard)
                {

                }
                CommonService.DocumentProtection_On(activeDocument, mstrProtectionPassword);
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

        private void Document_UnProtect()
        {
            string MyProcName = "Document_UnProtect";
            try
            {

                if (mblnCalledWizard)
                {

                }
                CommonService.DocumentProtection_Off(activeDocument, mstrProtectionPassword);
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

        private void UserForm__SHUTDOWN(qReturn pqrtQReturn)
        {
            mqrtQReturn = pqrtQReturn;
            if (!mblnCalledWizard)
            {
                MessageBox.Show("ERROR: Called UserForm__SHUTDOWN(" + pqrtQReturn + ")" + ", but this is not a CalledWizard.", MyCompName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (this.Visible)
            {
                this.Hide();
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



    }
}

