using MBRAuthoring.Business;
using MBRAuthoring.DataAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Linq;


namespace MBRAuthoring.Presentation
{
    public partial class FormEditModule : FormBase
    {
        const string MyCompName = "FormEditModule";
        protected override Panel progressBarPanel => panel1;
        System.Data.DataTable dtModuleAll = new System.Data.DataTable();
        System.Data.DataTable dtDocModule = new System.Data.DataTable();
        bool IsSuccess = false;
        bool IsNew = false;
        bool IsUpdate = false;
        bool IsSave = false;

        string TableRows;
        string strModule = string.Empty;
        string strPacket = string.Empty;
        public IDocumentManager _documentManager = null;
        //public ICommandBarManager _commandBarManager;
        private Word.Document activeDocument = Globals.ThisDocument.ActiveDocument;
        string mstrDocName;
        string mstrDocFullName;
        int mDocID;
        string mstrProtectionPassword;
        string mstrOpUnit;
        int mintOpUnitID;
        string mstrOpUnitType;
        int mintParmsValidated;
        const int cintParmsValidated = 16539;
        bool mblnCalledWizard;
        bool mblnUserMadeModuleChange;
        string[] astrModuleTableRows;
        bool blnSpnReachedMax;
        bool blnSpnReachedMin;
        //made changes to this as per requirement 05/11
        const int qdmColDocModuleID = 6;
        const int qdmColModuleBaseName = 0;
        const int qdmColModuleModuleTitle = 1;
        const int qdmColModuleSequence = 9;
        const int qdmColModuleRepetitions = 2;
        const int qdmColModuleUsage = 3;
        const int qdmColModulePacket = 4;
        const int qdmColModuleModuleTableRows = 5;
        const int qdmColModuleDocumentTitle = 7;
        qReturn mqrtQReturn;
        qRequest mqrqCalledRequest;
        const int cintEditModuleRun = 4592;
        public const int qdmColCount = 9;
        int mlngMyTop;
        int mlngMyLeft;
        bool blnDBUpdateSuccessful;
        //Performancetuning
        private System.Windows.Forms.Timer wordKeepAliveTimer;

        enum ModuleValidationMode
        {
            ModuleValidation_Add = 2177,
            ModuleValidation_Update = 17801
        }



        public FormEditModule()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();
            dgvModuleSelected.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                SetTargetDocumentName(activeDocument?.Name, activeDocument?.FullName);
            }
            catch
            {
            }
        }
        private void FormEditModule_Load(object sender, EventArgs e)
        {

            const string MyProcName = "FormEditModule_Load";
            try
            {
                if (!TryResolveActiveDocumentForForm())
                {
                    throw new InvalidOperationException("Unable to resolve the target document.Close the form and try again.");
                }
                InitializeTabs(tabEditModule, cmdBack, cmdNext, tabModule);//add two buttons
                ShowTab(0);

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
                
                if (!TryResolveActiveDocumentForForm())
                {
                    throw new InvalidOperationException("Unable to resolve the target document.Close the form and try again.");
                }

                var result = _documentManager.Document_SelectAttributes(activeDocument.Name);
                mstrDocName = activeDocument.Name;
                mDocID = result.DocumentID;
                mstrProtectionPassword = result.DocumentProtectionPassword;
                mstrOpUnit = result.DocumentOpUnit;
                mintOpUnitID = result.DocumentOpUnitID;
                mstrOpUnitType = result.DocumentOpUnitType;
                mblnCalledWizard = false;
                lblTitle.Text = "Modules for Operating Unit " + mstrOpUnit;
                lblModuleSelected.Text = "Modules selected for : " + mstrDocName;
                GetModuleAll();
                GetPacketbyDoc();

                lblModuleTitle.Visible = false;
                txtModuleTitle.Visible = false;
                lblModuleTableRows.Visible = false;
                SpnUpDown.Visible = false;
                chkTableRows.Visible = false;
                chkTableRows.Enabled = false;
                rdoBatchRecord.Enabled = true;
                rdoBatchRecord.Checked = true;
                btnModuleUpdate.Enabled = false;
                btnModuleDelete.Enabled = false;
                ModuleReset();
                SetToolTip(rdoBatchRecord, "Usage for this module");
                SetToolTip(rdoInfoOnly, "Usage for this module");
                SetToolTip(dgvModuleSelected, "dgvModuleSelected: DocID=" + mDocID + " Module");
                SetToolTip(btnCancel, "Terminate Function with NO Output");
                SetToolTip(btnFinish, "Save Document, then Terminate Function");
                SetToolTip(btnModuleAdd, "Add above selected Module");
                SetToolTip(cmdBack, "Show Previous Wizard Page");
                SetToolTip(cmdNext, "Show Next Wizard Page");
                SetToolTip(btnModuleDelete, "Delete Module");
                SetToolTip(btnModuleUpdate, "Update Module");
                SetToolTip(lblMap6, "To skip to this step, click here.");
                SetToolTip(shp0, "To skip to this step, click here.");
                SetToolTip(lblMap5, "To skip to this step, click here.");
                SetToolTip(lblMap0, "To skip to this step, click here.");

                //TDo: Move to Activate block

                //Performancetuning
                wordKeepAliveTimer = new System.Windows.Forms.Timer();
                wordKeepAliveTimer.Interval = 1 * 60 * 1000; // every 1 minute (adjust if needed)
                wordKeepAliveTimer.Tick += WordKeepAliveTimer_Tick;
                wordKeepAliveTimer.Start();
                //end
            }
            catch
            {

            }
        }
        //private void FormEditModule_Activate(object sender, EventArgs e)
        private void FormEditModule_Shown(object sender, EventArgs e)
        {
            bool blnIncludeLocked;
            bool blnIncludeAuthoring;
            bool blnIncludeApproved;
            bool blnIncludeArchived;
            blnIncludeLocked = true;
            blnIncludeAuthoring = true;
            blnIncludeApproved = true;
            blnIncludeArchived = true;
            //const string MyProcName = "FormEditModule_Activate";
            const string MyProcName = "FormEditModule_Shown";
            UserForm__Positioning();
            string strActiveTransaction;

            strActiveTransaction = DatabaseHelper.Transaction_Active();
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
                //int ColumnCount = dgvModuleSelected.Columns.Count;
                int ColumnCount = dgvModuleSelected.Columns.Count;
                if (ColumnCount == 1)
                {
                    ColumnCount = 9;
                }
                //Fill_Document_Module_ColumnWidths(ModuleListBox: lstModuleSelected);
            }
            else
            {
                btnFinish.Text = "Save";
                GetDocModule();
            }

            //ListBox_CopyAll(lstModuleSelected, lstModuleOriginal1);
            SpnUpDown.Minimum = 1;
            //SpnUpDown.Maximum = 1;
            //SpnUpDown.Value = 1;
            //Form_Accelerators_Finish(cmdFinish);
        }


        /*private void FormEditModule_Activate(object sender, EventArgs e)
        {
            bool blnIncludeLocked;
            bool blnIncludeAuthoring;
            bool blnIncludeApproved;
            bool blnIncludeArchived;
            blnIncludeLocked = true;
            blnIncludeAuthoring = true;
            blnIncludeApproved = true;
            blnIncludeArchived = true;
            const string MyProcName = "FormEditModule_Activate";
            UserForm__Positioning();
            string strActiveTransaction;

            strActiveTransaction = ""; // Todo: Transaction_Active()
            if (strActiveTransaction != "" && strActiveTransaction != mstrDocName)
            {
                MessageBox.Show("CANNOT RUN THIS LOGIC, DUE TO ACTIVE DATABASE TRANSACTION FOR '" + strActiveTransaction, MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.Close();
                return;
            }
            if (mblnCalledWizard)
            {
                btnFinish.Text = "Close";
                int ColumnCount = dgvModuleSelected.Columns.Count;
                if (ColumnCount == 1)
                {
                    ColumnCount = 9;
                }
                //Fill_Document_Module_ColumnWidths(ModuleListBox: lstModuleSelected);
            }
            else
            {
                btnFinish.Text = "Save";
                GetDocModule();
            }

            //ListBox_CopyAll(lstModuleSelected, lstModuleOriginal1);
            SpnUpDown.Minimum = 1;
            SpnUpDown.Maximum = 1;
            SpnUpDown.Value = 1;
            //Form_Accelerators_Finish(cmdFinish);
        }*/

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
        private void QInitialize(qRequest pqrqRequest)
        {
            const string MyProcName = "QInitialize";
            mblnCalledWizard = true;
            mqrqCalledRequest = pqrqRequest;
            switch (pqrqRequest)
            {
                case qRequest.qRequest_Display:
                    {
                        mintParmsValidated = 0;
                        break;
                    }

                case qRequest.qRequest_Validate:
                    {
                        mintParmsValidated = 0;
                        break;
                    }

                case qRequest.qRequest_Write:
                    {
                        if (mintParmsValidated != cintParmsValidated)
                        {
                            //Todo Error Handling Information.Err.Raise(42981, MyCompName + "." + MyProcName, "Attempting QWrite with unvalidated data");
                            var ex = new Exception();
                            ex.Data["errNum"] = "42981";
                            ex.Data["errDesc"] = "Attempting QWrite with unvalidated data";
                            throw ex;
                        }
                        break;
                    }
            }

            //shpMapCollapse.Visible = true;
        }

        public qReturn QDisplay(QDisplay_FirstLast QFirstLast)
        {
            QInitialize(qRequest.qRequest_Display);
            //this.Show();
            //this.ShowDialog();
            //this.Dispose();
            return mqrtQReturn;
        }

        public void QLet(int QTop, int QLeft, qpsEditModuleParms p)
        {
            const string MyProcName = "QLet";
            mlngMyTop = QTop;
            mlngMyLeft = QLeft;
            if (p.intEditModuleRun != 0 & p.intEditModuleRun != cintEditModuleRun)
            {
                //Todo Information.Err.Raise(41987, MyCompName + ".QLET", "INVALID RUNCODE=" + p.intEditModuleRun);
                var ex = new Exception();
                ex.Data["errNum"] = "41987";
                ex.Data["errDesc"] = "INVALID RUNCODE=" + p.intEditModuleRun;
                throw ex;
            }
            if (p.intEditModuleValidated != 0 & p.intEditModuleValidated != cintParmsValidated)
            {
                //Todo Information.Err.Raise(41989, MyCompName + ".QLET", "INVALID VALCODE=" + p.intEditModuleValidated);
                var ex = new Exception();
                ex.Data["errNum"] = "41989";
                ex.Data["errDesc"] = "INVALID RUNCODE=" + p.intEditModuleValidated;
                throw ex;
            }
            if (p.intEditModuleRun == 0)
            {
                p.strEditModuleSelected = new string[0, 0];

            }


            mintParmsValidated = p.intEditModuleValidated;
            if (p.strEditModuleSelected.GetUpperBound(1) > 0)
            {
                DataGridView_From_Array(dgvModuleSelected, p.strEditModuleSelected);
                if (dgvModuleSelected.ColumnCount != qdmColCount)
                {
                    MessageBox.Show("Expected " + qdmColCount + " Columns," + "detected " + dgvModuleSelected.ColumnCount + " Columns.", MyCompName + "." + MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                dgvModuleSelected.ColumnCount = qdmColCount;
            }
        }

        public void QGet(long QTop, long QLeft, qpsEditModuleParms p)
        {
            QTop = this.Top;
            QLeft = this.Left;
            p.intEditModuleRun = cintEditModuleRun;
            p.intEditModuleValidated = mintParmsValidated;
            DataGridView_To_Array(dgvModuleSelected, p.strEditModuleSelected);
        }

        public qReturn QValidate()
        {
            QInitialize(qRequest.qRequest_Validate);
            if (Valid_Entries())
            {
                return qReturn.qReturn_Valid;
            }
            else
            {
                this.Show();
                return mqrtQReturn;
            }
        }

        public qReturn QWrite(QWriteWhat pqwwWriteWhat)
        {
            const string MyProcName = "QWrite";
            QInitialize(qRequest.qRequest_Write);
            switch (pqwwWriteWhat)
            {
                case QWriteWhat.qWriteDBUpdate:
                    {
                        GetDocModule();
                        Setup_Database();
                        break;
                    }

                case QWriteWhat.qWriteMBRTable:
                    {
                        Setup_Document();
                        break;
                    }

                default:
                    {
                        MessageBox.Show("UNEXPECTED REQUEST '" + pqwwWriteWhat + "'", MyCompName + "." + MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }

            return qReturn.qReturn_Written;
        }

        public void QClose()
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }

        private void FormEditModule_FormClosed(object sender, FormClosedEventArgs e)
        {
           
            CommonService.CleanUpBackgroundWordProcesses();

        }

        private void GetModuleAll()
        {
            MBRAuthDAL objDALMBR = new MBRAuthDAL();

            List<string> lstModuleAll = new List<string>();
            var dict = new Dictionary<int, string>();
            int OpUnitID = mintOpUnitID;

            //lstModuleAll.Insert(0, "--Select--"); //Soundhar 010825
            dtModuleAll = objDALMBR.GetModuleAll(OpUnitID);
            DataRow selectdummy = dtModuleAll.NewRow();
            selectdummy[0] = "";
            selectdummy[1] = "--Select--";
            selectdummy[2] = "0";
            selectdummy[3] = "0";

            dtModuleAll.Rows.InsertAt(selectdummy, 0);
            if (dtModuleAll != null && dtModuleAll.Rows.Count > 0)
            {
                for (int intRow = 0; intRow <= dtModuleAll.Rows.Count - 1; intRow++)
                {

                    lstModuleAll.Add(dtModuleAll.Rows[intRow]["DocumentBaseName"].ToString() + " - " + dtModuleAll.Rows[intRow]["DocumentTitle"].ToString().Trim());
                }

                ddlModuleAll.DisplayMember = "DocumentBaseName";
                ddlModuleAll.ValueMember = "DocumentTitle";
                ddlModuleAll.DataSource = lstModuleAll;
                ddlModuleAll.Enabled = true;
                ddlModuleAll.SelectedIndex = 0;

            }

        }
        private void GetPacketbyDoc()
        {
            MBRAuthDAL objDALMBR = new MBRAuthDAL();
            System.Data.DataTable dtPacket = new System.Data.DataTable();
            int DocID = mDocID;

            dtPacket = objDALMBR.GetPacketbyDoc(DocID);
            if (dtPacket != null && dtPacket.Rows.Count > 0)
            {
                ddlPacket.DisplayMember = "Packet";
                ddlPacket.ValueMember = "Packet";
                ddlPacket.DataSource = dtPacket;
                ddlPacket.Enabled = true;
            }
        }


        private void GetDocModule()
        {
            MBRAuthDAL objDALMBR = new MBRAuthDAL();

            int DocID = mDocID;

            dtDocModule = objDALMBR.GetDocumentModule(DocID);
            if (dtDocModule != null && dtDocModule.Rows.Count > 0)
            {
                dgvModuleSelected.DataSource = dtDocModule;
                dgvModuleSelected.AutoGenerateColumns = false;

                dgvModuleSelected.Columns[0].Name = "ModuleBaseName";
                dgvModuleSelected.Columns[0].DataPropertyName = "ModuleBaseName";
                dgvModuleSelected.Columns[0].HeaderText = "ModuleName";
                dgvModuleSelected.Columns[1].Name = "ModuleTitle";
                dgvModuleSelected.Columns[1].DataPropertyName = "ModuleTitle";
                dgvModuleSelected.Columns[1].HeaderText = "ModuleTitle";
                dgvModuleSelected.Columns[2].Name = "ModuleRepetitions";
                dgvModuleSelected.Columns[2].DataPropertyName = "ModuleRepetitions";
                dgvModuleSelected.Columns[2].HeaderText = "Reps";
                dgvModuleSelected.Columns[3].Name = "ModuleUsage";
                dgvModuleSelected.Columns[3].DataPropertyName = "ModuleUsage";
                dgvModuleSelected.Columns[3].HeaderText = "Usage";
                dgvModuleSelected.Columns[4].Name = "Packet";
                dgvModuleSelected.Columns[4].DataPropertyName = "Packet";
                dgvModuleSelected.Columns[4].HeaderText = "Packets";
                dgvModuleSelected.Columns[5].Name = "ModuleTableRows";
                dgvModuleSelected.Columns[5].DataPropertyName = "ModuleTableRows";
                dgvModuleSelected.Columns[5].HeaderText = "TableRows";
                dgvModuleSelected.Columns[6].Name = "DocModuleID";
                dgvModuleSelected.Columns[6].DataPropertyName = "DocModuleID";
                dgvModuleSelected.Columns[6].HeaderText = "DocModuleID";
                dgvModuleSelected.Columns[7].Name = "DocumentTitle";
                dgvModuleSelected.Columns[7].DataPropertyName = "DocumentTitle";
                dgvModuleSelected.Columns[7].HeaderText = "DocumentTitle";
                dgvModuleSelected.Columns[8].Name = "ModuleComment";
                dgvModuleSelected.Columns[8].DataPropertyName = "ModuleComment";
                dgvModuleSelected.Columns[8].HeaderText = "ModuleComment";
                dgvModuleSelected.Columns[9].Name = "ModuleSequence";
                dgvModuleSelected.Columns[9].DataPropertyName = "ModuleSequence";
                dgvModuleSelected.Columns[9].HeaderText = "ModuleSequence";

                dgvModuleSelected.Columns[0].Width = 90;
                dgvModuleSelected.Columns[1].Width = 90;
                dgvModuleSelected.Columns[2].Width = 50;
                dgvModuleSelected.Columns[3].Width = 50;
                dgvModuleSelected.Columns[4].Width = 90;
                dgvModuleSelected.Columns[5].Width = 90;

                dgvModuleSelected.Columns[6].Visible = false;
                dgvModuleSelected.Columns[7].Visible = false;
                dgvModuleSelected.Columns[8].Visible = false;
                dgvModuleSelected.Columns[9].Visible = false;

                /*dgvModuleSelected.ColumnHeadersVisible = false;
                dgvModuleSelected.RowHeadersVisible = false;*/
            }
            else
            {

                MessageBox.Show("There are no properties available for this DocID");
            }

        }

        protected void ddlModuleAll_SelectedIndexChanged(object sender, EventArgs e)
        {

            int i = ddlModuleAll.SelectedIndex;
            SetToolTip(ddlModuleAll, "ddlModuleAll: OpUnitID=" +i+ " Module");

            if (Convert.ToInt32(dtModuleAll.Rows[i]["ModuleTitleCount"]) > 0)
            {
                lblModuleTitle.Visible = true;
                txtModuleTitle.Visible = true;

            }
            else
            {
                lblModuleTitle.Visible = false;
                txtModuleTitle.Visible = false;
            }
            if (Convert.ToInt32(dtModuleAll.Rows[i]["ModuleTableRowsCount"]) > 0)
            {
                object v = dtModuleAll.Rows[i]["ModuleTableRowsCount"];
                astrModuleTableRows = new string[(int)v];
                lblModuleTableRows.Visible = true;
                SpnUpDown.Visible = true;
                chkTableRows.Visible = true;
                chkTableRows.Enabled = true;
                SpnUpDown.Minimum = 1;
                //SpnUpDown.Maximum = Convert.ToInt32(dtModuleAll.Rows[i]["ModuleTableRowsCount"]);
                SpnUpDown.Value = Convert.ToInt32(dtModuleAll.Rows[i]["ModuleTableRowsCount"]); ;
                lblModuleTableRows.Text = "Table Rows 1 of 1";// + SpnUpDown.Maximum;
                lblModuleTableRows.Visible = true;

                //if (Convert.ToInt32(dtModuleAll.Rows[i]["ModuleTableRowsCount"]) > 1)
                //{
                //    SpnUpDown.Maximum = Convert.ToInt32(dtModuleAll.Rows[i]["ModuleTableRowsCount"]);
                //    blnSpnReachedMax = false;
                //}
                //else
                //{
                //    SpnUpDown.Maximum = 1;
                //    blnSpnReachedMax = false;
                //}
            }
            else
            {
                chkTableRows.Visible = false;
                lblModuleTableRows.Visible = false;
                //SpnUpDown.Value = 1;
                SpnUpDown.Visible = false;
            }

            strModule = ddlModuleAll.SelectedItem.ToString();
        }

        protected void ddlPacket_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = ddlPacket.SelectedIndex;
            SetToolTip(ddlPacket, "ddlPacket: DocID=" + mDocID + " Packet");
            if (index == -1)
            {
                MessageBox.Show("Please select a Packet.");
                ddlPacket.Focus();
                return;
            }
            else
            {
                strPacket = ddlPacket.SelectedValue.ToString();
            }
        }

        private void SpnUpDown_ValueChanged(object sender, EventArgs e)
        {
            chkTableRows.Checked = false;
        }


        private void btnModuleAdd_Click(object sender, EventArgs e)
        {
            int intListIndex;
            IsNew = true;
            mblnUserMadeModuleChange = true;


            if (btnModuleAdd.Text == "Reset")
            {
                ModuleReset();
                return;
            }



            if (!Module_Validation())
            {
                return;
            }

            ModuleReset();
            btnFinish.Enabled = true;
        }

        private void btnModuleUpdate_Click(object sender, EventArgs e)
        {
            IsUpdate = true;
            mblnUserMadeModuleChange = true;

            if (!Module_Validation())
            {
                return;
            }

            ModuleReset();


        }

        private void btnModuleDelete_Click(object sender, EventArgs e)
        {
            mblnUserMadeModuleChange = true;
            try
            {
                MBRAuthDAL objDALMBR = new MBRAuthDAL();
                int selectedModuleIndex = dgvModuleSelected.CurrentCell.RowIndex;
                //int index = dgvModuleSelected.CurrentCell.RowIndex;
                //if (index < 0)
                if (selectedModuleIndex < 0)
                {
                    MessageBox.Show("An item must be selected in order to delete it.");
                }
                else
                {
                    DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete this information ? ", "Confirm Deletion", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                       // dgvModuleSelected.Rows.RemoveAt(dgvModuleSelected.SelectedRows[0].Index);
                        dtDocModule.Rows.RemoveAt(dgvModuleSelected.SelectedRows[0].Index); // Fix for move up/down after deletion
                        dgvModuleSelected.Refresh(); 
                        //int ModuleSequence = Convert.ToInt32(dgvModuleSelected.SelectedRows[0].Cells["ModuleSequence"].Value);
                        //int ModuleSequence = Convert.ToInt32(dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleSequence"].Value);
                        //objDALMBR.DeleteModule(mDocID, ModuleSequence);
                        //GetDocModule();

                    }
                    else if (dialogResult == DialogResult.No)
                    {

                        return;
                    }

                }
                btnFinish.Enabled = true;
                ModuleReset();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "btnModuleDelete_Click");
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "btnModuleDelete_Click");
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private bool Module_Validation()
        {
            int intModTblRwsArrayCount;
            int intModTblRwsArrayScan;
            bool ModuleValid = true;
            IsSuccess = true;
            string strUsage = string.Empty;
            if (dgvModuleSelected.Columns.Count < 9)
            {
                MessageBox.Show("PROGRAM ERROR: lstModuleSelected.ColumnCount=" + dgvModuleSelected.Columns.Count + ", expecting " + 9, MyCompName, MessageBoxButtons.OK);
                return ModuleValid;
            }


            if (strModule == string.Empty)
            {
                MessageBox.Show("Please select a Module.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                IsSuccess = false;
                ddlModuleAll.Focus();
                return ModuleValid;
            }
            else if(strModule == "--Select--")
            {
                MessageBox.Show("Please select a Module.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                IsSuccess = false;
                ddlModuleAll.Focus();
                return ModuleValid;
            }
            else
            {
                strModule = strModule.Substring(strModule.LastIndexOf("-") + 1);
            }

            if (Convert.ToInt32(dtModuleAll.Rows[ddlModuleAll.SelectedIndex]["ModuleTitleCount"]) > 0)
            {
                if (txtModuleTitle.Text.Trim() == "")

                {

                    MessageBox.Show("Please enter a TITLE for Module.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    IsSuccess = false;
                    txtModuleTitle.Focus();
                    return ModuleValid;
                }
            }
            else
            {
                txtModuleTitle.Text = "N/A";
            }

            strPacket = ddlPacket.Text;
            DataTable pktDataTable = (DataTable)ddlPacket.DataSource;
            bool blnIsDuplicate = pktDataTable.AsEnumerable().Any(a => string.Equals(a[0].ToString(), strPacket, StringComparison.OrdinalIgnoreCase));

            if (!blnIsDuplicate)
            {
                DataRow dataRow = pktDataTable.NewRow();
                dataRow[0] = strPacket;
                pktDataTable.Rows.Add(dataRow);
            }
            if (Convert.ToInt32(dtModuleAll.Rows[ddlModuleAll.SelectedIndex]["ModuleTableRowsCount"]) > 0)
            {

                if (SpnUpDown.Value < 1 && chkTableRows.Enabled == false && !chkTableRows.Checked)
                {
                    MessageBox.Show("TABLE ROWS has to be set greater than 0 or enable 'P' for PalletCount.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }


            }
            if (IsSuccess)
            {
                if (txtRepetitions.Text == string.Empty)
                {
                    txtRepetitions.Text = "1";


                }
                if (rdoBatchRecord.Checked) //Soundhar 010825
                {
                    strUsage = "B";
                }

                else if (rdoInfoOnly.Checked)  //Soundhar 010825
                {
                    strUsage = "I";
                }

                if (chkTableRows.Enabled && chkTableRows.Visible && chkTableRows.Checked)
                {
                    TableRows = "P";
                }
                else if (SpnUpDown.Value > 0 && SpnUpDown.Visible && !chkTableRows.Checked)
                {
                    TableRows = Convert.ToString(SpnUpDown.Value);
                }
                else
                {
                    TableRows = "";
                }
                if (IsNew)
                {
                    if (dtDocModule.Columns.Count != 10)
                    {
                        MessageBox.Show("EXPECTING 10 NOT " + dtDocModule.Columns.Count);
                    }

                    DataRow dr = dtDocModule.NewRow();
                    dr["DocModuleID"] = -1;
                    dr["ModuleSequence"] = (dtDocModule.Rows.Count) + 1;
                    dr["ModuleBaseName"] = dtModuleAll.Rows[ddlModuleAll.SelectedIndex]["DocumentBaseName"].ToString();
                    dr["ModuleUsage"] = strUsage;
                    dr["ModuleRepetitions"] = Convert.ToInt32(txtRepetitions.Text);
                    dr["Packet"] = strPacket;
                    dr["ModuleTitle"] = txtModuleTitle.Text;
                    dr["ModuleTableRows"] = TableRows;
                    dr["DocumentTitle"] = dtModuleAll.Rows[ddlModuleAll.SelectedIndex]["DocumentTitle"].ToString();
                    dtDocModule.Rows.Add(dr);
                    dgvModuleSelected.DataSource = dtDocModule;
                    IsNew = false;

                }
                if (IsUpdate)
                {
                    int i = dgvModuleSelected.SelectedRows[0].Index; //Soundhar 010825
                    dtDocModule.Rows[i]["ModuleBaseName"] = dtModuleAll.Rows[ddlModuleAll.SelectedIndex]["DocumentBaseName"].ToString();
                    dtDocModule.Rows[i]["ModuleRepetitions"] = Convert.ToInt32(txtRepetitions.Text);
                    dtDocModule.Rows[i]["Packet"] = strPacket;
                    dtDocModule.Rows[i]["ModuleUsage"] = strUsage;
                    dtDocModule.Rows[i]["ModuleTitle"] = txtModuleTitle.Text;
                    dtDocModule.Rows[i]["ModuleTableRows"] = TableRows;
                    dgvModuleSelected.DataSource = dtDocModule;
                    IsUpdate = false;
                }

                btnModuleAdd.Text = "Reset";

                /*if (IsSave)
                {
                    int intMain;
                    MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
                    string strModuleComment = "";
                    for (intMain = 0; intMain <= dgvModuleSelected.Rows.Count - 1; intMain++)
                    {
                        objMBRAuthDAL.InsertUpdateDocModule(DocID: mDocID,
                                             ModuleBaseName: dgvModuleSelected.Rows[intMain].Cells[qdmColModuleBaseName].Value.ToString(),
                                             ModuleComment: strModuleComment,
                                             ModuleSequence: intMain + 1,
                                             ModuleRepetitions: Convert.ToInt32(dgvModuleSelected.Rows[intMain].Cells[qdmColModuleRepetitions].Value),
                                             Packet: dgvModuleSelected.Rows[intMain].Cells[qdmColModulePacket].Value.ToString(),
                                             ModuleTitle: dgvModuleSelected.Rows[intMain].Cells[qdmColModuleModuleTitle].Value.ToString(),
                                             ModuleTableRows: dgvModuleSelected.Rows[intMain].Cells[qdmColModuleModuleTableRows].Value.ToString(),
                                             ModuleUsage: dgvModuleSelected.Rows[intMain].Cells[qdmColModuleUsage].Value.ToString());
                    }                    
                }*/
            }

            else
            {
                MessageBox.Show("The data entered is not valid");
            }
            return ModuleValid;
        }



        private void dgvModuleSelected_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            int selectedModuleIndex = dgvModuleSelected.CurrentCell.RowIndex;
            dgvModuleSelected.Rows[selectedModuleIndex].Selected = true;
            if (selectedModuleIndex < 0)
            {
                return;
            }

            
            string selectedModuleValue = dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleBaseName"].Value.ToString().Trim() + " - " + dgvModuleSelected.Rows[selectedModuleIndex].Cells["DocumentTitle"].Value.ToString().Trim();
            ddlModuleAll.Enabled = false;
            //ddlModuleAll.SelectedItem = selectedModuleValue;
            for (int i = 0; i <= dtModuleAll.Rows.Count - 1; i++)
            {
                if (dtModuleAll.Rows[i]["DocumentBaseName"].ToString() + " - " + dtModuleAll.Rows[i]["DocumentTitle"].ToString().Trim() == selectedModuleValue)
                {
                    //ddlModuleAll.SelectedIndex = i;
                    ddlModuleAll.SelectedIndex = i; //Soundhar
                    break;
                }
            }

            dgvModuleSelected.ReadOnly = true;
            txtRepetitions.Text = Convert.ToString(dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleRepetitions"].Value);
            ddlPacket.Text = Convert.ToString(dgvModuleSelected.Rows[selectedModuleIndex].Cells["Packet"].Value);
            txtModuleTitle.Text = Convert.ToString(dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleTitle"].Value);

            if (dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleUsage"].Value.ToString() == "B")
                rdoBatchRecord.Checked = true;
            else if (dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleUsage"].Value.ToString() == "I")
                rdoInfoOnly.Checked = true;
            if ((dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleTableRows"].Value).ToString() == "P")
            {
                chkTableRows.Enabled = true;
                chkTableRows.Checked = true; //Soundhar
            }
            else
            {
                int number;
                if (int.TryParse((dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleTableRows"].Value).ToString(), out number))
                {
                    if(Convert.ToInt32(dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleTableRows"].Value) > 0) //060225
                    {
                        lblModuleTableRows.Visible = true;
                        SpnUpDown.Visible = true;
                        chkTableRows.Visible = true;
                        chkTableRows.Enabled = true;
                        SpnUpDown.Minimum = 1;
                        //SpnUpDown.Maximum = Convert.ToInt32(dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleTableRows"].Value);
                        SpnUpDown.Value = Convert.ToInt32(dgvModuleSelected.Rows[selectedModuleIndex].Cells["ModuleTableRows"].Value);
                        lblModuleTableRows.Text = "Table Rows 1 of 1";// + SpnUpDown.Maximum;
                        lblModuleTableRows.Visible = true;
                    }                    
                }

            }

            btnModuleAdd.Text = "Reset";
            dgvModuleSelected.ReadOnly = true;
            btnModuleUpdate.Enabled = true;
            btnModuleDelete.Enabled = true;
            btnMoveUp.Enabled = true;
            btnMoveDown.Enabled = true;


        }



        private void ModuleReset()
        {
            dgvModuleSelected.ClearSelection();
            if (ddlModuleAll.SelectedIndex > 0)
            {
                ddlModuleAll.SelectedIndex = 0;
            }
            ddlModuleAll.Enabled = true;
            txtRepetitions.Text = "1";
            ddlPacket.Text = "Master";
            SpnUpDown.Value = 1;
            txtModuleTitle.Text = "";
            rdoBatchRecord.Enabled = true;
            rdoInfoOnly.Enabled = true;
            ddlModuleAll.Select();
            btnModuleAdd.Text = "Add";
            dgvModuleSelected.Enabled = true;
            btnModuleUpdate.Enabled = false;
            btnModuleDelete.Enabled = false;
            btnMoveUp.Enabled = false;
            btnMoveDown.Enabled = false;
        }

        private void btnMoveDown_Click(object sender, EventArgs e)
        {
            mblnUserMadeModuleChange = true;
            MoveRow(1);

        }

        private void btnMoveUp_Click(object sender, EventArgs e)
        {
            mblnUserMadeModuleChange = true;
            MoveRow(-1);
        }

        private void MoveRow(int direction)
        {
            if (dgvModuleSelected.Rows.Count > 0)
            {
                int selectedRowIndex = dgvModuleSelected.SelectedRows[0].Index;

                //int selectedRowIndex = dgvModuleSelected.CurrentCell.RowIndex;

                //new row index calculate
                int newRowIndex = selectedRowIndex + direction;
                int currentSrollPosition = dgvModuleSelected.FirstDisplayedScrollingRowIndex;
                if (selectedRowIndex == 0 && direction == -1)
                {
                    MessageBox.Show("Cannot Move the FIRST item Up", "DataGrid - MoveUp", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                if (selectedRowIndex == dtDocModule.Rows.Count - 1 && direction == 1)
                {
                    MessageBox.Show("Cannot move LAST item DOWN", "DataGrid - MoveDown", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                if (newRowIndex >= 0 && newRowIndex < dtDocModule.Rows.Count)
                {
                    //copy the current row to new position
                    DataRow tempRow = dtDocModule.NewRow();
                    tempRow.ItemArray = dtDocModule.Rows[selectedRowIndex].ItemArray;
                    dtDocModule.Rows[selectedRowIndex].ItemArray = dtDocModule.Rows[newRowIndex].ItemArray;
                    dtDocModule.Rows[newRowIndex].ItemArray = tempRow.ItemArray;

                    //Reselect the row
                    dgvModuleSelected.ClearSelection();
                    dgvModuleSelected.Rows[newRowIndex].Selected = true;

                }
                //Restore the scroll position
                if (currentSrollPosition >= 0 && currentSrollPosition <= dgvModuleSelected.Rows.Count)
                {
                    dgvModuleSelected.FirstDisplayedScrollingRowIndex = currentSrollPosition;
                }

                if (newRowIndex >= 0 && newRowIndex < dgvModuleSelected.Rows.Count)
                {
                    dgvModuleSelected.FirstDisplayedScrollingRowIndex = Math.Max(0, newRowIndex);
                }
            }
        }


        /*private void MoveRow(int direction)
        {
            if (dgvModuleSelected.Rows.Count > 0)
            {
                //int selectedRowIndex = dgvModuleSelected.SelectedRows[0].Index;
                
                int selectedRowIndex = dgvModuleSelected.CurrentCell.RowIndex;

                //new row index calculate
                int newRowIndex = selectedRowIndex + direction;
                if (selectedRowIndex == 0 && direction == -1)
                {
                    MessageBox.Show("Cannot Move the FIRST item Up", "DataGrid - MoveUp", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                if (selectedRowIndex == dtDocModule.Rows.Count - 1 && direction == 1)
                {
                    MessageBox.Show("Cannot move LAST item DOWN", "DataGrid - MoveDown", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                if (newRowIndex >= 0 && newRowIndex < dtDocModule.Rows.Count)
                {
                    //copy the current row to new position
                    DataRow tempRow = dtDocModule.NewRow();
                    tempRow.ItemArray = dtDocModule.Rows[selectedRowIndex].ItemArray;
                    dtDocModule.Rows[selectedRowIndex].ItemArray = dtDocModule.Rows[newRowIndex].ItemArray;
                    dtDocModule.Rows[newRowIndex].ItemArray = tempRow.ItemArray;

                    //Reselect the row
                    dgvModuleSelected.ClearSelection();
                    dgvModuleSelected.Rows[newRowIndex].Selected = true;

                }
            }
        }*/

        private bool Valid_Entries()
        {
            bool Valid_Entries = false;
            mintParmsValidated = 0;

            if (dgvModuleSelected.Rows.Count == 0)
            {
                DialogResult dialogResult = MessageBox.Show("Do you want to SAVE with NO MODULES? ", "Confirm No Modules", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {

                }
                else if (dialogResult == DialogResult.No)
                {
                    MessageBox.Show("You must select at least one Module.", "Module Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    ddlModuleAll.Select();
                }

            }
            Valid_Entries = true;
            mintParmsValidated = cintParmsValidated;

            return Valid_Entries;
        }


        //private void spnTableNum_SpinUp()
        //{
        //    if (SpnUpDown.Value <= Convert.ToInt32(dtModuleAll.Rows[ddlModuleAll.SelectedIndex]["ModuleTableRowsCount"]))
        //    {


        //        if (SpnUpDown.Value == Convert.ToInt32(dtModuleAll.Rows[ddlModuleAll.SelectedIndex]["ModuleTableRowsCount"]))
        //        {
        //            blnSpnReachedMax = true;
        //            if (SpnUpDown.Maximum > 1)
        //            {
        //                lblModuleTableRows.Text = "Table Rows " + SpnUpDown.Value + " of " + SpnUpDown.Maximum;
        //                int value = (int)SpnUpDown.Value;
        //                rdoTableRows.Checked = Convert.ToBoolean(astrModuleTableRows[value]);
        //                if (SpnUpDown.Value == SpnUpDown.Maximum)
        //                    blnSpnReachedMax = true;
        //            }
        //        }
        //    }
        //}

        //private void spnTableNum_SpinDown()
        //{
        //    if (SpnUpDown.Value >= SpnUpDown.Minimum)
        //    {
        //        if (SpnUpDown.Value == SpnUpDown.Minimum)
        //            blnSpnReachedMin = true;
        //        else
        //            blnSpnReachedMin = false;
        //        lblModuleTableRows.Text = "Table Rows " + SpnUpDown.Value + " of " + SpnUpDown.Maximum;
        //        int value = (int)SpnUpDown.Value;
        //        rdoTableRows.Enabled = Convert.ToBoolean(astrModuleTableRows[value]);
        //        if (SpnUpDown.Value == SpnUpDown.Minimum)
        //            blnSpnReachedMin = true;
        //    }
        //}

        private void Document_Protect()
        {
            if (!TryResolveActiveDocumentForForm())
            {
                MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.Close();
                return;
            }

            if (mblnCalledWizard)
            {

            }
            CommonService.DocumentProtection_On(activeDocument, mstrProtectionPassword);

        }

        private void Document_UnProtect()
        {
            if (!TryResolveActiveDocumentForForm())
            {
                MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.Close();
                return;
            }

            if (mblnCalledWizard)
            {

            }
            CommonService.DocumentProtection_Off(activeDocument, mstrProtectionPassword);

        }

        private void Setup_Document()
        {
            int intModule;
            //            int intTableRow = -1;
            int intTableRow = 0;
            Word.Table tblModule = null;
            int intModuleCount;
            if (dgvModuleSelected.Rows.Count > 0)
            {
                intModuleCount = dgvModuleSelected.Rows.Count;
            }
            else
            {
                intModuleCount = 1;
            }
            if (!TryResolveActiveDocumentForForm())
            {
                MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.Close();
                return;
            }
            //CommonService.AutoText_Insert("ModuleList", ref intModuleCount, intTableRow, tableAutoText: tblModule, wordApp: activeDocument.Application);
            CommonService.AutoText_Insert("ModuleList", ref intTableRow, intModuleCount, tableAutoText: tblModule, wordApp: ExecuteWordComWithRetry(() => activeDocument.Application));
            {
                tblModule = CommonService.Table_Find_ModuleMBR(ExecuteWordComWithRetry(() => activeDocument));
                tblModule.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            }

            for (intModule = 0; intModule <= dgvModuleSelected.Rows.Count - 1; intModule++)
            {
                //tblModule.Cell(intTableRow, 1).Range.Text = dgvModuleSelected.Columns[qdmColModuleBaseName].ToString();
                tblModule.Cell(intTableRow, 1).Range.Text = dgvModuleSelected.Rows[intModule].Cells["ModuleBaseName"].Value.ToString();
                if (dgvModuleSelected.Rows[intModule].Cells["ModuleUsage"].Value.ToString() == "B")
                {
                    tblModule.Cell(intTableRow, 2).Range.Text = "Batch";
                }

                else if (dgvModuleSelected.Rows[intModule].Cells["ModuleUsage"].Value.ToString() == "I")
                {
                    tblModule.Cell(intTableRow, 2).Range.Text = "Info";
                }
                string strText;
                strText = dgvModuleSelected.Rows[intModule].Cells["DocumentTitle"].Value.ToString();
                if (dgvModuleSelected.Rows[intModule].Cells["ModuleTitle"].Value.ToString() != "N/A")
                {
                    strText = strText + "\r\n" + "Module Title: " + dgvModuleSelected.Rows[intModule].Cells["ModuleTitle"].Value.ToString();
                }

                if (dgvModuleSelected.Rows[intModule].Cells["ModuleTableRows"].Value.ToString() != "N /A" & dgvModuleSelected.Rows[intModule].Cells["ModuleTableRows"].Value.ToString() != "")
                    strText = strText + "\r\n" + "Module Table Rows: " + dgvModuleSelected.Rows[intModule].Cells["ModuleTableRows"].Value.ToString();
                tblModule.Cell(intTableRow, 3).Range.Text = strText;
                tblModule.Cell(intTableRow, 4).Range.Text = dgvModuleSelected.Rows[intModule].Cells["Packet"].Value.ToString();
                tblModule.Cell(intTableRow, 5).Range.Text = dgvModuleSelected.Rows[intModule].Cells["ModuleRepetitions"].Value.ToString();
                intTableRow = intTableRow + 1;

                /*else if (dgvModuleSelected.Rows[intModule].Cells["ModuleUsage"].Value.ToString() == "I")
                {
                    tblModule.Cell(intTableRow, 2).Range.Text = "Info";
                    string strText;
                    strText = dgvModuleSelected.Rows[intModule].Cells["ModuleTitle"].Value.ToString();
                    if (dgvModuleSelected.Rows[intModule].Cells["ModuleTitle"].Value.ToString() != "N/A")
                    {
                        strText = strText + "\r\n" + "Module Title: " + dgvModuleSelected.Rows[intModule].Cells["ModuleTitle"].Value.ToString();
                    }

                    if (dgvModuleSelected.Rows[intModule].Cells["ModuleTableRows"].Value.ToString() != "N /A" & dgvModuleSelected.Rows[intModule].Cells["ModuleModuleTableRows"].Value.ToString() != "")
                        strText = strText + "\r\n" + "Module Table Rows: " + dgvModuleSelected.Rows[intModule].Cells["ModuleTableRows"].Value.ToString();
                    tblModule.Cell(intTableRow, 3).Range.Text = strText;
                    tblModule.Cell(intTableRow, 4).Range.Text = dgvModuleSelected.Rows[intModule].Cells["ModulePacket"].Value.ToString();
                    tblModule.Cell(intTableRow, 5).Range.Text = dgvModuleSelected.Rows[intModule].Cells["ModuleRepetitions"].Value.ToString();
                    intTableRow = intTableRow + 1;
                }*/
            }
        }

        private void Setup_Database()
        {
            int intMain;
            MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
            string strModuleComment = "";
            //if (Transaction_Active == "")
            //{
            //    Transaction_Begin(mstrDocName);
            //}
            blnDBUpdateSuccessful = false;
            try
            {


                for (intMain = 0; intMain <= dgvModuleSelected.Rows.Count - 1; intMain++)
                {
                    objMBRAuthDAL.InsertUpdateDocModule(DocID: mDocID,
                                         ModuleBaseName: dgvModuleSelected.Rows[intMain].Cells[qdmColModuleBaseName].Value.ToString(),
                                         ModuleComment: strModuleComment,
                                         ModuleSequence: intMain + 1,
                                         ModuleRepetitions: Convert.ToInt32(dgvModuleSelected.Rows[intMain].Cells[qdmColModuleRepetitions].Value),
                                         Packet: dgvModuleSelected.Rows[intMain].Cells[qdmColModulePacket].Value.ToString(),
                                         ModuleTitle: dgvModuleSelected.Rows[intMain].Cells[qdmColModuleModuleTitle].Value.ToString(),
                                         ModuleTableRows: dgvModuleSelected.Rows[intMain].Cells[qdmColModuleModuleTableRows].Value.ToString(),
                                         ModuleUsage: dgvModuleSelected.Rows[intMain].Cells[qdmColModuleUsage].Value.ToString());
                }
                int ModuleSequence = dgvModuleSelected.Rows.Count;
                objMBRAuthDAL.DeleteModule(mDocID, ModuleSequence);
                blnDBUpdateSuccessful = true;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "LoadDocumentProperty");
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "LoadDocumentProperty");
                }
            }
        }

        private void btnFinish_Click(object sender, EventArgs e)
        {

            Cursor.Current = Cursors.WaitCursor;
            string MyProcName = "btnFinish";
            Word.Table tbl;
            bool blnProtect;
            bool btnWordUpdate = false;
            try
            {
                //activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorWait;

                if (!Valid_Entries())
                {
                    return;
                }
                if (mblnCalledWizard)
                {
                    UserForm__SHUTDOWN(qReturn.qReturn_Finish);
                    return;
                }


                this.Left = Application.OpenForms[0].Width - tabModule.Left;
                Setup_Database();
                if (!TryResolveActiveDocumentForForm())
                {
                    btnWordUpdate = false;
                    //throw new InvalidOperationException("Unable to access the target document, Click on Refresh menu to update the document.");
                    if (blnDBUpdateSuccessful)
                    {

                        MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        this.Close();
                        return;
                    }
                    else
                    {
                        MessageBox.Show("Data save operation failed. Try again.", "Failed to Save", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                if (blnDBUpdateSuccessful != true)
                {
                    MessageBox.Show("Data save operation failed. Try again.", "Failed to Save", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (ExecuteWordComWithRetry(() => activeDocument.ProtectionType == Word.WdProtectionType.wdNoProtection))
                {
                    blnProtect = false;
                }
                else
                {
                    blnProtect = true;
                    Document_UnProtect();
                }
                if (!TryResolveActiveDocumentForForm())
                {
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        this.Close();
                        return;
                }
                tbl = CommonService.Table_Find(ExecuteWordComWithRetry(() => activeDocument), "Module List");


                if (tbl == null)
                {
                    MessageBox.Show("Existing Module-Table NOT FOUND - Skipping UPDATE", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if (!TryResolveActiveDocumentForForm())
                    {
                        MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        this.Close();
                        return;
                    }
                    CommonService.Table_Delete(ptblToDelete: tbl, pdocCurrent: ExecuteWordComWithRetry(() => activeDocument), pstrProtectionPassword: mstrProtectionPassword);
                    Setup_Document();
                }

                if (blnProtect)
                {
                    Document_Protect();
                }
                if (!TryResolveActiveDocumentForForm())
                {
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }
                SaveManager.Document_Save_FileAndDatabase(pdocCurrent: ExecuteWordComWithRetry(() => activeDocument), docpasswordProtect: mstrProtectionPassword, pblnRestoreCommandBar: false);//passed pass

                if (!TryResolveActiveDocumentForForm())
                {
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }
                CommonService.DocumentProtection_Assure(pdocProtectedDocument: ExecuteWordComWithRetry(() => activeDocument), pstrProtectionPassword: mstrProtectionPassword);
                btnWordUpdate = true;
                //activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;
                if (!TryResolveActiveDocumentForForm())
                {
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }
                Cursor.Current = Cursors.Default;
                ExecuteWordComWithRetry(() => activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal); // To suppress RPC
                this.Close();
                //return;

            }
            catch (Exception ex)
            {
             
                if (TryResolveActiveDocumentForForm() && ExecuteWordComWithRetry(() => activeDocument)?.Application != null)
                {
                    ExecuteWordComWithRetry(() => activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal);
                }
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
               // Cursor = Cursors.Default;
                //performancetuning
                StopWordKeepAliveTimer();
                if (blnDBUpdateSuccessful)
                {

                    CommonService.CleanUpBackgroundWordProcesses();
                }
            }
        }

        //FormCommon Methods
        public void DataGridView_From_Array(DataGridView dgvTempGrid, string[,] pstrArray, DataGridView dgvTempGridAll = null, int pintKeyColumn = -1)
        {
            const string MyProcName = "DataGridView_From_Array";
            int intRow;
            int intRowAll;
            int intCol;
            int intColumnCount;
            int intRowCount;
            try
            {
                if (dgvTempGridAll == null & pintKeyColumn != -1)
                {
                    //Todo Information.Err.Raise(41611, MyCompName + "." + MyProcName, "Cannot specify KeyColumn=" + pintKeyColumn + " without 'All' ListBox");
                    var ex = new Exception();
                    ex.Data["errNum"] = "41611";
                    ex.Data["errDesc"] = "Cannot specify KeyColumn=" + pintKeyColumn + " without 'All' ListBox";
                    throw ex;
                    return;
                }

                if (!(dgvTempGridAll == null) & pintKeyColumn == -1)
                {
                    //Todo Information.Err.Raise(41612, MyCompName + "." + MyProcName, "Cannot specify ListBoxAll=" + plstListBoxAll.Name + " with KeyColumn=" + pintKeyColumn);
                    var ex = new Exception();
                    ex.Data["errNum"] = "41612";
                    ex.Data["errDesc"] = "Cannot specify ListBoxAll=" + dgvTempGrid.Name + " with KeyColumn=" + pintKeyColumn;
                    throw ex;
                    return;
                }

                dgvTempGrid.ClearSelection();

                //On Error GoTo Array_Is_Undefined


                intRowCount = pstrArray.GetUpperBound(1);
                intColumnCount = pstrArray.GetUpperBound(2) + 1;

                if (intColumnCount < 1)
                {
                    MessageBox.Show("Cannot process Array with Ubound(2)=" + intColumnCount, MyCompName + "." + MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (dgvTempGridAll == null)
                {
                    if (pintKeyColumn != -1)
                    {
                        MessageBox.Show("CANNOT SPECIFY KeyColumn=" + pintKeyColumn + " WITHOUT SPECIFYING 'GridAll'", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

                if (pintKeyColumn != -1)
                {
                    if (dgvTempGridAll == null)
                    {
                        MessageBox.Show("CANNOT SPECIFY GridAll=" + dgvTempGridAll.Name + " WITHOUT SPECIFYING 'KeyColumn'", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

                dgvTempGrid.ClearSelection();
                dgvTempGrid.ColumnCount = intColumnCount;
                for (intRow = 0; intRow <= intRowCount - 1; intRow++)
                {
                    dgvTempGrid.Rows.Add();
                    for (intCol = 0; intCol <= intColumnCount - 1; intCol++)
                    {
                        dgvTempGrid.Rows[intRow].Cells[intCol].Value = pstrArray[intRow + 1, intCol];
                    }
                }

                if (pintKeyColumn != -1)
                {
                    DataGridView_Exclude(dgvMaster: dgvTempGrid, dgvToExclude: dgvTempGridAll, pintKeyColumn: pintKeyColumn);
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                pstrArray = new string[0, 0];
                intRowCount = 0;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "LoadDocumentProperty");
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "LoadDocumentProperty");
                }
            }
        }

        public void DataGridView_To_Array(DataGridView dgvTempGrid, string[,] pstrArray)
        {
            const string MyProcName = "ListBox_To_Array";
            int intRow;
            int intCol;
            int intColumnCount;
            int intRowCount;
            intRowCount = dgvTempGrid.Rows.Count;
            intColumnCount = dgvTempGrid.Columns.Count;
            if (intColumnCount < 1)
            {
                MessageBox.Show("Cannot process Grid '" + dgvTempGrid.Name + "' with " + intColumnCount + " Columns", MyCompName + "." + MyProcName, MessageBoxButtons.OK);
                return;
            }

            pstrArray = new string[intRowCount + 1, intColumnCount - 1 + 1];
            for (intRow = 0; intRow <= intRowCount - 1; intRow++)
            {
                for (intCol = 0; intCol <= intColumnCount - 1; intCol++)
                {
                    pstrArray[intRow + 1, intCol] = dgvTempGrid.Rows[intRow].Cells[intCol].Value.ToString();
                }
            }
        }

        public void DataGridView_Exclude(DataGridView dgvMaster, DataGridView dgvToExclude, int pintKeyColumn)
        {
            int intRow;
            int intRowAll;
            for (intRow = 0; intRow <= dgvMaster.Rows.Count - 1; intRow++)
            {
                for (intRowAll = 0; intRowAll <= dgvToExclude.Rows.Count - 1; intRowAll++)
                {
                    if (dgvToExclude.Rows[intRowAll].Cells[pintKeyColumn].Value == dgvMaster.Rows[intRow].Cells[pintKeyColumn].Value)
                    {

                        dgvToExclude.Rows.RemoveAt(intRowAll);
                        intRowAll = dgvToExclude.Rows.Count;
                    }
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            /*if (mqrqCalledRequest == CommonService.qRequest_Add)
            {
                UserForm__SHUTDOWN(qReturn.qReturn_Cancel);
                return;
            }
            YieldChanged();*/
            try
            { 
            if (mblnUserMadeModuleChange)
            {
                DialogResult result = MessageBox.Show(
                               "OK to DISCARD MODULE CHANGES?.",
                               "Confirm CANCEL",
                               MessageBoxButtons.YesNo,
                              MessageBoxIcon.Question
                             );
                switch (result)
                {
                    case DialogResult.Yes:
                        break;
                    case DialogResult.No:
                        return;
                        break;
                    default:
                        MessageBox.Show("unexpected response");
                        return;
                        break;
                }

            }

            if (mblnCalledWizard)
            {
                UserForm__SHUTDOWN(qReturn.qReturn_Cancel);
            }
            else
            {
                this.Close();
            }
        }
       finally
            {
                CommonService.CleanUpBackgroundWordProcesses();
            }
}

        private void lblMap5_Click(object sender, EventArgs e)
        {
            btnFinish_Click(sender, e);
        }

        //Performancetuning
        private void WordKeepAliveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if(!TryResolveActiveDocumentForForm())
                {
                    return;
                }
                if (ExecuteWordComWithRetry(() => activeDocument) != null)//&& IsWordRunning())
                {
                    string wordVersion = ExecuteWordComWithRetry(() => activeDocument).Application.Version;
                }
            }
            catch (Exception ex)
            {

                //MessageBox.Show("Keep-alive failed: " + ex.Message);
            }
        }

        //Performancetuning
        private void StopWordKeepAliveTimer()
        {
            if (wordKeepAliveTimer != null)
            {
                wordKeepAliveTimer.Stop();
                wordKeepAliveTimer.Tick -= WordKeepAliveTimer_Tick;
                wordKeepAliveTimer.Dispose();
                wordKeepAliveTimer = null;
            }
        }

        public void SetTargetDocumentName(string documentName, string documentFullName = "")
        {
            if (!string.IsNullOrWhiteSpace(documentName))
            {
                mstrDocName = documentName;
            }

            if (!string.IsNullOrWhiteSpace(documentFullName))
            {
                mstrDocFullName = documentFullName;
            }
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
        // Resolves the Word document this form should operate on.
        // Preference order:
        // 1) Reuse current activeDocument if it is still valid and matches.
        // 2) Find an open document by full name/name using Word windows/documents.
        // 3) Fall back to the currently active Word document.
        // Returns false when no suitable document can be resolved.
        private bool TryResolveActiveDocumentForForm()
        {
            try
            {
                // Fast-path: keep the existing reference when it is still usable.
                if (activeDocument != null)
                {
                    string currentDocumentName = activeDocument.Name;
                    if (string.IsNullOrWhiteSpace(mstrDocName) ||
                        string.Equals(currentDocumentName, mstrDocName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrWhiteSpace(mstrDocName))
                        {
                            mstrDocName = currentDocumentName;
                        }

                        return true;
                    }
                }
            }
            catch
            {
                // COM proxy may be stale; force a fresh resolve path.
                activeDocument = null;
            }

            // Retry a few times to handle transient Word COM timing issues.
            for (int retry = 0; retry < 3; retry++)
            {
                try
                {
                    Word.Document resolvedDocument = null;
                    Word.Application wordApp = GetWordApplicationForForm();

                    //// First try from the document that launched/hosts this add-in.
                    //if (Globals.ThisDocument?.ActiveDocument != null)
                    //{
                    //    resolvedDocument = FindOpenDocument(Globals.ThisDocument.ActiveDocument.Application, mstrDocName, mstrDocFullName);
                    //}

                    // Then try from the global Word application object.
                    if (resolvedDocument == null)
                    {
                        resolvedDocument = FindOpenDocument(wordApp, mstrDocName, mstrDocFullName);
                    }

                    // If a specific target name was provided but not found, retry briefly.
                    if (resolvedDocument == null && !string.IsNullOrWhiteSpace(mstrDocName))
                    {
                        if (retry < 2)
                        {
                            System.Threading.Thread.Sleep(100);
                            continue;
                        }

                        activeDocument = null;
                        return false;
                    }

                    // Last fallback: use Word's current active document.
                    if (resolvedDocument == null)
                    {
                        // Word.Application wordApp = Globals.ThisDocument?.Application;//commented
                        if (wordApp != null)
                        {
                            resolvedDocument = wordApp.ActiveDocument;
                        }
                    }

                    if (resolvedDocument == null)
                    {
                        activeDocument = null;
                        return false;
                    }

                    // Save the resolved target for downstream operations.
                    activeDocument = resolvedDocument;

                    if (string.IsNullOrWhiteSpace(mstrDocName))
                    {
                        mstrDocName = activeDocument.Name;
                    }

                    return true;
                }
                catch (COMException)
                {
                    // Transient COM failures can occur while Word is switching documents.
                    activeDocument = null;
                    if (retry < 2)
                    {
                        System.Threading.Thread.Sleep(100);
                        continue;
                    }

                    return false;
                }
                catch
                {
                    activeDocument = null;
                    return false;
                }
            }

            activeDocument = null;
            return false;
        }

        private Word.Application GetWordApplicationForForm()
        {
            try
            {
                if (activeDocument != null)
                {
                    return activeDocument.Application;
                }
            }
            catch
            {
            }

            try
            {
                if (Globals.ThisDocument?.Application != null)
                {
                    return Globals.ThisDocument.Application;
                }
            }
            catch
            {
            }

            try
            {
                return Marshal.GetActiveObject("Word.Application") as Word.Application;
            }
            catch
            {
                return null;
            }
        }

        // Attempts to locate an already-open document using both open windows and the document list.
        private static Word.Document FindOpenDocument(Word.Application wordApp, string documentName, string documentFullName)
        {
            if (wordApp == null)
            {
                return null;
            }

            Word.Document windowDocument = FindOpenDocumentByWindows(wordApp, documentName, documentFullName);
            if (windowDocument != null)
            {
                return windowDocument;
            }

            return FindOpenDocumentByDocuments(wordApp, documentName, documentFullName);
        }

        // Scans open Word windows first, which typically reflects visible user context.
        private static Word.Document FindOpenDocumentByWindows(Word.Application wordApp, string documentName, string documentFullName)
        {
            for (int windowIndex = 1; windowIndex <= wordApp.Windows.Count; windowIndex++)
            {
                Word.Window openWindow = wordApp.Windows[windowIndex];
                Word.Document openDocument = openWindow.Document;

                if (DocumentMatches(openDocument, documentName, documentFullName))
                {
                    return openDocument;
                }
            }

            return null;
        }

        // Scans Word's documents collection as a secondary lookup path.
        private static Word.Document FindOpenDocumentByDocuments(Word.Application wordApp, string documentName, string documentFullName)
        {
            for (int documentIndex = 1; documentIndex <= wordApp.Documents.Count; documentIndex++)
            {
                Word.Document openDocument = wordApp.Documents[documentIndex];
                if (DocumentMatches(openDocument, documentName, documentFullName))
                {
                    return openDocument;
                }
            }

            return null;
        }

        // Matches by full path when available; otherwise falls back to document name.
        private static bool DocumentMatches(Word.Document openDocument, string documentName, string documentFullName)
        {
            if (openDocument == null)
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(documentFullName))
            {
                return string.Equals(openDocument.FullName, documentFullName, StringComparison.OrdinalIgnoreCase);
            }

            if (!string.IsNullOrWhiteSpace(documentName))
            {
                return string.Equals(openDocument.Name, documentName, StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }
    }

}
