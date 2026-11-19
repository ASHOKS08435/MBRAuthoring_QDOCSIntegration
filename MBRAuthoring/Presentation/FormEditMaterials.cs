using MBRAuthoring.Business;
using MBRAuthoring.DataAccess;
using MBRAuthoring.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MBRAuthoring.Presentation
{
    public partial class FormEditMaterials : FormBase
    {
        protected override Panel progressBarPanel => panel1;
        public IDocumentManager _documentManager = null;
        private Word.Document activeDocument = Globals.ThisDocument.ActiveDocument;
        Word.Range mrngHold;
        Word.Range mrngYield;

        const string MyCompName = "frmEditMaterial";

        string mstrDocName;
        bool mblnYieldActive;
        int mDocIDNegative;
        string DocIsModule;
        string DocHasModule;
        string DocHasMaterial;
        string DocHasProduct;
        string DocHasProperty;
        string DocHasWorkCenter;
        string DocHasWorkCenterAlt;
        string DocPermitsQuantitySufficient;
        int mDocID;
        const string MaterialCategory_ShowAll = "[ Show All ]";
        const string mstrSelectMaterialCategory = "<Select Category>";
        const int cfraMaterialHeight_Expand = 266;//218;
        const int cfraMaterialHeight_Normal = 182;///134; In VB the zise is 134
        const int ccmdMaterialTop_Normal = 154; //112;In vb taken top property for button in fraMaterial; In c# I have took location Y value
        const int ccmdMaterialTop_Expand = 234;//192; 
        const int ctxtBoxHeight_Normal = 20;//16;
        const int ctxtBoxHeight_Expand = 117;//113;
        const int ctxtBOMBoxHeight_Expand = 104;//100;
        const int cintParmsValidated = 19217;
        const int cintEditMaterialRun = 13862;
        const int cintMatlNam = 1;
        const int cintItemCode = 2;
        const int cintMatlAMT = 3;
        const int cintMatlUOM = 4;

        int mintSwapCount;
        int mintParmsValidated;
        bool mblnUserMadeYieldChange;
        bool mblnEnableBackNext;
        int mMaterialImprint_TextMaxLen;// = 255; //Hard Codeded. Refer Material_TextMaxLens() method.
        int mMaterialPerUnitData_TextMaxLen;// = 255; //Hard Codeded. Refer Material_TextMaxLens() method.
        int mMaterialBOMText_TextMaxLen;// = 255; //Hard Codeded. Refer Material_TextMaxLens() method.
        bool mblnDocumentProtected;
        string mstrProtectionPassword;
        bool mblnCalledWizard;
        bool MaterialYieldAvailable = false;
        string mstrOriginalYield;
        int mintListIndex;
        int mintOpUnitID;
        string mstrOpUnitType;
        bool mblnUserMadeMaterialChange;
        int mqrqCalledRequest = 0;
        string mstrSuppressUnload;
        int mlngMyTop;
        int mlngMyLeft;
        bool IsImprintExpanded;
        bool IsBOMExpanded;
        qReturn mqrtQReturn;
        private DataTable dt;
        //Performancetuning
        private System.Windows.Forms.Timer wordKeepAliveTimer;
        bool blnDBUpdateSuccessful;
        string mstrDocFullName;
        enum MaterialValidationMode
        {
            MaterialValidation_Add = 31908,
            MaterialValidation_Update = 25632
        }

        public FormEditMaterials()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();
            try
            {
                SetTargetDocumentName(activeDocument?.Name, activeDocument?.FullName);
            }
            catch
            {
            }
            lstMaterialsWithAmounts.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            lstMaterialsWithAmounts.ColumnHeadersVisible = false;
            dt = new DataTable();
            dt.Columns.Add("MaterialDisplay");
            dt.Columns.Add("Amount");
            dt.Columns.Add("Units");
            dt.Columns.Add("Usage");
            dt.Columns.Add("MaterialDBID");
            dt.Columns.Add("MaterialItemCode");
            dt.Columns.Add("MaterialQS");
            dt.Columns.Add("MaterialTitle");
            dt.Columns.Add("MaterialCategory");
            dt.Columns.Add("MaterialImprint");
            dt.Columns.Add("MaterialPerUnitData");
            dt.Columns.Add("MaterialBOMText");
            dt.Columns.Add("MaterialUseCount");

            lstMaterialsWithAmounts.DataSource = dt;
            lstMaterialsWithAmounts.Columns["MaterialDisplay"].Width = 240;
            lstMaterialsWithAmounts.Columns["Amount"].Width = 70;
            lstMaterialsWithAmounts.Columns["Units"].Width = 80;
            lstMaterialsWithAmounts.Columns["Usage"].Width = 90;

            lstMaterialsWithAmounts.AutoGenerateColumns = false;
            lstMaterialsWithAmounts.Columns["MaterialDBID"].Visible = false;
            lstMaterialsWithAmounts.Columns["MaterialItemCode"].Visible = false;
            lstMaterialsWithAmounts.Columns["MaterialQS"].Visible = false;
            lstMaterialsWithAmounts.Columns["MaterialTitle"].Visible = false;
            lstMaterialsWithAmounts.Columns["MaterialPerUnitData"].Visible = false;
            lstMaterialsWithAmounts.Columns["MaterialBOMText"].Visible = false;
            lstMaterialsWithAmounts.Columns["MaterialImprint"].Visible = false;
            lstMaterialsWithAmounts.Columns["MaterialCategory"].Visible = false;
            lstMaterialsWithAmounts.Columns["MaterialUseCount"].Visible = false;

        }

        private void FormEditMaterials_Load(object sender, EventArgs e)
        {
            try
            {
                if (!TryResolveActiveDocumentForForm())
                {
                    throw new InvalidOperationException("Unable to resolve the target document.Close the form and try again.");
                }
                Cursor = Cursors.WaitCursor;
                tabControl1.TabPages.Remove(tabPage2);
                InitializeTabs(tabControl1, cmdBack, cmdNext, tabPage1);
                ShowTab(0);
                SetToolTip(cboMaterialCategory, "cboMaterialCategory: MaterialCategory");
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

                var result = _documentManager.Document_SelectAttributes(activeDocument.Name);
                mstrDocName = activeDocument.Name;
                mDocID = result.DocumentID;
                mstrProtectionPassword = result.DocumentProtectionPassword;
                mintOpUnitID = result.DocumentOpUnitID;
                mstrOpUnitType = result.DocumentOpUnitType;
                _documentManager.Document_Has(mstrDocName, ref DocIsModule, ref DocHasModule, ref DocHasMaterial, ref DocHasProduct, ref DocHasProperty, ref DocHasWorkCenter, ref DocHasWorkCenterAlt, ref DocPermitsQuantitySufficient);

                mblnYieldActive = true;
                mDocIDNegative = mDocID * -1;
                mrngHold = activeDocument.Application.Selection.Range;

                lblTitle2.Text = "Materials for Operating Unit " + result.DocumentOpUnit;

                CommonService.SwapClear();

                lblMaterialCategory.Visible = true;
                cboMaterialCategory.Visible = true;

                GetMaterialCategory();

                Fill_MaterialUsage();

                cboMaterialUnits.DataSource = _formManager.Fill_MaterialUnit();
                cboMaterialUnits.DisplayMember = "MaterialUOM";
                cboMaterialUnits.ValueMember = "MaterialUOM";
                SetToolTip(cboMaterialUnits, "cboMaterialUnits: MaterialUnit");
                lblOrderBy.Visible = true;

                optMaterialOrderByNumber.Checked = false;
                optMaterialOrderByNumber.Checked = true;

                Load_lstMaterialsWithAmounts();
                MaterialReset();

                mblnUserMadeYieldChange = false;
                shpMapCollapse.Visible = false;

                mblnEnableBackNext = true;

                if (cboMaterialCategory.SelectedValue.ToString() == mstrSelectMaterialCategory)
                {
                    fraMaterial.Visible = false;
                }
                Material_TextMaxLens();

                txtBOMText.Multiline = true;
                txtBOMText.AcceptsReturn = true;

                txtImprintOrPerUnitData.Multiline = true;
                txtImprintOrPerUnitData.AcceptsReturn = true;
                _formManager.NotifyUser_UnIncludedMaterials(mDocID);
                Cursor = Cursors.Default;
                activeDocument.Application.StatusBar = "";

                //Performancetuning
                wordKeepAliveTimer = new System.Windows.Forms.Timer();
                wordKeepAliveTimer.Interval = 1 * 60 * 1000; // every 1 minutes (adjust if needed)
                wordKeepAliveTimer.Tick += WordKeepAliveTimer_Tick;
                wordKeepAliveTimer.Start();
                //end
            }
            catch (Exception ex)
            {
                Document_Protect();
                Cursor = Cursors.Default;
                if (activeDocument != null)
                {
                    activeDocument.Application.StatusBar = "";
                    activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;
                }

                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, "FormEditMaterials_Load");
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "FormEditMaterials_Load");
                }
            }
        }

        private void Load_lstMaterialsWithAmounts()
        {
            SetToolTip(lstMaterialsWithAmounts, lstMaterialsWithAmounts.Name + ": DocID=" + mDocID + " Material");
            List<DocMaterial_SelectDetails> docMaterials = _formManager.Fill_Document_Material(mDocID);
            foreach (var material in docMaterials)
            {
                DataRow dataRow = dt.NewRow();

                dataRow["MaterialDisplay"] = material.MaterialDisplay;
                dataRow["MaterialTitle"] = material.MaterialTitle;
                dataRow["MaterialDBID"] = material.MaterialID;
                dataRow["MaterialItemCode"] = material.ItemCode;

                if (Convert.ToDouble(material.MaterialAmountDisplay) != 0)
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

                dt.Rows.Add(dataRow);
            }

        }

        private void Fill_MaterialUsage()
        {

            cboMaterialUsage.Items.Clear();

            List<MaterialUsageDetails> materialUsages = _formManager.Fill_MaterialUsage();

            List<MaterialUsageDetails> usageDetails = new List<MaterialUsageDetails>();
            if (materialUsages.Any(item => item.MaterialUsage == "Standard"))
            {
                var standardItem = materialUsages.FirstOrDefault(item => item.MaterialUsage == "Standard");
                usageDetails = materialUsages.Where(item => item.MaterialUsage != "Standard").ToList();
                if (standardItem != null)
                {
                    usageDetails.Insert(0, standardItem);
                }

            }

            cboMaterialUsage.DisplayMember = "MaterialUsage";
            cboMaterialUsage.ValueMember = "MaterialUsage";
            cboMaterialUsage.DataSource = usageDetails;

            SetToolTip(cboMaterialUsage, cboMaterialUsage.Name + ": MaterialUsage");


        }

        public void GetMaterialCategory()
        {
            List<MaterialCategoryDetails> MaterialCboValues = new List<MaterialCategoryDetails>()
                {
                    new MaterialCategoryDetails{ MaterialCategory = mstrSelectMaterialCategory, MaterialCategoryDescription= "not applicable" , Material_ItemCodeFormat=" " , Product_ItemCodeFormat=" ",MaterialHasImprint= "N",MaterialHasPerUnitData= "N", MaterialHasBOMText="N" },
                    new MaterialCategoryDetails{ MaterialCategory = MaterialCategory_ShowAll, MaterialCategoryDescription= "not applicable" , Material_ItemCodeFormat=" " , Product_ItemCodeFormat=" ",MaterialHasImprint= "N",MaterialHasPerUnitData= "N", MaterialHasBOMText="N" },

                };

            List<MaterialCategoryDetails> materialCategories = _formManager.Fill_MaterialCategory(mintOpUnitID);
            MaterialCboValues.AddRange(materialCategories);

            cboMaterialCategory.DisplayMember = "MaterialCategory";
            cboMaterialCategory.ValueMember = "MaterialCategory";
            cboMaterialCategory.DataSource = MaterialCboValues;

            cboMaterialCategory.Enabled = true;
            cboMaterialCategory.DrawMode = DrawMode.OwnerDrawFixed;
            cboMaterialCategory.DrawItem += cboMaterialCategory_DrawItem;
            cboMaterialCategory.Format += cboMaterialCategory_Format;


        }

        private void cboMaterialCategory_Format(object sender, ListControlConvertEventArgs e)
        {
            var MaterialCatecoryDetails = (MaterialCategoryDetails)e.ListItem;
            e.Value = MaterialCatecoryDetails.MaterialCategory;
        }

        private void cboMaterialCategory_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            if (e.Index >= 0)
            {
                var materialCatecoryDetails = (MaterialCategoryDetails)cboMaterialCategory.Items[e.Index];

                System.Drawing.Font drawFont = e.Font;
                int padding = 180;

                e.Graphics.DrawString(materialCatecoryDetails.MaterialCategory, drawFont, Brushes.Black, e.Bounds.X, e.Bounds.Y);
                e.Graphics.DrawString(materialCatecoryDetails.MaterialCategoryDescription, drawFont, Brushes.Black, e.Bounds.X + padding, e.Bounds.Y);
            }

            e.DrawFocusRectangle();
        }


        private void FormEditMaterials_Activate(object sender, EventArgs e)
        {
            const string myProcName = "FormEditMaterials_Activate";
            int intListScan;
            string strColumnWidths;
            int intRowsWithAmounts;
            string strActiveTransaction;

            UserForm__Positioning();
            strActiveTransaction = DatabaseHelper.Transaction_Active();
            if (!string.IsNullOrEmpty(strActiveTransaction) && strActiveTransaction != mstrDocName)
            {
                MessageBox.Show("CANNOT RUN THIS LOGIC, DUE TO ACTIVE DATABASE TRANSACTION FOR '" + strActiveTransaction, myProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.Close();
                CommonService.CleanUpBackgroundWordProcesses();
                return;
            }

            try
            {
                if (mblnCalledWizard)
                {
                    mrngYield = null;
                    cmdFinish.Text = "&Close";
                }
                else
                {
                    Yield_Initialize();
                    cmdFinish.Text = "&Save";
                }

                strColumnWidths = "";

                intRowsWithAmounts = 0;
                for (intListScan = 0; intListScan < lstMaterialsWithAmounts.Rows.Count; intListScan++)
                {
                    if (lstMaterialsWithAmounts.Rows[intListScan].Cells["Amount"].Value.ToString() != "0")
                    {
                        intRowsWithAmounts = intRowsWithAmounts + 1;
                        break;
                    }
                }

                if (intRowsWithAmounts == 0)
                {
                    Material_Amounts();
                }

                SetToolTip(cmdFinish, "Save Document, then Terminate Function (TABLE will be REPLACED)");
                SetToolTip(cboMaterial, cboMaterial.Text);
                SetToolTip(btnCopyMaterialName, "Click to send current Full Material Name to Paste Buffer");
                SetToolTip(cmdImprintOrPerUnitDataExpand, "Click to expand Imprint/Per Unit Data Text field");
                SetToolTip(cmdBOMTextExpand, "Click to expand B.O.M. Text field");
                SetToolTip(cmdPasteImprint, "Click to enter the Paste buffer into Imprint/Per Unit Text field");
                SetToolTip(cmdPasteBOMText, "Click to enter the Paste buffer into B.O.M. Text field");
                SetToolTip(cmdCharMap, "Click to open Character Map to Select/Copy Special Characters");
                SetToolTip(cmdPlaceName, "Single-Click a Material directly above, then click button to place that Material name in document");
                SetToolTip(lblMap6, "To skip to this step, click here.");
                SetToolTip(shp0, "To skip to this step, click here.");
                SetToolTip(lblMap5, "To skip to this step, click here.");
                SetToolTip(lblMap0, "To skip to this step, click here.");
                SetToolTip(lblItemWithVolition, "MICK Material Item Code with Volition");
                SetToolTip(cmdCancel, "Terminate Function with NO Output");
                SetToolTip(cmdBack, "Show Previous Wizard Page");
                SetToolTip(cmdNext, "Show Next Wizard Page");

            }

            catch (Exception ex)
            {
                if (activeDocument.ProtectionType == Word.WdProtectionType.wdNoProtection)
                {
                    Document_Protect();

                }
                Cursor = Cursors.Default;
                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, myProcName);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, myProcName);
                }

            }

        }

        private void Material_Amounts()
        {
            Word.Table tblDocMaterial;
            int intColumnCount;
            int intRowCount;
            int intRowScan;
            string strMaterialName;
            string strIdentifier;
            string strCellThree;
            double strAmount;
            string strUnitOM;
            int intListScan;
            int intListUpdateCount;
            int intCommaPos;
            string[,] astrMaterial = new string[4, 101];

            Document_UnProtect();
            try
            {
                tblDocMaterial = CommonService.Table_Find(activeDocument, "Material Name", "EDP No.");
                if (tblDocMaterial is null)
                {
                    Document_Protect();
                    return;
                }
                intColumnCount = tblDocMaterial.Columns.Count;
                intRowCount = tblDocMaterial.Rows.Count;
                for (intRowScan = 0; intRowScan < intRowCount; intRowScan++)
                {
                    strMaterialName = CommonService.Table_CellText(tblDocMaterial, intRowScan, 1);
                    strIdentifier = CommonService.Table_CellText(tblDocMaterial, intRowScan, 2);
                    strCellThree = CommonService.Table_CellText(tblDocMaterial, intRowScan, 3);
                    intCommaPos = strCellThree.IndexOf(",", 0);

                    if (intCommaPos > 1)
                    {
                        strCellThree = strCellThree.Substring(0, intCommaPos) + strCellThree.Substring(intCommaPos + 1);
                    }

                    strAmount = double.Parse(strCellThree);
                    switch (intColumnCount)
                    {
                        case 3:
                            if (strCellThree.ToUpper().IndexOf("KG") >= 0)
                            {
                                strUnitOM = "KG";
                            }
                            else if (strCellThree.ToUpper().IndexOf("L") >= 0)
                            {
                                strUnitOM = "LITERS";
                            }
                            else
                            {
                                strUnitOM = "KG";
                            }
                            break;

                        case 4:
                            strUnitOM = CommonService.Table_CellText(tblDocMaterial, intRowScan, 4);
                            break;
                        default:
                            strUnitOM = "";
                            break;

                    }

                    astrMaterial[cintMatlNam, intRowScan] = strMaterialName;
                    astrMaterial[cintItemCode, intRowScan] = strIdentifier;
                    astrMaterial[cintMatlAMT, intRowScan] = strAmount.ToString();
                    astrMaterial[cintMatlUOM, intRowScan] = strUnitOM;

                }
                Document_Protect();
                intListUpdateCount = 0;
                for (intListScan = 0; intListScan < lstMaterialsWithAmounts.Rows.Count; intListScan--)
                {
                    if (lstMaterialsWithAmounts.Rows[intListScan].Cells["Amount"].Value.ToString() == "0")
                    {
                        for (intRowScan = 1; intRowScan <= intRowCount; intRowScan++)
                        {
                            if (lstMaterialsWithAmounts.Rows[intListScan].Cells["MaterialItemCode"].Value.ToString() == astrMaterial[cintItemCode, intRowScan])
                            {
                                intListUpdateCount = intListUpdateCount + 1;
                                break;
                            }

                        }
                    }
                    if (lstMaterialsWithAmounts.Rows[intListScan].Cells["Amount"].Value.ToString() == "0")
                    {
                        for (intRowScan = 1; intRowScan <= intRowCount; intRowScan++)
                        {
                            if (lstMaterialsWithAmounts.Rows[intListScan].Cells["MaterialDisplay"].Value.ToString() == astrMaterial[cintMatlNam, intRowScan]
)
                            {

                                intListUpdateCount = intListUpdateCount + 1;
                                break;
                            }
                        }
                    }
                }

                if (intListUpdateCount > 0)
                {
                    MessageBox.Show("NOTE: the database has NO Amounts for this Document's Materials." + "\r\n" + "\r\n" + "This function has copied " + intListUpdateCount + " Amounts from the Document to this Wizard's Material-List." + "\r\n" + "\r\n" + "If you <Save>, these AMOUNTS will be Saved to the Database (as Changes)" + "\r\n" + "\r\n" + "If you <Cancel>, you will be warned about discarding changes, and Wizards may show no Materials.", "Edit Material Conversion", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

            }
            catch (Exception ex)
            {
                Document_Protect();
            }

        }
        private void Yield_Initialize()
        {
            const string MyProcName = "Yield_Initialize";
            Word.Section secWork;
            Word.Range rng;
            bool blnFound;
            Word.Paragraph parYield;
            string strStandardYieldPhrase;

            if (!MaterialYieldAvailable)
            {
                return;
            }
            secWork = null;
            foreach (Word.Section section in activeDocument.Sections)
            {
                if (section.ProtectedForForms == true)
                {
                    secWork = section;
                    break;
                }
            }
            if (secWork is null)
            {
                CommonService.Error_Handler("44092", "NO PROTECTED SECTION", MyCompName, MyProcName);
                return;
            }
            Document_UnProtect();
            try
            {
                rng = secWork.Range;
                CommonService.Find_Setup(rng.Find);
                rng.Find.Text = "Standard Yield*: ";
                rng.Find.Wrap = Word.WdFindWrap.wdFindContinue;
                rng.Find.MatchWildcards = true;
                blnFound = rng.Find.Execute();

                if (blnFound == true)
                {
                    rng.Select();
                    strStandardYieldPhrase = rng.Text;
                    rng.Move(Word.WdUnits.wdWord, 1);
                    rng.EndOf(Word.WdUnits.wdParagraph, Word.WdMovementType.wdExtend); // Extend selection to the end of the paragraph
                    rng.End -= 1; // Decreament the end position by 1 character


                    mrngYield = rng;
                    txtStdYield.Text = rng.Text;
                    mstrOriginalYield = rng.Text;
                    parYield = mrngYield.Paragraphs[1];

                }
                else
                {
                    txtStdYield.Text = "";
                    mrngYield = null;
                }
                Document_Protect();
            }
            catch (Exception ex)
            {
                Document_Protect();
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


        private void Document_Protect()
        {
            if (mblnDocumentProtected)
            {
                CommonService.DocumentProtection_On(activeDocument, mstrProtectionPassword);
            }
        }

        private void Document_UnProtect()
        {

            if (activeDocument.ProtectionType == Word.WdProtectionType.wdNoProtection)
            {
                mblnDocumentProtected = false;
            }
            else
            {
                mblnDocumentProtected = true;
            }
            CommonService.DocumentProtection_Off(activeDocument, mstrProtectionPassword);

        }


        private void Setup_Document()
        {
            Document_UnProtect();
            Setup_Document_MBR(true);
        }

        private void Setup_Document_MBR(bool pblnReplaceTable)
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
            if (!TryResolveActiveDocumentForForm())
            {
                //throw new InvalidOperationException("Unable to access the target document, Close the form and Save again.");
                MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.Close();
                return;
            }
            if (pblnReplaceTable)
            {
                tblMaterial = CommonService.Table_Find_MaterialMBR(activeDocument);

                if (tblMaterial is null)
                {
                    MessageBox.Show("COULD NOT FIND MATERIAL TABLE - TERMINATING", MyCompName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                CommonService.Table_Delete(tblMaterial, activeDocument, pstrProtectionPassword: mstrProtectionPassword);
            }
            
            activeDocument.Application.Selection.Range.Select();
            activeDocument.Application.Selection.Font.Size = 4;

            CommonService.AutoText_Insert("MaterialsMBR", ref intRowScan, lstMaterialsWithAmounts.Rows.Count, tableAutoText: tblMaterial, wordApp: activeDocument.Application);
            tblMaterial = CommonService.Table_Find_MaterialMBR(activeDocument);
            tblMaterial.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            for (intListScan = 0; intListScan < lstMaterialsWithAmounts.Rows.Count; intListScan++)
            {
                if (lstMaterialsWithAmounts.Rows[intListScan].IsNewRow) continue; // added :there are no empty rows (optional check)

                int tableMaterialRowIndex = intRowScan + intListScan;

                tblMaterial.Cell(tableMaterialRowIndex, 1).Range.Text = lstMaterialsWithAmounts.Rows[intListScan].Cells["MaterialItemCode"].Value.ToString();
                MaterialUsage = lstMaterialsWithAmounts.Rows[intListScan].Cells["Usage"].Value.ToString();

                if (MaterialUsage.ToUpper() != "STANDARD")
                {
                    rngWork = tblMaterial.Cell(tableMaterialRowIndex, 1).Range;
                    rngWork.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    rngWork.Text = MaterialUsage + " ";
                    rngWork.Font.Italic = 1;
                }

                MaterialPerUnitData = lstMaterialsWithAmounts.Rows[intListScan].Cells["MaterialPerUnitData"].Value.ToString();
                MaterialImprint = lstMaterialsWithAmounts.Rows[intListScan].Cells["MaterialImprint"].Value.ToString();
                MaterialBOMText = lstMaterialsWithAmounts.Rows[intListScan].Cells["MaterialBOMText"].Value.ToString();

                tblMaterial.Cell(tableMaterialRowIndex, 2).Range.Text = lstMaterialsWithAmounts.Rows[intListScan].Cells["MaterialTitle"].Value.ToString();

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

                //MaterialUsageDetails selectedUsage = (MaterialUsageDetails)cboMaterialUsage.SelectedItem;
                //blnAllowAmount = selectedUsage.MaterialUsageAllowAmount;

                blnAllowAmount = cboMaterialUsage.Items.OfType<MaterialUsageDetails>().FirstOrDefault(x => x.MaterialUsage == MaterialUsage).MaterialUsageAllowAmount; // Fix for user-handle error: Clicking Save instead of Update 

                if (blnAllowAmount)
                {
                    if (lstMaterialsWithAmounts.Rows[intListScan].Cells["MaterialQS"].Value.ToString() == CommonService.qdlMaterialAmount_qs_No_Amount)
                    {
                        if (lstMaterialsWithAmounts.Rows[intListScan].Cells["Amount"].Value.ToString() == "")
                        {
                            tblMaterial.Cell(tableMaterialRowIndex, 3).Range.Text = CommonService.qdlMaterialAmount_qs_No_Amount;
                        }
                        else
                        {
                            tblMaterial.Cell(tableMaterialRowIndex, 3).Range.Text = CommonService.qdlMaterialAmount_qs_With_Amount + " " + lstMaterialsWithAmounts.Rows[intListScan].Cells["Amount"].Value.ToString() + " " + lstMaterialsWithAmounts.Rows[intListScan].Cells["Units"].Value.ToString();
                        }
                    }
                    else
                    {
                        tblMaterial.Cell(tableMaterialRowIndex, 3).Range.Text = lstMaterialsWithAmounts.Rows[intListScan].Cells["Amount"].Value.ToString() + " " + lstMaterialsWithAmounts.Rows[intListScan].Cells["Units"].Value.ToString();
                    }
                }
            }

        }

        private void Material_TextMaxLens()
        {
            const string cMTL_BOMText = "MTL_BOMText";
            const string cMTL_Imprint = "MTL_Imprint";
            const string cMTL_PerUnitData = "MTL_PerUnitData";
            List<SelctedProperty> properties = _formManager.Property_Select("N", "");
            Dictionary<string, int> nameTextLengthMap = new Dictionary<string, int>();
            foreach (var prop in properties)
            {
                if (prop.PropertyName == cMTL_Imprint || prop.PropertyName == cMTL_PerUnitData || prop.PropertyName == cMTL_BOMText)
                {
                    nameTextLengthMap[prop.PropertyName] = prop.PropertyTextMaxLen;
                }
            }

            if (nameTextLengthMap.ContainsKey(cMTL_Imprint))
            {
                mMaterialImprint_TextMaxLen = nameTextLengthMap[cMTL_Imprint];
            }

            if (nameTextLengthMap.ContainsKey(cMTL_PerUnitData))
            {
                mMaterialPerUnitData_TextMaxLen = nameTextLengthMap[cMTL_PerUnitData];
            }

            if (nameTextLengthMap.ContainsKey(cMTL_BOMText))
            {
                mMaterialBOMText_TextMaxLen = nameTextLengthMap[cMTL_BOMText];
            }
        }
        private void cboMaterialCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedItem = (MaterialCategoryDetails)cboMaterialCategory.SelectedItem;
            if (selectedItem != null)
            {
                cboMaterialCategory.Text = selectedItem.MaterialCategory;
            }
            cboMaterialCategory.Refresh();
            if (cboMaterialCategory.SelectedIndex < 1)
            {
                return;
            }

            fraMaterial.Visible = true;
            MaterialImprintOrPerUnitData_Visible(true, cboMaterialCategory.Text);
            MaterialBOMText_Visible(true, cboMaterialCategory.Text);
            cboMaterial__Fill();

            SetToolTip(cboMaterial, cboMaterial.Text);
        }

        private void cboMaterial__Fill()
        {
            string orderBy;
            string materialCategory_Value;
            MaterialCategoryDetails selectedCategory = cboMaterialCategory.SelectedItem as MaterialCategoryDetails;
            cboMaterialCategory.Text = selectedCategory.MaterialCategory;

            cboMaterial.Visible = false;
            lblMaterialFilling.Visible = true;
            this.Refresh();

            if (optMaterialOrderByNumber.Checked)
            {
                orderBy = "N";
            }
            else
            {
                orderBy = "T";
            }
            if (cboMaterialCategory.Items.Count == 0)
            {
                return;
            }
            if (cboMaterialCategory.Text == MaterialCategory_ShowAll)
            {
                materialCategory_Value = "";
            }
            else
            {
                materialCategory_Value = cboMaterialCategory.Text;
            }

            Fill_OpUnit_Material_LoadLogic(orderBy, materialCategory_Value);
            cboMaterial.Visible = (cboMaterial.Items.Count > 0);
            lblMaterialFilling.Visible = false;
        }

        private void Fill_OpUnit_Material_LoadLogic(string orderBy, string materialCategory_Value)
        {
            List<OpUnitMaterial> opUnitMaterials = _formManager.Fill_OpUnit_Material(mintOpUnitID, mstrOpUnitType, orderBy, materialCategory_Value);
            if (orderBy == "T")
            {
                opUnitMaterials = opUnitMaterials.OrderBy(m => m.MaterialTitle_ItemCode).ToList();
            }
            else
            {
                opUnitMaterials = opUnitMaterials.OrderBy(m => m.ItemCode_MaterialTitle).ToList();

            }
            if (opUnitMaterials.Count > 0)
            {
                opUnitMaterials.First();

                if (orderBy == "T")
                {

                    cboMaterial.DisplayMember = "MaterialTitle_ItemCode";
                    cboMaterial.ValueMember = "MaterialTitle_ItemCode";
                }
                else
                {
                    cboMaterial.DisplayMember = "ItemCode_MaterialTitle";
                    cboMaterial.ValueMember = "ItemCode_MaterialTitle";
                }

                cboMaterial.DataSource = opUnitMaterials;
            }
        }

        private void MaterialBOMText_Visible(bool visible, string materialCategory)
        {
            int materialCategoryIndex;
            bool controlsVisible;

            MaterialCategoryDetails selectedCategory = cboMaterialCategory.SelectedItem as MaterialCategoryDetails;
            materialCategoryIndex = cboMaterialCategory.SelectedIndex;
            if (selectedCategory != null)
            {
                controlsVisible = visible && (selectedCategory.MaterialHasBOMText != "N");
                cmdBOMTextExpand.Visible = controlsVisible;
                txtBOMText.Visible = controlsVisible;
                lblBOMText.Visible = controlsVisible;
                cmdPasteBOMText.Visible = controlsVisible;

                if (controlsVisible == true)
                {
                    txtBOMText.Font = new Font("Times New Roman", 8.25F);

                }
                if (selectedCategory.MaterialHasBOMText == "R")
                {
                    lblBOMText.Text = "* B.O.M. Text:";
                }
                else
                {
                    lblBOMText.Text = "B.O.M. Text:";
                }
            }
        }


        public void MaterialImprintOrPerUnitData_Visible(bool visible, string materialCategory)
        {
            int materialCategoryIndex;
            bool controlsVisible;

            MaterialCategoryDetails selectedCategory = cboMaterialCategory.SelectedItem as MaterialCategoryDetails;
            materialCategoryIndex = cboMaterialCategory.SelectedIndex;
            if (selectedCategory != null)
            {
                controlsVisible = visible && ((selectedCategory.MaterialHasImprint != "N") || (selectedCategory.MaterialHasPerUnitData != "N"));
                cmdImprintOrPerUnitDataExpand.Visible = controlsVisible;
                txtImprintOrPerUnitData.Visible = controlsVisible;
                lblImprintOrPerUnitData.Visible = controlsVisible;
                cmdPasteImprint.Visible = controlsVisible;

                if (selectedCategory.MaterialHasImprint != "N")
                {
                    if (selectedCategory.MaterialHasImprint == "R")
                    {
                        lblImprintOrPerUnitData.Text = "* Imprint:";
                    }
                    else
                    {
                        lblImprintOrPerUnitData.Text = "Imprint:";
                    }
                }

                if (selectedCategory.MaterialHasPerUnitData != "N")
                {
                    if (selectedCategory.MaterialHasPerUnitData == "R")
                    {
                        lblImprintOrPerUnitData.Text = "* Per Unit:";
                    }
                    else
                    {
                        lblImprintOrPerUnitData.Text = "Per Unit:";
                    }
                }
            }
        }

        private void cboMaterialUnits_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void MaterialReset()
        {
            lstMaterialsWithAmounts.ClearSelection();
            if (cboMaterial.Items.Count > 0)
            {
                cboMaterial.SelectedIndex = 0;
            }
            cboMaterial.Enabled = true;
            cboMaterialCategory.Enabled = true;
            txtMaterialAmount.Text = "";
            chkQuantitySufficient.Checked = false;
            cboMaterialUnits.SelectedIndex = 0;

            if (cboMaterial.SelectedIndex > 0 && cboMaterial.Visible)
            {
                cboMaterial.Focus();
            }
            cmdMaterialAdd.Text = "Add";
            cmdSwapWithAmounts.Enabled = true;
            lstMaterialsWithAmounts.Enabled = true;
            cmdMaterialUpdate.Enabled = false;
            cmdMaterialDelete.Enabled = false;
            cmdMoveUpWithAmounts.Enabled = false;
            cmdMoveDownWithAmounts.Enabled = false;

            MaterialImprintOrPerUnitData_Visible(true, cboMaterialCategory.Text);
            MaterialBOMText_Visible(true, cboMaterialCategory.Text);
            if (this.txtImprintOrPerUnitData.Visible == true)
            {
                txtImprintOrPerUnitData.Text = "";
            }
            if (this.txtBOMText.Visible == true)
            {
                txtBOMText.Text = "";
            }
            cmdMaterialUpdate.Top = ccmdMaterialTop_Normal;
            cmdMaterialAdd.Top = ccmdMaterialTop_Normal;
            cmdSwapWithAmounts.Top = ccmdMaterialTop_Normal;
            fraMaterial.Height = cfraMaterialHeight_Normal;
            lstMaterialsWithAmounts.Visible = true;

        }
        private bool Material_Valid(MaterialValidationMode pmvmMode)
        {
            const int cintColCount = 10;
            double dblMaterialAMount;
            int MaterialAmountScale;
            bool material_Valid = false;



            if (CommonService.qdlColCount != cintColCount)
            {
                MessageBox.Show("PROGRAM ERROR: This logic expects " + cintColCount + " Columns, not " + CommonService.qdlColCount, "MBR Edit Material", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (cboMaterialCategory.SelectedValue.ToString() != mstrSelectMaterialCategory)
            {
                if (cboMaterial.Text == "")
                {
                    if (cboMaterial.Visible)
                    {
                        cboMaterial.Focus();
                    }

                    MessageBox.Show("Please select a Material.", "MBR Edit Material", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
            }
            else
            {
                MessageBox.Show("Please select a Material Category.", "MBR Edit Material", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
            if (txtMaterialAmount.Visible)
            {
                if (txtMaterialAmount.Text == "")
                {
                    if (!chkQuantitySufficient.Checked)
                    {
                        if (chkQuantitySufficient.Visible == true)
                        {
                            DialogResult result = MessageBox.Show(
                            "No quantity Amount has been entered, or Quantity Sufficient checked.  Is this correct?",
                            "No \"Quantity\" or \"q.s. \"",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.ServiceNotification);
                            switch (result)
                            {
                                case DialogResult.Yes:
                                    break;
                                case DialogResult.No:
                                    txtMaterialAmount.Focus();
                                    return false;
                                    //break;
                            }

                        }
                        else
                        {
                            DialogResult result = MessageBox.Show(
                            "No quantity Amount has been entered.  Is this correct?",
                           "No \"Quantity\" entered",
                            MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question,
                           MessageBoxDefaultButton.Button1,
                           MessageBoxOptions.ServiceNotification);
                            switch (result)
                            {
                                case DialogResult.Yes:
                                    break;
                                case DialogResult.No:
                                    txtMaterialAmount.Focus();
                                    return false;
                            }
                        }
                    }
                }
                else
                {
                    try
                    {
                        dblMaterialAMount = Convert.ToDouble(txtMaterialAmount.Text);
                        if (dblMaterialAMount <= 0)
                        {
                            txtMaterialAmount.Focus();
                            MessageBox.Show("Please enter a positive amount; " + "\r\n" + "\r\n" + "optionally blank if 'q.s. (Quantity Sufficient)' is checked.", "MBR Edit Material", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return false;
                        }
                        MaterialAmountScale = _formManager.GetMaterialAmountScale(txtMaterialAmount.Text);

                        if (MaterialAmountScale > 4)
                        {
                            txtMaterialAmount.Focus();
                            MessageBox.Show("Material Amount is limited to 4 digits after the decimal point.", "'Decimals' too long", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            txtMaterialAmount.Focus();
                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        if (txtMaterialAmount.Visible == true && (!chkQuantitySufficient.Checked))
                        {
                            txtMaterialAmount.Focus();
                            MessageBox.Show("Please enter a numeric value.");
                        }
                        return material_Valid = false;
                    }
                }
            }
            else
            {
                txtMaterialAmount.Text = "";
            }

            MaterialCategoryDetails selectedMaterialCategory = (MaterialCategoryDetails)cboMaterialCategory.SelectedItem;
            if (selectedMaterialCategory.MaterialHasImprint == "R")
            {
                if (txtImprintOrPerUnitData.Text.Trim() == "")
                {
                    txtImprintOrPerUnitData.Focus();
                    MessageBox.Show("Imprint data is Required for this Material Category", "'Imprint' Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            if (selectedMaterialCategory.MaterialHasImprint != "N")
            {
                if ((txtImprintOrPerUnitData.Text).Length > mMaterialImprint_TextMaxLen)
                {
                    txtImprintOrPerUnitData.Focus();
                    MessageBox.Show("Imprint data is limited to " + mMaterialImprint_TextMaxLen + " characters", "'Imprint' too long", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            if (selectedMaterialCategory.MaterialHasPerUnitData == "R")
            {
                if (txtImprintOrPerUnitData.Text.Trim() == "")
                {
                    txtImprintOrPerUnitData.Focus();
                    MessageBox.Show("Per-Unit-Data is Required for this Material Category", "'Per-Unit-Data' Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            if (selectedMaterialCategory.MaterialHasPerUnitData != "N")
            {
                if (txtImprintOrPerUnitData.Text.Length > mMaterialPerUnitData_TextMaxLen)
                {
                    txtImprintOrPerUnitData.Focus();
                    MessageBox.Show("Per-Unit-Data is limited to " + mMaterialPerUnitData_TextMaxLen + " characters", "'PerUnitData' too long", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;

                }
            }

            if (selectedMaterialCategory.MaterialHasBOMText == "R")
            {
                if (txtBOMText.Text.Trim() == "")
                {
                    txtBOMText.Focus();
                    MessageBox.Show("BOM Text is required for this Category.", "'B.O.M. Text' Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            if (selectedMaterialCategory.MaterialHasBOMText != "N")
            {
                if ((txtBOMText.Text).Length > mMaterialBOMText_TextMaxLen)
                {
                    txtBOMText.Focus();
                    MessageBox.Show("BOMText is limited to " + mMaterialBOMText_TextMaxLen + " characters", "'BOMText' too long", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            return material_Valid = true;
        }

        private void cmdMaterialAdd_Click(object sender, EventArgs e)
        {
            int intToRow;
            if (cmdMaterialAdd.Text == "Reset")
            {
                MaterialReset();
                return;
            }
            if (!Material_Valid(MaterialValidationMode.MaterialValidation_Add))
            {
                return;
            }

            MaterialApply();
            MaterialReset();
            mblnUserMadeMaterialChange = true;
            if (mqrqCalledRequest == CommonService.qRequest_Add)
            {
                UserForm__SHUTDOWN(qReturn.qReturn_Finish);
            }

            if (IsImprintExpanded)
            {

                cmdImprintOrPerUnitDataExpand_Click(null, EventArgs.Empty);
                IsImprintExpanded = false;
            }
            if (IsBOMExpanded)
            {
                cmdBOMTextExpand_Click(null, EventArgs.Empty);
                IsBOMExpanded = false;
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
        public void MaterialApply()
        {
            string MaterialImprint;
            string MaterialPerUnitData;
            string MaterialBOMText;
            var selectedCboMaterial = (OpUnitMaterial)cboMaterial.Items[cboMaterial.SelectedIndex];
            var selectedCboMaterialCategory = (MaterialCategoryDetails)cboMaterialCategory.Items[cboMaterialCategory.SelectedIndex];
            var MaterialCategory = selectedCboMaterial.MaterialCategory;
            var MaterialDisplay = cboMaterial.Text?.ToString().Trim();
            var MaterialAmount = txtMaterialAmount.Text?.Trim();
            var MaterialUsage = cboMaterialUsage.Text?.ToString().Trim();
            var MateriaID = selectedCboMaterial.MaterialID;
            var MaterialItemCode = selectedCboMaterial.ItemCode;
            var MaterialQS = chkQuantitySufficient.Checked ? CommonService.qdlMaterialAmount_qs_No_Amount : "";
            var MaterialTitle = selectedCboMaterial.MaterialTitle;
            var MaterialUnits = !string.IsNullOrEmpty(MaterialAmount) ? cboMaterialUnits.Text?.ToString() : string.Empty;
            MaterialBOMText = txtBOMText.Text.Trim();

            DataRow dataRow = dt.NewRow();

            if (lstMaterialsWithAmounts.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = lstMaterialsWithAmounts.SelectedRows[0];
                DataRowView rowView = (DataRowView)selectedRow.DataBoundItem;
                DataRow row = rowView.Row;
                if (Convert.ToInt32(row["MaterialDBID"]) == MateriaID)
                {
                    row["Amount"] = MaterialAmount;
                    if (txtMaterialAmount.Text == "")
                    {
                        row["Units"] = "";
                    }
                    else
                    {
                        row["Units"] = MaterialUnits;
                    }
                    if (selectedCboMaterialCategory.MaterialHasImprint == "N")
                    {
                        MaterialImprint = "";
                    }
                    else
                    {
                        MaterialImprint = txtImprintOrPerUnitData.Text.Trim();
                    }
                    if (selectedCboMaterialCategory.MaterialHasPerUnitData == "N")
                    {
                        MaterialPerUnitData = "";
                    }
                    else
                    {
                        MaterialPerUnitData = txtImprintOrPerUnitData.Text.Trim();
                    }
                    row["Usage"] = MaterialUsage;
                    row["MaterialBOMText"] = MaterialBOMText;
                    row["MaterialItemCode"] = MaterialItemCode;
                    row["MaterialTitle"] = MaterialTitle;
                    row["MaterialPerUnitData"] = MaterialPerUnitData;
                    row["MaterialImprint"] = MaterialImprint; 
                }
            }

            else
            {
                dataRow["MaterialDBID"] = MateriaID;
                dataRow["MaterialDisplay"] = MaterialDisplay;
                dataRow["Amount"] = MaterialAmount;
                if (txtMaterialAmount.Text == "")
                {
                    dataRow["Units"] = "";
                }
                else
                {
                    dataRow["Units"] = MaterialUnits;
                }
                dataRow["Usage"] = MaterialUsage;
                if (selectedCboMaterialCategory.MaterialHasImprint == "N")
                {
                    MaterialImprint = "";
                }
                else
                {
                    MaterialImprint = txtImprintOrPerUnitData.Text.Trim();
                }
                if (selectedCboMaterialCategory.MaterialHasPerUnitData == "N")
                {
                    MaterialPerUnitData = "";
                }
                else
                {
                    MaterialPerUnitData = txtImprintOrPerUnitData.Text.Trim();
                }

                dataRow["MaterialBOMText"] = MaterialBOMText;
                dataRow["MaterialImprint"] = MaterialImprint;
                dataRow["MaterialPerUnitData"] = MaterialPerUnitData;
                dataRow["MaterialCategory"] = MaterialCategory;
                dataRow["MaterialQS"] = MaterialQS;
                dataRow["MaterialItemCode"] = MaterialItemCode;
                dataRow["MaterialTitle"] = MaterialTitle;

                dt.Rows.Add(dataRow);
            }
            lstMaterialsWithAmounts.Refresh();
            txtBOMText.Clear();
            txtImprintOrPerUnitData.Clear();
        }

        private string DocMaterial_MULTI(string materialCategory, string materialImprint, string materialPerUnitData, string materialBOMText)
        {
            string[] multiArray = new string[4];
            multiArray[CommonService.qdlSubColMaterialCategory - CommonService.qdlSubMaterial_MULTI_Offset - 1] = materialCategory;
            multiArray[CommonService.qdlSubColMaterialImprint - CommonService.qdlSubMaterial_MULTI_Offset - 1] = materialImprint;
            multiArray[CommonService.qdlSubColMaterialPerUnitData - CommonService.qdlSubMaterial_MULTI_Offset - 1] = materialPerUnitData;
            multiArray[CommonService.qdlSubColMaterialBOMText - CommonService.qdlSubMaterial_MULTI_Offset - 1] = materialBOMText;

            return string.Join(CommonService.Multi_Array_Separator, multiArray);
        }

        private void cmdSwapWithAmounts_Click(object sender, EventArgs e)
        {
            string strMaterialOLD;
            string strMaterialNEW;
            int intMaterialRowOLD;
            int intMaterialRowNEW;

            if (cboMaterial.SelectedIndex < 0)
            {
                cboMaterial.Focus();
                MessageBox.Show("Please specify a Material to replace with.");
                return;
            }

            int selectedMaterialIndex = lstMaterialsWithAmounts.CurrentCell.RowIndex;
            if (selectedMaterialIndex < 0)
            {
                lstMaterialsWithAmounts.Focus();
                MessageBox.Show("Please specify a Material to be replaced.");
                return;
            }

            strMaterialOLD = lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["MaterialDisplay"].Value?.ToString();
            strMaterialNEW = cboMaterial.Text;
            intMaterialRowOLD = selectedMaterialIndex;
            intMaterialRowNEW = cboMaterial.SelectedIndex;

            if (CommonService.SwapInvolvement(strMaterialOLD))
            {
                lstMaterialsWithAmounts.Focus();
                SwapErrorMsg(strMaterialOLD);
                return;
            }
            if (CommonService.SwapInvolvement(strMaterialNEW))
            {
                cboMaterial.Focus();
                SwapErrorMsg(strMaterialNEW);
                return;
            }

            mblnUserMadeMaterialChange = true;
            CommonService.SwapRequest(strMaterialOLD, strMaterialNEW);

            lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialDisplay"].Value = cboMaterial.Text;
            lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialDBID"].Value = ((OpUnitMaterial)cboMaterial.SelectedItem).MaterialID;
            lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialItemCode"].Value = ((OpUnitMaterial)cboMaterial.SelectedItem).ItemCode;
            lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialTitle"].Value = ((OpUnitMaterial)cboMaterial.SelectedItem).MaterialTitle;

            //lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["Amount"].Value = txtMaterialAmount.Text;
            //lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["Units"].Value = cboMaterialUnits.Text;
            //lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["Usage"].Value = cboMaterialUsage.Text;

            //if (((MaterialCategoryDetails)cboMaterialCategory.SelectedItem).MaterialHasImprint == "P")
            //{
            //    lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialImprint"].Value = txtImprintOrPerUnitData.Text;
            //}
            //else  
            //{
            //    lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialImprint"].Value = string.Empty;
            //}

            //if (((MaterialCategoryDetails)cboMaterialCategory.SelectedItem).MaterialHasPerUnitData == "P")
            //{
            //    lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialPerUnitData"].Value = txtImprintOrPerUnitData.Text;
            //}
            //else 
            //{
            //    lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialPerUnitData"].Value = string.Empty;
            //}
            //if (((MaterialCategoryDetails)cboMaterialCategory.SelectedItem).MaterialHasBOMText == "P")
            //{
            //    lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialBOMText"].Value = txtBOMText.Text;
            //}
            //else 
            //{
            //    lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialBOMText"].Value = string.Empty;
            //}
            //lstMaterialsWithAmounts.Rows[intMaterialRowOLD].Cells["MaterialQS"].Value = cboMaterialUsage.Text;
        }

        private void SwapErrorMsg(string materialName)
        {
            MessageBox.Show("Cannot Swap" + "\r\n" + "\r\n" + "This Material is already being Swapped:" + "\r\n" + "\r\n" + materialName, "Swap Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void cmdMaterialUpdate_Click(object sender, EventArgs e)
        {
            if (!Material_Valid(MaterialValidationMode.MaterialValidation_Update))
            {
                return;
            }
            if (txtBOMText.Height != ctxtBoxHeight_Normal)
            {
                cmdBOMTextExpand_Click(sender, e);
            }
            if (txtImprintOrPerUnitData.Height != ctxtBoxHeight_Normal)
            {
                cmdImprintOrPerUnitDataExpand_Click(sender, e);
            }
            int seletedMaterialRow = lstMaterialsWithAmounts.CurrentCell.RowIndex;
            if (lstMaterialsWithAmounts.Rows[seletedMaterialRow].Cells["MaterialDisplay"].ToString() == txtMaterialAmount.Text &&
                lstMaterialsWithAmounts.Rows[seletedMaterialRow].Cells["Amount"].ToString() == cboMaterialUnits.Text &&
                lstMaterialsWithAmounts.Rows[seletedMaterialRow].Cells["Usage"].ToString() == cboMaterialUsage.Text &&


           lstMaterialsWithAmounts.Rows[seletedMaterialRow].Cells["Units"].ToString() == cboMaterialUnits.Text)
            {
                MessageBox.Show("No Update: you did not make a change.", "MBR Edit Material Error", MessageBoxButtons.OK);
                return;
            }
            MaterialApply();
            MaterialReset();
            mblnUserMadeMaterialChange = true;

        }

        private void cmdBOMTextExpand_Click(object sender, EventArgs e)
        {
            if (txtBOMText.Height == ctxtBoxHeight_Normal)
            {
                txtBOMText.Height = ctxtBOMBoxHeight_Expand;
                MaterialFrame_Expand(true);
                MaterialListbox_Visible(false);
                MaterialImprintOrPerUnitData_Visible(false, cboMaterialCategory.Text);
                IsBOMExpanded = true;
            }
            else if (txtBOMText.Height != ctxtBoxHeight_Normal)
            {
                txtBOMText.Height = ctxtBoxHeight_Normal;
                MaterialFrame_Expand(false);
                MaterialListbox_Visible(true);
                MaterialImprintOrPerUnitData_Visible(true, cboMaterialCategory.Text);
                IsBOMExpanded = false;
            }
            txtBOMText.Focus();
        }

        private void MaterialListbox_Visible(bool visible)
        {
            lstMaterialsWithAmounts.Visible = visible;
            cmdPlaceName.Visible = visible;
            cmdMoveDownWithAmounts.Visible = visible;
            cmdMoveUpWithAmounts.Visible = visible;
            cmdMaterialDelete.Visible = visible;
            cmdSwapWithAmounts.Visible = visible;

        }

        private void MaterialFrame_Expand(bool expand)
        {
            if (expand)
            {
                fraMaterial.Height = cfraMaterialHeight_Expand;
                cmdMaterialUpdate.Top = ccmdMaterialTop_Expand;
                cmdMaterialAdd.Top = ccmdMaterialTop_Expand;
                cmdSwapWithAmounts.Top = ccmdMaterialTop_Expand;
                cmdCharMap.Top = ccmdMaterialTop_Expand;
            }
            else
            {
                cmdMaterialUpdate.Top = ccmdMaterialTop_Normal;
                cmdMaterialAdd.Top = ccmdMaterialTop_Normal;
                cmdCharMap.Top = ccmdMaterialTop_Normal;
                cmdSwapWithAmounts.Top = ccmdMaterialTop_Normal;
                fraMaterial.Height = cfraMaterialHeight_Normal;

            }
        }

        private void cmdImprintOrPerUnitDataExpand_Click(object sender, EventArgs e)
        {
            if (txtImprintOrPerUnitData.Height == ctxtBoxHeight_Normal)
            {
                txtImprintOrPerUnitData.Height = ctxtBoxHeight_Expand;
                MaterialFrame_Expand(true);
                MaterialListbox_Visible(false);
                MaterialBOMText_Visible(false, cboMaterialCategory.Text);
                IsImprintExpanded = true;
            }
            else if (txtImprintOrPerUnitData.Height != ctxtBoxHeight_Normal)
            {
                txtImprintOrPerUnitData.Height = ctxtBoxHeight_Normal;
                MaterialFrame_Expand(false);
                MaterialListbox_Visible(true);
                MaterialBOMText_Visible(true, cboMaterialCategory.Text);
                IsImprintExpanded = false;
            }
            if (txtImprintOrPerUnitData.Visible)
            {
                txtImprintOrPerUnitData.Focus();
            }
        }

        public void lstMaterialsWithAmounts_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            MaterialSelected();
        }

        private void MaterialSelected()
        {
            try
            {

                int selectedMaterialIndex = lstMaterialsWithAmounts.CurrentCell.RowIndex;
                if (selectedMaterialIndex < 0)
                {
                    return;
                }

                cboMaterialCategory.SelectedValue = lstMaterialsWithAmounts.CurrentRow.Cells["MaterialCategory"].Value.ToString().Trim();

                optMaterialOrderByNumber.Checked = true;
                cboMaterial.SelectedValue= lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["MaterialDisplay"].Value.ToString().Trim();
                var selectedCboMaterial = (OpUnitMaterial)cboMaterial.SelectedItem;



                cboMaterial.Enabled = false;
                cboMaterialCategory.Enabled = false;
                txtMaterialAmount.Text = lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["Amount"].Value.ToString();
                chkQuantitySufficient.Checked = (lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["MaterialQS"].Value?.ToString() == CommonService.qdlMaterialAmount_qs_No_Amount);
                if (!string.IsNullOrEmpty(lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["Amount"].Value.ToString()))
                {
                    cboMaterialUnits.SelectedValue = lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["Units"].Value.ToString().Trim();
                }

                cboMaterialUsage.SelectedValue = lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["Usage"].Value.ToString().Trim();
                txtBOMText.Text = lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["MaterialBOMText"].Value.ToString();

                if (lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["MaterialImprint"].Value.ToString() != "N")
                {
                    if (string.IsNullOrEmpty(txtImprintOrPerUnitData.Text))
                    {
                        txtImprintOrPerUnitData.Text = lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["MaterialImprint"].Value.ToString();
                    }
                }
                if (lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["MaterialPerUnitData"].Value.ToString() != "N")
                {
                    if (string.IsNullOrEmpty(txtImprintOrPerUnitData.Text)) 
                    {
                        txtImprintOrPerUnitData.Text = lstMaterialsWithAmounts.Rows[selectedMaterialIndex].Cells["MaterialPerUnitData"].Value.ToString();
                    }
                }
                if (lblMaterialFilling.Visible == false)
                {
                    lblItemWithVolition.Text = selectedCboMaterial.MICK_Material_Item_Code;

                }
                if (txtMaterialAmount.Visible)
                {
                    txtMaterialAmount.Focus();
                }
                else
                {
                    cboMaterialUsage.Focus();
                }
                cmdMaterialAdd.Text = "Reset";
                cmdSwapWithAmounts.Enabled = false;
                lstMaterialsWithAmounts.Enabled = false;
                cmdMaterialUpdate.Enabled = true;
                cmdMaterialDelete.Enabled = true;
                cmdMoveUpWithAmounts.Enabled = true;
                cmdMoveDownWithAmounts.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Invalid Property Value. Please try again!", "DataGrid - Material Select", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.Close();
                CommonService.CleanUpBackgroundWordProcesses();
            }

        }
        private string DocMaterial_Category(string multiValue)
        {
            string[] multiArray = multiValue.Split(new[] { CommonService.Multi_Array_Separator }, StringSplitOptions.None);
            if (multiArray.Length <= CommonService.qdlSubColMaterialCategory - CommonService.qdlSubMaterial_MULTI_Offset - 1)
            {
                return string.Empty;
            }
            else
            {
                return multiArray[CommonService.qdlSubColMaterialCategory - CommonService.qdlSubMaterial_MULTI_Offset - 1];
            }
        }

        private void cmdCharMap_Click(object sender, EventArgs e)
        {
            string charMapPath = "C:/WINDOWS/System32/charmap.exe";
            System.Diagnostics.Process.Start(charMapPath);
        }
        private void cmdMoveDownWithAmounts_Click(object sender, EventArgs e)
        {
            MoveRow(1);
        }

        private void cmdMoveUpWithAmounts_Click(object sender, EventArgs e)
        {
            MoveRow(-1);
        }
        private void MoveRow(int direction)
        {
            if (lstMaterialsWithAmounts.Rows.Count > 0)
            {
                int selectedRowIndex = lstMaterialsWithAmounts.SelectedRows[0].Index;
                //new row index calculate
                int newRowIndex = selectedRowIndex + direction;
                int currentSrollPosition = lstMaterialsWithAmounts.FirstDisplayedScrollingRowIndex;
                if (selectedRowIndex == 0 && direction == -1)
                {
                    MessageBox.Show("Cannot Move the FIRST item Up", "DataGrid - MoveUp", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return; // to move the selected row, 1 row above
                }
                if (selectedRowIndex == dt.Rows.Count - 1 && direction == 1)
                {
                    MessageBox.Show("Cannot move LAST item DOWN", "DataGrid - MoveDown", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return; // to move the selected row, 1 row below
                }

                if (newRowIndex >= 0 && newRowIndex < dt.Rows.Count)
                {
                    //copy the current row to new position
                    DataRow tempRow = dt.NewRow();
                    tempRow.ItemArray = dt.Rows[selectedRowIndex].ItemArray;
                    dt.Rows[selectedRowIndex].ItemArray = dt.Rows[newRowIndex].ItemArray;
                    dt.Rows[newRowIndex].ItemArray = tempRow.ItemArray;

                    //Reselect the row
                    lstMaterialsWithAmounts.ClearSelection();
                    lstMaterialsWithAmounts.Rows[newRowIndex].Selected = true;

                }
                //Restore the scroll
                if (currentSrollPosition >= 0 && currentSrollPosition <= lstMaterialsWithAmounts.Rows.Count)
                {
                    lstMaterialsWithAmounts.FirstDisplayedScrollingRowIndex = currentSrollPosition;
                }

                if (newRowIndex >= 0 && newRowIndex < lstMaterialsWithAmounts.Rows.Count)
                {
                    lstMaterialsWithAmounts.FirstDisplayedScrollingRowIndex = Math.Max(0, newRowIndex);
                }
            }
        }
        private void cmdMaterialDelete_Click(object sender, EventArgs e)
        {
            int selectedIndex = lstMaterialsWithAmounts.CurrentRow.Index;
            string propertyAssigned;
            if (selectedIndex == 0 && lstMaterialsWithAmounts.Rows.Count > 1)
            {
                if (lstMaterialsWithAmounts.CurrentRow.Cells["Amount"].Value.ToString() == "")
                {
                    MessageBox.Show("Cannot Move ALTERNATE Material to FIRST position", "MoveDown", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
            propertyAssigned = _formManager.DocMaterial_UsedByProperty(mDocID, Convert.ToInt32(lstMaterialsWithAmounts.CurrentRow.Cells["MaterialDBID"].Value));
            if (!string.IsNullOrEmpty(propertyAssigned))
            {
                MessageBox.Show("Cannot delete Material assigned to Property " + propertyAssigned);
                return;
            }

            int useCountValue = 0;
            var useCount = lstMaterialsWithAmounts.Rows[selectedIndex].Cells["MaterialUseCount"].Value;
            if (useCount != null && int.TryParse(useCount.ToString(), out useCountValue))
            {
                if (useCountValue > 0)
                {
                    MessageBox.Show("Cannot delete Material: used " + lstMaterialsWithAmounts.CurrentRow.Cells["MaterialUseCount"].Value + " times;" + "\r\n" + "You may SWAP with another material");
                    return;
                }
            }
            if (MessageBox.Show("Are you sure you want to delete " + lstMaterialsWithAmounts.CurrentRow.Cells["MaterialDisplay"].Value + "?", "MBR Material Delete", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }
            dt.Rows.RemoveAt(selectedIndex);
            mblnUserMadeMaterialChange = true;
            MaterialReset();
        }

        private void cmdFinish_Click(object sender, EventArgs e)
        {
            mstrSuppressUnload = "";
            bool btnWordUpdate = false;
            try
            {
                Cursor = Cursors.WaitCursor;
                /*if (!Valid_Entries()) // To avoid RPC & performance lag
                {
                    return;
                }*/
                if (mblnCalledWizard)
                {
                    UserForm__SHUTDOWN(qReturn.qReturn_Finish);
                    return;
                }
                if (lstMaterialsWithAmounts.Rows.Count < 1 && DocHasMaterial == "Y")
                {
                    lstMaterialsWithAmounts.Focus();
                    MessageBox.Show("There must be at least one Material selected for this document.");
                    return;
                }
                //YieldChanged(); //No need
                //Performancetuning added
                Setup_Database();
                if (!blnDBUpdateSuccessful) {
                    MessageBox.Show("Data save operation failed. Try again.", "Failed to Save", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
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
                ExecuteWordComWithRetry(() => activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorWait);
                this.Left = Application.OpenForms[0].Width - tabControl1.Left;
                if (!TryResolveActiveDocumentForForm())
                {
                    //throw new InvalidOperationException("Unable to access the target document, Close the form and Save again.");
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }
                Word.Selection selection = ExecuteWordComWithRetry(() => activeDocument.Application.Selection);//wordApp.Selection;
                mrngHold = selection.Range;
                //Performancetuning - commented
                // Setup_Database();
                Setup_Document();
                if (!TryResolveActiveDocumentForForm())
                {
                    //throw new InvalidOperationException("Unable to access the target document, Close the form and Save again.");
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }
                CommonService.SwapExecute(ExecuteWordComWithRetry(() => activeDocument));
                mrngHold.Select();
                if (!TryResolveActiveDocumentForForm())
                {
                    //throw new InvalidOperationException("Unable to access the target document, Close the form and Save again.");
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }
                SaveManager.Document_Save_FileAndDatabase(ExecuteWordComWithRetry(() => activeDocument), docpasswordProtect: mstrProtectionPassword, pblnRestoreCommandBar: false);
                Cursor = Cursors.Default;
                if (mstrSuppressUnload == "")
                {
                    this.Close();
                }
                else
                {
                    MessageBox.Show("FORM IS KEPT OPEN, for: " + mstrSuppressUnload);
                }
                if (!TryResolveActiveDocumentForForm())
                {
                    //throw new InvalidOperationException("Unable to access the target document, Close the form and Save again.");
                    MessageBox.Show("Data saved successfully, but unable to access the target document, Click on Refresh menu to update the document.", "Document Idle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }
                btnWordUpdate = true;
                ExecuteWordComWithRetry(() => activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
               
                if (TryResolveActiveDocumentForForm() && ExecuteWordComWithRetry(() => activeDocument?.Application != null))
                {
                    activeDocument.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;
                }
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
                //performancetuning
                StopWordKeepAliveTimer();
                CommonService.CleanUpBackgroundWordProcesses();
            }
        }

        private void Setup_Database()
        {
            if (DatabaseHelper.Transaction_Active() == "")
            {
                DatabaseHelper.Transaction_Begin(mstrDocName);
            }
            Setup_Database_Detail(true);
        }

        private void Setup_Database_Detail(bool blnUseAmounts)
        {
            int intMain;
            int materialDBID;
            string strAmount;
            string strMaterialUOM;
            string MaterialUsage;
            int DocMaterialQuantitySufficient;
            string MaterialImprint;
            string MaterialPerUnitData;
            string MaterialBOMText;

            blnDBUpdateSuccessful = false;

            for (intMain = 0; intMain < lstMaterialsWithAmounts.Rows.Count; intMain++)
            {
                materialDBID = Convert.ToInt32(lstMaterialsWithAmounts.Rows[intMain].Cells["MaterialDBID"].Value);
                MaterialUsage = lstMaterialsWithAmounts.Rows[intMain].Cells["Usage"].Value.ToString();
                MaterialImprint = lstMaterialsWithAmounts.Rows[intMain].Cells["MaterialImprint"].Value.ToString();
                MaterialPerUnitData = lstMaterialsWithAmounts.Rows[intMain].Cells["MaterialPerUnitData"].Value.ToString();
                MaterialBOMText = lstMaterialsWithAmounts.Rows[intMain].Cells["MaterialBOMText"].Value.ToString();

                if (!blnUseAmounts)
                {
                    strAmount = "-1";
                    strMaterialUOM = "";

                }
                else if (lstMaterialsWithAmounts.Rows[intMain].Cells["Amount"].Value.ToString() == "")
                {
                    strAmount = "";
                    strMaterialUOM = "";

                }
                else
                {
                    strAmount = lstMaterialsWithAmounts.Rows[intMain].Cells["Amount"].Value.ToString();
                    strMaterialUOM = lstMaterialsWithAmounts.Rows[intMain].Cells["Units"].Value.ToString();

                }
                if (lstMaterialsWithAmounts.Rows[intMain].Cells["MaterialQS"].Value.ToString() == CommonService.qdlMaterialAmount_qs_No_Amount)
                {
                    DocMaterialQuantitySufficient = 1;
                }
                else
                {
                    DocMaterialQuantitySufficient = 0;
                }
                _formManager.DocMaterial_InsertUpdate(mDocID, materialDBID, strAmount, strMaterialUOM, intMain, DocMaterialQuantitySufficient, MaterialUsage, MaterialImprint.Trim(), MaterialPerUnitData.Trim(), MaterialBOMText.Trim());

            }
            _formManager.DocMaterial_Delete(mDocID, lstMaterialsWithAmounts.Rows.Count - 1);
            mblnUserMadeMaterialChange = true;
            blnDBUpdateSuccessful = true;
        }

        private void YieldChanged()
        {
            mblnUserMadeYieldChange = false;
            if (txtStdYield.Text != mstrOriginalYield)
            {
                mblnUserMadeYieldChange = true;
            }
        }

        private bool Valid_Entries()
        {
            const string cstrMsgBoxTitle = "Edit Material Error";

            mintParmsValidated = 0;
            if (lstMaterialsWithAmounts.Rows.Count == 0 && DocHasMaterial == "Y")
            {
                MessageBox.Show("You must select at least one material.", cstrMsgBoxTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
            if (MaterialYieldAvailable)
            {


                if (fraYieldUnits.Enabled)
                {
                    if (fraYieldUnits__UOMExtract() == "")
                    {

                        return false;
                    }

                }
            }
            mintParmsValidated = cintParmsValidated;
            return true;
        }
        private string fraYieldUnits__UOMExtract()
        {
            if (optYieldUnitsGm.Checked)
            {
                return optYieldUnitsGm.Text;
            }
            else if (optYieldUnitsKg.Checked)
            {
                return optYieldUnitsKg.Text;
            }
            else
            {
                return "";
            }
        }

        private void cmdBack_Click(object sender, EventArgs e)
        {
            ShowPrevTab();
        }

        private void cmdNext_Click(object sender, EventArgs e)
        {
            ShowNextTab();
        }
        private void btnCopyMaterialName_Click(object sender, EventArgs e)
        {
            if (cboMaterial.SelectedValue != null)
            {
                string material = cboMaterial.Text;
                Thread t = new Thread(() =>
                {
                    Clipboard.SetText(material);
                });

                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();

                cboMaterial.Focus();
            }
        }
        private void cmdPasteImprint_Click(object sender, EventArgs e)
        {
            txtImprintOrPerUnitData.Focus();
            SendKeys.Send("^v");
            mblnUserMadeMaterialChange = true;
        }
        private void cmdPasteBOMText_Click(object sender, EventArgs e)
        {
            txtBOMText.Focus();
            SendKeys.Send("^v");
            mblnUserMadeMaterialChange = true;
        }
        private void cmdPlaceName_Click(object sender, EventArgs e)
        {
            if (lstMaterialsWithAmounts.SelectedRows.Count > 0)
            {
                CommonService.TypeStepText(StepTextType.qtImmediate, lstMaterialsWithAmounts.SelectedRows[0].Cells["MaterialDisplay"].Value.ToString(), boldAll: true, useNameAsText: true);
                cmdCancel_Click(sender, e);
                activeDocument.Application.Selection.Font.Bold = 0;
                activeDocument.Application.Selection.TypeText(" ");
            }
            else
            {
                MessageBox.Show("Please first single-click one of the selected Materials of this Master " + "\r\n" + "Record to have its name immediately placed as text in the document.", "Choose material first", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
        private void cmdCancel_Click(object sender, EventArgs e)
        {
            try
            {
                if (mqrqCalledRequest == CommonService.qRequest_Add)
                {
                    UserForm__SHUTDOWN(qReturn.qReturn_Cancel);
                    return;
                }
                YieldChanged();
                if (mblnUserMadeMaterialChange)
                {
                    DialogResult result = MessageBox.Show(
                                   "OK to DISCARD MATERIAL CHANGES?.",
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
        private void cboMaterialUsage_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool allowAmount;
            if (cboMaterialUsage.SelectedIndex < 0)
            {
                return;
            }

            var selectedUsage = (MaterialUsageDetails)cboMaterialUsage.SelectedItem;
            SetToolTip(cboMaterialUsage, selectedUsage.MaterialUsageDescription);

            allowAmount = selectedUsage.MaterialUsageAllowAmount;

            if (!allowAmount)
            {
                txtMaterialAmount.Text = "";
                chkQuantitySufficient.Checked = false;
            }
            lblMaterialAmount.Visible = allowAmount;
            txtMaterialAmount.Visible = allowAmount;
            lblMaterialUnits.Visible = allowAmount;
            cboMaterialUnits.Visible = allowAmount;

            if (DocPermitsQuantitySufficient == "Y")
            {
                chkQuantitySufficient.Visible = allowAmount;
            }
            else
            {
                chkQuantitySufficient.Visible = false;
            }
        }
        private void cboMaterial_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedCboMaterial = (OpUnitMaterial)cboMaterial.SelectedItem;

            if (lblMaterialFilling.Visible == false)
            {
                lblItemWithVolition.Text = selectedCboMaterial.MICK_Material_Item_Code;

                SetToolTip(cboMaterial, cboMaterial.Text);
            }
        }
        private void lblMap5_Click(object sender, EventArgs e)
        {
            cmdFinish_Click(sender, e);
        }
        private void lblMap6_Click(object sender, EventArgs e)
        {
            cmdFinish_Click(sender, e);
        }
        private void optMaterialOrderByTitle_CheckedChanged(object sender, EventArgs e)
        {
            cboMaterial__Fill();
        }
        private void optMaterialOrderByNumber_CheckedChanged(object sender, EventArgs e)
        {
            cboMaterial__Fill();
        }
        private void cboMaterialCategory_KeyDown(object sender, KeyEventArgs e)
        {
            //Todo:   
            //AutoCompleteCombo(cboMaterial, e.KeyCode)
            SetToolTip(cboMaterial, cboMaterial.Text);
        }
        private void cboMaterial_KeyDown(object sender, KeyEventArgs e)
        {
            //Todo:   
            //AutoCompleteCombo(cboMaterialCategory, e.KeyCode)
        }
        private void cboMaterialUnits_KeyDown(object sender, KeyEventArgs e)
        {
            //Todo:   
            //AutoCompleteCombo(cboMaterialUnits, e.KeyCode)
        }

        private void UserForm__Positioning()
        {
            if (mlngMyTop > 0 && mlngMyLeft > 0)
            {
                this.Top = mlngMyTop;
                this.Left = mlngMyLeft;
            }

        }

        public void QInitialize(qRequest qrqRequest)
        {
            const string MyProcName = "QInitialize";
            mblnCalledWizard = true;
            mqrqCalledRequest = (int)qrqRequest;
            switch (qrqRequest)
            {
                case qRequest.qRequest_Display:
                case qRequest.qRequest_Validate:
                    mintParmsValidated = 0;
                    break;
                case qRequest.qRequest_Write:
                    if (mintParmsValidated != cintParmsValidated)
                    {
                        var ex = new Exception();
                        ex.Data["errNum"] = "42981";
                        ex.Data["errDesc"] = "Attempting QWrite with unvalidated data";
                        throw ex;

                    }
                    break;
            }

            shpMapCollapse.Visible = true;
            mblnEnableBackNext = false;
        }
        public void QLet(int QTop, int QLeft, qpsEditMaterialParms p)
        {
            const string MyProcName = "QLet";
            mlngMyTop = QTop;
            mlngMyLeft = QLeft;
            if (p.intEditMaterialRun != 0 && p.intEditMaterialRun != cintEditMaterialRun)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41987";
                ex.Data["errDesc"] = "INVALID RUNCODE=" + p.intEditMaterialRun;
                throw ex;

            }

            if (p.intEditMaterialValidated != 0 && p.intEditMaterialValidated != cintParmsValidated)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41989";
                ex.Data["errDesc"] = "INVALID VALCODE=" + p.intEditMaterialValidated;
                throw ex;
            }

            mintParmsValidated = p.intEditMaterialValidated;
            if (p.intEditMaterialRun == 0)
            {

            }
            if (p.strEditMaterialSelected?.GetUpperBound(1) > 0)
            {

                if (lstMaterialsWithAmounts.ColumnCount != 10)
                {
                    MessageBox.Show("Expected " + 10 + " Columns," + "detected " + lstMaterialsWithAmounts.ColumnCount + " Columns.", MyCompName + "." + MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                }

            }

            if (p.strEditMaterialYieldIgnore == "I")
            {
                mblnYieldActive = false;

            }
        }
        public void QGet(int QTop, int QLeft, qpsEditMaterialParms p)
        {

            QTop = this.Top;
            QLeft = this.Left;
            p.intEditMaterialRun = cintEditMaterialRun;
            p.intEditMaterialValidated = mintParmsValidated;
            CommonService.ListBox_To_Array(lstMaterialsWithAmounts, p.strEditMaterialSelected);
            p.strEditMaterialStandardYield = txtStdYield.Text;

        }
        public qReturn QDisplay(QDisplay_FirstLast QFirstLast)
        {
            QInitialize(qRequest.qRequest_Display);

            return mqrtQReturn;
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
                case QWriteWhat.qWriteWhatever:
                    Setup_Document();
                    break;
                case QWriteWhat.qWriteMBRTable:
                    Setup_Document_MBR(false);
                    break;
                default:
                    MessageBox.Show("UNEXPECTED REQUEST '" + pqwwWriteWhat + "'", MyCompName + "." + MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
            return qReturn.qReturn_Written;
        }
        public qReturn QAdd(int plngMaterialID)
        {
            int index = 0;
            QInitialize((qRequest)CommonService.qRequest_Add);
            for (int intScan = 0; intScan <= cboMaterial.Items.Count - 1; intScan++)
            {
                if (Convert.ToInt32(lstMaterialsWithAmounts.Rows[intScan].Cells["MaterialDBID"].Value) == plngMaterialID)
                {
                    cboMaterial.SelectedIndex = intScan;
                    index = intScan;
                    break;
                }
            }

            if (index > cboMaterial.Items.Count - 1)
            {
                MessageBox.Show("DID NOT FIND Materialid=" + plngMaterialID + " for ADD-LIST", MyCompName, MessageBoxButtons.OK);
                return mqrtQReturn;
            }
            txtMaterialAmount.Text = "0";
            lstMaterialsWithAmounts.Visible = false;
            lblListAmount.Visible = false;
            lblListMaterial.Visible = false;
            lblListUnits.Visible = false;
            cmdMoveDownWithAmounts.Visible = false;
            cmdMoveUpWithAmounts.Visible = false;
            cmdMaterialDelete.Visible = false;
            cmdSwapWithAmounts.Visible = false;
            cmdMaterialUpdate.Visible = false;
            cmdFinish.Enabled = false;
            cboMaterial.Enabled = false;
            lblMap6.Visible = false;
            lblMap5.Visible = false;

            this.Show();
            if (mqrtQReturn == qReturn.qReturn_Finish)
            {
                mintParmsValidated = cintParmsValidated;

            }
            return mqrtQReturn;
        }
       //Performance Tuning
        private void WordKeepAliveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!TryResolveActiveDocumentForForm())
                {
                    return;
                }
                if (activeDocument != null)//&& IsWordRunning())
                {
                    string wordVersion = activeDocument.Application.Version;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Keep-alive failed: " + ex.Message);
            }
        }
        public qReturn QDelete(int plngMaterialID)
        {
            int index = 0;
            int intMaterialCount;

            bool blnDummy = false;
            QInitialize((qRequest)CommonService.qRequest_Delete);
            for (int intScan = 0; intScan <= lstMaterialsWithAmounts.RowCount - 1; intScan++)
            {
                if (Convert.ToInt32(lstMaterialsWithAmounts.Rows[intScan].Cells["MaterialDBID"].Value) == plngMaterialID)
                {
                    lstMaterialsWithAmounts.Rows[intScan].Selected = true;
                    lstMaterialsWithAmounts_CellDoubleClick(null, (DataGridViewCellEventArgs)DataGridViewCellEventArgs.Empty);// need to check
                    index = intScan;
                    break;
                }
            }
            if (index > lstMaterialsWithAmounts.RowCount - 1)
            {
                MessageBox.Show("DID NOT FIND Materialid=" + plngMaterialID + " for DELETE-LIST", MyCompName, MessageBoxButtons.OK);
                UserForm__SHUTDOWN(qReturn.qReturn_Cancel);
                return mqrtQReturn;
            }
            intMaterialCount = lstMaterialsWithAmounts.RowCount;
            cmdMaterialDelete_Click(null, EventArgs.Empty);// need to check
            if (intMaterialCount == lstMaterialsWithAmounts.RowCount)
            {
                UserForm__SHUTDOWN(qReturn.qReturn_Cancel);
            }
            else if (intMaterialCount == lstMaterialsWithAmounts.RowCount - 1)
            {
                UserForm__SHUTDOWN(qReturn.qReturn_Finish);
            }
            else
            {
                MessageBox.Show("MULTIPLE Materials with Materialid=" + plngMaterialID + " have been deleted. Ensure that the Materials are correct.", MyCompName, MessageBoxButtons.OK);
                UserForm__SHUTDOWN(qReturn.qReturn_Finish);
            }

            if (mqrtQReturn != qReturn.qReturn_Finish && mqrtQReturn != qReturn.qReturn_Cancel)
            {
                UserForm__SHUTDOWN(qReturn.qReturn_Cancel);
            }
            Setup_Document_MBR(true);
            return mqrtQReturn;
        }
        public void QClose()
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }

        private void FormEditMaterials_FormClosing(object sender, FormClosingEventArgs e)
        {
            _formManager.Wizard_Terminate(activeDocument, mblnCalledWizard, mstrProtectionPassword);            

        }
        private void FormEditMaterials_FormClosed(object sender, FormClosedEventArgs e)
        {

            CommonService.CleanUpBackgroundWordProcesses();

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
