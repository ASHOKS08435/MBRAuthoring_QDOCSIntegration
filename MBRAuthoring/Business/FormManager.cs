using MBRAuthoring.DataAccess;
using MBRAuthoring.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DataTable = System.Data.DataTable;

namespace MBRAuthoring.Business
{
    public interface IFormManager
    {
        List<TblOpUnit> Fill_OpUnit();
        List<OpUnitDocType> Fill_OpUnit_DocType(int mintOpUnitIDCreate, string mstrOpUnitTypeCreate);

        List<Material> Fill_OpUnit_Product(int mintOpUnitIDCreate, string mstrOpUnitTypeCreate, bool isProductOrderByNumber);

        List<DocumentDetails> Fill_Document(int opUnitID, int docTypeID, bool IncludeLocked, bool IncludeAuthoring, bool IncludeApproved, bool IncludeArchived, bool IncludeBlocked);

        List<DocumentImageLog> Fill_DocImageLog(int docId);

        int GetMaterialAmountScale(string amount); //In ModDataBase

        string GetDocumentVerifyPath();

        string AppStatus_Select(string keyName);

        List<DocumentValidateDetails> Fill_Document_Validation(int DocID);

        string VerifyProperty(int DocId);

        int SetDocumentToLockedState(int DocID);
        List<MaterialCategoryDetails> Fill_MaterialCategory(int OpUnitID);
        List<OpUnitMaterial> Fill_OpUnit_Material(int mintOpUnitID, string mstrOpUnitType, string orderBy, string materialCategory_Value);
        List<MaterialUnit> Fill_MaterialUnit();
        List<MaterialUsageDetails> Fill_MaterialUsage();
        void NotifyUser_UnIncludedMaterials(int docID);
        DocImageLog DocImageLog_SelectLast(int docID);
        string DocMaterial_UsedByProperty(int docId, int materailId);
        void DocMaterial_InsertUpdate(int mDocID, int materialDBID, string strAmount, string strMaterialUOM, int intMain, int docMaterialQuantitySufficient, string materialUsage, string materialImprint, string materialPerUnitData, string materialBOMText);
        void DocMaterial_Delete(int mDocID, int sequence);
        void Wizard_BookMarkRemoval(Document pdocDocument);
        List<DocMaterial_SelectDetails> Fill_Document_Material(int mDocIDNegative);
        Document_Details FillDocument_Details(int DocID);
        int Document_UpdateTitle(int docID, string docTitle);
        bool DocProperty_Apply(int docID, int propertyIndex, string propertyName, string propertyValue, string propertyUOM, int materialID, int workCenterID);
        List<SelctedProperty> Property_Select(string docIsModule, string propertyCategory);

        DataTable Document_Property_Select(string docIsModule, string propertyCategory);
        List<DocBaseNameDetails> Fill_DocBaseName(int docTypeID, bool isDocOrderByBaseName);
        BuildDocumentNameDetail BuildDocumentName(int opUnitID, int docTypeID, long materialID, long workCenterID, string freeText, string docBaseName);
        BuildDocumentNameVersionDetail BuildDocumentNameVersion(int docID, string variantText);

        FileCreateDetail File_Create(int opUnitID, long docTypeID, string documentTitle, long workCenterID, long materialID, string freeText, string docBaseName, long superDocID);

        FileCopyVersDetail File_CopyVers(long oldDOCId, string variantText);

        FileCopyDiff_Detail File_CopyDiff(long oldDOCId, int opUnitID, int docTypeID, int materialID, string docBaseName);
        DocumentType DocType_Select(int docTypeID);
        void DocumentSaved_DB_Record(int docId, Document document);
        void Wizard_Terminate(Document pdocDocument, bool pblnCalledWizard, string pstrProtectionPassword);
    }

    public class FormManager : IFormManager
    {
        public FormManager()
        {

        }
       
        public List<TblOpUnit> Fill_OpUnit()
        {
            List<TblOpUnit> opUnits = new List<TblOpUnit>();
            opUnits = DatabaseHelper.OpUnit_SelectALL();

            return opUnits;
        }
        public List<OpUnitDocType> Fill_OpUnit_DocType(int mintOpUnitIDCreate, string mstrOpUnitTypeCreate)
        {
            List<OpUnitDocType> opUnitDocTypeList = new List<OpUnitDocType>();
            opUnitDocTypeList = DatabaseHelper.OpUnitDocType_Select(mintOpUnitIDCreate);

            return opUnitDocTypeList;

        }

        
        public List<Material> Fill_OpUnit_Product(int mintOpUnitIDCreate, string mstrOpUnitTypeCreate, bool isProductOrderByNumber)
        {
            return DatabaseHelper.OpUnitProduct_SelectALL(mintOpUnitIDCreate, isProductOrderByNumber);
        }

        
        public BuildDocumentNameDetail BuildDocumentName(int opUnitID, int docTypeID, long materialID, long workCenterID, string freeText, string docBaseName)
        {
            return DatabaseHelper.BuildDocumentName(opUnitID, docTypeID, materialID, workCenterID, freeText, docBaseName);
        }

        
        public BuildDocumentNameVersionDetail BuildDocumentNameVersion(int docID, string variantText)
        {
            return DatabaseHelper.BuildDocumentNameVersion(docID, variantText);
        }

        //public (List<object>, Dictionary<string, object>) LoadBuildDocumentName(int opUnitID, int docTypeID, int materialID, string moduleID)
        //{
        //    object docBaseName;
        //    if (moduleID != null) { docBaseName = moduleID; } else { docBaseName = DBNull.Value; }
        //    SqlParameter[] parameters = new SqlParameter[]
        //    {
        //        new SqlParameter("@OpUnitID",SqlDbType.Int){SqlValue =opUnitID },
        //        new SqlParameter("@DocTypeID",System.Data.SqlDbType.Int){SqlValue = docTypeID},
        //        new SqlParameter("@MaterialID",System.Data.SqlDbType.Int){SqlValue = materialID},
        //        new SqlParameter("@VersionMajor",System.Data.SqlDbType.Int){SqlValue =1 },
        //        new SqlParameter("@VersionMinor",System.Data.SqlDbType.Int){SqlValue = 0},
        //        new SqlParameter("@WorkCenterID",System.Data.SqlDbType.Int){SqlValue = 0},
        //        new SqlParameter("@FreeText",System.Data.SqlDbType.VarChar){SqlValue = "-"},
        //        new SqlParameter("@DocBaseName",System.Data.SqlDbType.VarChar){SqlValue = docBaseName},
        //        new SqlParameter("@return_value",System.Data.SqlDbType.Int){Direction = ParameterDirection.ReturnValue},
        //        new SqlParameter("@ErrorReturn",System.Data.SqlDbType.VarChar,300){SqlValue = "",Direction = ParameterDirection.InputOutput},
        //        new SqlParameter("@DocumentReturn",System.Data.SqlDbType.VarChar,100){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput},
        //        new SqlParameter("@BaseNameReturn",System.Data.SqlDbType.VarChar,100){SqlValue = DBNull.Value, Direction = ParameterDirection.InputOutput},
        //        new SqlParameter("@ErrorNumber",System.Data.SqlDbType.Int){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput},
        //    };

        //    //Func<IDataReader, BuildDocumentNameResult> mapper = reader =>
        //    //                {
        //    //                    BuildDocumentNameResult buildDocumentNameResult = new BuildDocumentNameResult();
        //    //                    buildDocumentNameResult.ErrorReturn = reader.GetString(reader.GetOrdinal("ErrorReturn"));
        //    //                    buildDocumentNameResult.ErrorNumber = reader.GetString(reader.GetOrdinal("ErrorNumber"));
        //    //                    buildDocumentNameResult.DocumentReturn = reader.GetString(reader.GetOrdinal("DocumentReturn"));
        //    //                    buildDocumentNameResult.BaseNameReturn = reader.GetString(reader.GetOrdinal("BaseNameReturn"));
        //    //                    return buildDocumentNameResult;
        //    //                };

        //    var resultSet = DatabaseHelper.ExecuteStoredProcedure<object>("SP_BuildDocumentName", parameters, null, out var outputValues, false);
        //    //Dictionary<string,object> item2 =  resultSet.Item2;
        //    //if (item2.ContainsKey("@ErrorReturn"))
        //    //{
        //    //    string err = item2["@ErrorReturn"].ToString();
        //    //}

        //    return resultSet;
        //}


        public List<DocBaseNameDetails> Fill_DocBaseName(int docTypeID, bool isDocOrderByBaseName)
        {
            return DatabaseHelper.DocBaseName_SelectAll(docTypeID, isDocOrderByBaseName);
        }
        

        public FileCreateDetail File_Create(int opUnitID, long docTypeID, string documentTitle, long workCenterID, long materialID, string freeText, string docBaseName, long superDocID)
        {
            return DatabaseHelper.File_Create(opUnitID, docTypeID, documentTitle, workCenterID, materialID, freeText, docBaseName, superDocID);
        }

        

        public FileCopyVersDetail File_CopyVers(long oldDOCId, string variantText)
        {
            return DatabaseHelper.File_CopyVers(oldDOCId, variantText);
        }

        

        public FileCopyDiff_Detail File_CopyDiff(long oldDOCId, int opUnitID, int docTypeID, int materialID, string docBaseName)
        {
            return DatabaseHelper.File_CopyDiff(oldDOCId, opUnitID, docTypeID, materialID, docBaseName);
        }

        

        #region FileOpen
        public List<DocumentDetails> Fill_Document(int opUnitID, int docTypeID, bool IncludeLocked, bool IncludeAuthoring, bool IncludeApproved, bool IncludeArchived, bool IncludeBlocked)
        {
            //return DatabaseHelper.Document_SelectByOpUnit(opUnitID, docTypeID, IncludeLocked, IncludeAuthoring, IncludeApproved, IncludeArchived, IncludeBlocked);
            List<DocumentDetails> selectedDocDetails = DatabaseHelper.Document_SelectByOpUnit(opUnitID, docTypeID, IncludeLocked, IncludeAuthoring, IncludeApproved, IncludeBlocked,IncludeArchived);
            return selectedDocDetails.Where(x => x != null).Select(x =>
            {
                var documentDetail = new DocumentDetails
                {
                    DocID = x.DocID,
                    DocumentName = x.DocumentName,
                    DocLocation = x.DocLocation,
                    DocPasswordRead = x.DocPasswordRead,
                    DocPasswordWrite = x.DocPasswordWrite,
                    DocumentPath = x.DocumentPath,
                    DocumentTitle = x.DocumentTitle,
                    DisplayDocumentName = x.DocumentName
                };

                switch (x.DocLocation)
                {
                    case CommonService.gintDocLocation_LOCK:
                        documentDetail.DisplayDocumentName += "(Locked)";
                        break;
                    case CommonService.gintDocLocation_AUTH:
                        documentDetail.DisplayDocumentName += "";
                        break;
                    case CommonService.gintDocLocation_APPR:
                        documentDetail.DisplayDocumentName += "(APPR)";
                        break;
                    case CommonService.gintDocLocation_ARCH:
                        documentDetail.DisplayDocumentName += "(ARCHIVE)";
                        break;
                    case CommonService.gintDocLocation_BLCK:
                        documentDetail.DisplayDocumentName += "(BLOCKED)";
                        break;
                    default:
                        documentDetail.DisplayDocumentName += "(Loc=" + documentDetail.DocLocation + ")";
                        break;
                }

                return documentDetail;
            }).ToList();
        }

        public DocImageLog DocImageLog_SelectLast(int docID)
        {
            return DatabaseHelper.DocImageLog_SelectLast(docID);
        }

        public Document_Details FillDocument_Details(int DocID)
        {
            return DatabaseHelper.Document_Details(DocID);
        }

        public int Document_UpdateTitle(int docID, string docTitle)
        {
            return DatabaseHelper.Document_UpdateTitle(docID, docTitle);
        }
        #endregion

        public DocumentType DocType_Select(int docTypeID)
        {
            return DatabaseHelper.DocType_Select(docTypeID);
        }

        public List<DocumentImageLog> Fill_DocImageLog(int DocId)
        {
            return DatabaseHelper.DocImageLog_Select(DocId);
        }

        public string GetDocumentVerifyPath()
        {
            return DatabaseHelper.OpUnitDocType_DocumentPath_Select();
        }

        public string AppStatus_Select(string cDocVerifyAuth)
        {
            return DatabaseHelper.AppStatus_Select(cDocVerifyAuth);
        }

        public List<DocumentValidateDetails> Fill_Document_Validation(int DocID)
        {
            return DatabaseHelper.DocumentValidate_Select(DocID);
        }

        public string VerifyProperty(int DocId)
        {
            return DatabaseHelper.VerifyProperty(DocId);
        }

        public int SetDocumentToLockedState(int DocID)
        {
            return DatabaseHelper.SetDocumentToLockedState(DocID);
        }
        #region EditMaterials

        public List<MaterialCategoryDetails> Fill_MaterialCategory(int OpUnitID)
        {
            return DatabaseHelper.MaterialCategory_Select(OpUnitID);
        }

        public List<OpUnitMaterial> Fill_OpUnit_Material(int mintOpUnitID, string mstrOpUnitType, string orderBy, string materialCategory_Value)
        {
            return DatabaseHelper.OpUnitMaterial_SelectALL(mintOpUnitID, materialCategory_Value);
        }

        public List<MaterialUnit> Fill_MaterialUnit()
        {
            return DatabaseHelper.MaterialUnit_Select();
        }

        public List<MaterialUsageDetails> Fill_MaterialUsage()
        {
            return DatabaseHelper.MaterialUsage_Select();
        }

        public string DocMaterial_UsedByProperty(int docId, int materailId)
        {
            return DatabaseHelper.DocMaterial_UsedByProperty(docId, materailId);
        }

        public void DocMaterial_InsertUpdate(int mDocID, int materialDBID, string strAmount, string strMaterialUOM, int intMain, int docMaterialQuantitySufficient, string materialUsage, string materialImprint, string materialPerUnitData, string materialBOMText)
        {
            DatabaseHelper.DocMaterial_InsertUpdate(mDocID, materialDBID, strAmount, strMaterialUOM, intMain, docMaterialQuantitySufficient, materialUsage, materialImprint, materialPerUnitData, materialBOMText);

        }
        public void DocMaterial_Delete(int mDocID, int sequence)
        {
            DatabaseHelper.DocMaterial_Delete(mDocID, sequence);
        }
        public int GetMaterialAmountScale(string amount) //In ModDataBase
        {
            string strDecimal;
            int intDecimalPos;
            string SearchChar;
            amount = amount.Trim();
            SearchChar = ".";
            intDecimalPos = amount.IndexOf(SearchChar, 1);
            if (intDecimalPos > 0)
            {
                return amount.Length - intDecimalPos;
            }
            else
            {
                return 0;
            }
        }

        public List<DocMaterial_SelectDetails> Fill_Document_Material(int docId)
        {
            bool includeZeroAmount = false;
            if (docId > 0)
            {
                includeZeroAmount = true;
            }
            return DatabaseHelper.DocMaterial_Select(docId, includeZeroAmount);
        }
        #endregion


        public void NotifyUser_UnIncludedMaterials(int docID)
        {
            DatabaseHelper.NotifyUser_UnIncludedMaterials(docID);
            //    // Todo DatabaseHelper.NotifyUser_UnIncludedMaterials(docID);
            //    SqlParameter[] parameters = new SqlParameter[]
            //   {
            //        new SqlParameter("@DOcID",SqlDbType.Int){SqlValue =docID }
            //   };
            //    var resultSet = DatabaseHelper.ExecuteStoredProcedure<object>("usp_GetUnincludedMaterial", parameters, null, out var outputValues, false);
        }

        // Removing extra bookmarks if exists 
        public void Wizard_BookMarkRemoval(Document pdocDocument)
        {
            const string cstrSTART = "Start";
            const string cstrNOTESTART = "NoteStart";
            const string cstrEND = "End";
            try
            {
                if (pdocDocument.Bookmarks.Exists(cstrSTART))
                {
                    pdocDocument.Bookmarks[cstrSTART].Delete();

                }
                if (pdocDocument.Bookmarks.Exists(cstrNOTESTART))
                {
                    pdocDocument.Bookmarks[cstrNOTESTART].Delete();
                }
                if (pdocDocument.Bookmarks.Exists(cstrEND))
                {
                    pdocDocument.Bookmarks[cstrEND].Delete();
                }
            }
            catch (Exception ex)
            {

            }
        }

        public bool DocProperty_Apply(int docID, int propertyIndex, string propertyName, string propertyValue, string propertyUOM, int materialID, int workCenterID)
        {
            return DatabaseHelper.DocProperty_Apply(docID, propertyIndex, propertyName, propertyValue, propertyUOM, materialID, workCenterID);
        }

        public List<SelctedProperty> Property_Select(string docIsModule, string propertyCategory)
        {
            return DatabaseHelper.Property_Select(docIsModule, propertyCategory);
        }

        public DataTable Document_Property_Select(string docIsModule, string propertyCategory)
        {
            return DatabaseHelper.Document_Property_Select(docIsModule, propertyCategory);
        }
        public void DocumentSaved_DB_Record(int docId, Document document)
        {
            DatabaseHelper.DocumentSaved_DB_Record(docId, document);
        }

        public void Wizard_Terminate(Document pdocDocument, bool pblnCalledWizard, string pstrProtectionPassword)
        {
            try
            {

                if (!pblnCalledWizard)
                {
                    Wizard_BookMarkRemoval(pdocDocument);
                    CommonService.DocumentProtection_Assure(pdocDocument, pstrProtectionPassword);
                }
            }
            catch (Exception ex)
            {

            }
        }


    }
    
}
