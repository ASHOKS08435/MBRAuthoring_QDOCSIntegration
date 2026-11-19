using MBRAuthoring.Utilities.CustomAttributes;
using Microsoft.Office.Interop.Word;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace MBRAuthoring.Models
{
    public class TblStepText
    {
        public string StepName { get; set; }
        public string StepText { get; set; }
        public string IsFound { get; set; }
    }

    public class DBDetails
    {
        public string SERVERNAME { get; set; }
        public string DBNAME { get; set; }
        public string VERSION { get; set; }
        public UserDataResult UserDetails { get; set; }
    }

    public class DBConnectionDetails
    {
        public DateTime ConnectStartTime { get; set; }
        public DateTime ConnectEndTime { get; set; }
        public int DatabaseRequestsIn { get; set; }
        public int DatabaseRequestsOut { get; set; }
        public int DatabaseConnections { get; set; }

    }

    public class TblOpUnit
    {
        public int OpUnitID { get; set; }
        public string OpUnit { get; set; }
        public string OpUnitType { get; set; }
        public string OUBuilding { get; set; }
        public int? WorkCenterOpUnitID { get; set; }
        public int? MaterialOpUnitID { get; set; }

        public TblOpUnit(int OpUnitID, string OpUnit, string OpUnitType)
        {
            this.OpUnitID = OpUnitID;
            this.OpUnit = OpUnit;
            this.OpUnitType = OpUnitType;
        }
        public TblOpUnit()
        {

        }
    }

    public class User
    {
        public int? UserID { get; set; }
        public int? UserTypeID { get; set; }
        public string Type { get; set; }
        public int? OpUnitID { get; set; }
        public string OpUnitName { get; set; }
    }

    public class TblLookupType
    {
        public string LookupType { get; set; }
        public string LookupTypeDefault { get; set; }
        public string LookupText { get; set; }
        public string LookupAttributeDefinition { get; set; }
    }

    public class TblLookupData
    {
        public string LookupText { get; set; }
        public string LookupAttributeValue { get; set; }
        public string LookupAttributeDefinition { get; set; }
    }

    public class UserDataResult
    {
        public User User { get; set; }
        public string ErrMsg { get; set; }
    }

    public class OpUnitDocType
    {
        public int DocTypeID { get; set; }
        public string DocType { get; set; }
        public string DocTypeAbb { get; set; }
        public string DocumentPath { get; set; }
        public string ProjectNameUsed { get; set; }
        public string DocIsModule { get; set; }
        public string DocHasModule { get; set; }
        public string DocHasMaterial { get; set; }
        public string DocHasWorkCenter { get; set; }
        public string DocHasWorkCenterAlt { get; set; }
        public string DocHasProduct { get; set; }
        public string DocHasProperty { get; set; }
        public string DocPermitsQuantitySufficient { get; set; }
        public int? DocNameIncludes_OpUnit { get; set; }
        public int? DocNameIncludes_DocTypeAbb { get; set; }
        public int? DocNameIncludes_WorkCenter { get; set; }
        public int? DocNameIncludes_VerMajor { get; set; }
        public int? DocNameIncludes_VerMinor { get; set; }
        public int? DocNameIncludes_FreeText { get; set; }
        public int? DocNameIncludes_DocBaseName { get; set; }
        public int? DocNameIncludes_ProductIdent { get; set; }

        [NonDBBacked]
        public bool Includes_DocBaseName
        {
            get
            {
                return DocNameIncludes_DocBaseName.HasValue && DocNameIncludes_DocBaseName.Value != 0;
            }
        }
        [NonDBBacked]
        public bool Includes_ProductIdent
        {
            get
            {
                return DocNameIncludes_ProductIdent.HasValue && DocNameIncludes_ProductIdent.Value != 0;
            }
        }
        [NonDBBacked]
        public bool Includes_FreeText
        {
            get
            {
                return DocNameIncludes_FreeText.HasValue && DocNameIncludes_FreeText.Value != 0;
            }
        }
        [NonDBBacked]
        public bool Includes_WorkCenter
        {
            get
            {
                return DocNameIncludes_WorkCenter.HasValue && DocNameIncludes_WorkCenter.Value != 0;
            }
        }
    }

    public class Material
    {
        public int MaterialID { get; set; }
        public string MaterialTitle { get; set; }
        public string MaterialDisplay { get; set; }
        public string ItemCode { get; set; }
    }

    public class BuildDocumentNameResult
    {
        public string ErrorReturn { get; set; }
        public string ErrorNumber { get; set; }
        public string DocumentReturn { get; set; }
        public string BaseNameReturn { get; set; }
    }

    public class DocBaseNameDetails
    {
        public string DocumentBaseName { get; set; }
        public string DocumentTitle { get; set; }
    }

    #region FileOpen
    public class DocumentDetails
    {
        public int DocID { get; set; }
        public string DocumentName { get; set; }
        public int DocLocation { get; set; }
        public string DocPasswordRead { get; set; }
        public string DocPasswordWrite { get; set; }
        public string DocumentPath { get; set; }
        public string DocumentTitle { get; set; }
        public string DisplayDocumentName { get; set; }

    }

    public class Document_Details
    {
        public string DocumentName { get; set; }
        public string DocumentTitle { get; set; }
        public string ItemCode { get; set; }
        public string MaterialName { get; set; }

    }

    #endregion

    public class DocumentImageLog
    {
        // public int DocID { get; set; }
        public DateTime LogDate { get; set; }
        public string LogNTUser { get; set; }
        public int DocEndPageNumber { get; set; }
        public int DocSectionsCount { get; set; }
        public int DocParagraphsCount { get; set; }
        public int DocWordsCount { get; set; }
        public int DocCharactersCount { get; set; }
        public int DocTablesCount { get; set; }
        public int DocCustomPropertiesCount { get; set; }
        public int DocFieldsCount { get; set; }

    }

    //Document Verify Wizard 
    public class DocumentValidateDetails
    {
        public string MasterDocumentName { get; set; }
        public string MasterLocationName { get; set; }
        public string ModuleDocumentName { get; set; }
        public string ModuleLocationName { get; set; }
        public string PassFail { get; set; }
    }

    public class ClsDocument
    {
        const string MyCompName = "ClsDocument";
        Word.Document docThis;
        public bool mblnUsable;
        public string mstrDocumentName;
        public string mstrOpUnit;
        public bool mblnTrackChangesReq;
        public int mintOpUnitID;
        public string mstrOpUnitType;
        public int mintDocTypeID;
        public string mstrDocType;
        public string mstrDocTypeAbb;
        public int mlngMaterialID;
        public string mstrMaterialName;
        public string mstrMaterialEDPNUM;
        public string mManufacturing_Product_Code;
        public string mNDC_Code;
        public string mProduct_Identification_Code;
        public string mRCS_Sales_Code;
        public string mGlobal_Trade_ID_Number;
        //added for GTIN - PANNEJ02
        public string mVialLabelGlobalTradeIDNumber;
        public string mstrDocumentPath;
        public bool mblnPendingUpdate;
        public string mDocIsModule;
        public string mProjectNameUsed;
        public string mDocPermitsQuantitySufficient;
        public string mDocHasMaterial;
        public string mDocHasModule;
        public string mDocHasProduct;
        public string mDocHasProperty;

        public ClsDocument()
        {
            // rstDocument = Nothing;
            docThis = null;
            mblnUsable = false;
        }
        
        public int DocID { get; set; }
        public string DocumentName { get; set; }
        public string DocumentBaseName { get; set; }
        public string DocumentTitle { get; set; }
        public int DocTypeID
        {
            get
            {
                if (!mblnUsable)
                {
                    //var ex = new Exception();
                    //ex.Data["errNum"] = "41060";
                    //ex.Data["errDesc"] = ModChangeHistory.gstrAppAcronym + " Document Object has not been .Read";
                    //throw ex;
                }
                return mintDocTypeID;

            }
            set
            {
                mintDocTypeID = value;
            }
        }
        public int OpUnitID
        {
            get
            {
                if (!mblnUsable)
                {
                    //var ex = new Exception();
                    //ex.Data["errNum"] = "46001";
                    //ex.Data["errDesc"] = ModChangeHistory.gstrAppAcronym + " Document Object has not been .Read";
                    //throw ex;
                }
                return mintOpUnitID;
            }
            set
            {
                mintOpUnitID = value;
            }
        }
        public int DocLocation { get; set; }
        public string DocLocationName { get; set; }
        public string DocLocationComment { get; set; }
        public string ProjectName { get; set; }
        public int WorkCenterID { get; set; }
        public int MaterialID { get; set; }
        public string StepCode { get; set; }
        public string Free_Text { get; set; }
        public int VersionMajor { get; set; }
        public int VersionMinor { get; set; }
        public string VariantText { get; set; }
        public int SupercedesDocID { get; set; }
        public int FormatVersion { get; set; }
        public int ScaledValues { get; set; }
        public string DocPasswordRead { get; set; }
        public string DocPasswordWrite { get; set; }
        public string DocPasswordProtect { get; set; }
        public string DocTrackChanges { get; set; }
        
    }

    public class DocumentType
    {
        public int DocTypeID { get; set; }
        public string DocType { get; set; }
        public string DocTypeAbb { get; set; }
        public string RecordType { get; set; }
    }
    public class ProductDetail
    {
        public int MaterialID { get; set; }
        public int OpUnitID { get; set; }
        public string ItemCode { get; set; }
        public string MaterialTitle { get; set; }
        public string MaterialDisplay { get; set; }
        public string MaterialCategory { get; set; }
        public string Product_Identification_Code { get; set; }
        public string Manufacturing_Product_Code { get; set; }
        public string NDC_Code { get; set; }
        public string RCS_Sales_Code { get; set; }
        public string Global_Trade_ID_Number { get; set; }
        //added GTIN PANNEj02
        public string VialLabelGlobalTradeIDNumber { get; set; }
        public string Itemcode { get; set; }
        public string MaterialEDP { get; set; }
        public string MICK_Material_Item_Code { get; set; }
    }

    public class DocumentInfo
    {
        public ClsDocument GetClsDocument { get; set; }
        public ProductDetail GetProductDetail { get; set; }
        public TblOpUnit GetTblOpUnit { get; set; }
        public DocumentType GetDocumentType { get; set; }
        public OpUnitDocType GetOpUnitDocType { get; set; }
    }

    public class DocumentSelectAttribute
    {
        public string DocumentName { get; set; }
        public int DocumentID { get; set; }
        public bool DocumentCleaningRecord { get; set; }
        public bool DocumentMasterRecord { get; set; }
        public string DocumentOpUnit { get; set; }
        public int DocumentOpUnitID { get; set; }
        public string DocumentOpUnitType { get; set; }
        public string DocumentDocType { get; set; }
        public string DocumentDocTypeAbb { get; set; }
        public int DocumentDocTypeID { get; set; }
        public string DocumentProtectionPassword { get; set; }
        public bool DocumentTrackChangesReq { get; set; }

    }

    public class DocImageLog
    {
        public DateTime LogDate { get; set; }
        public string LogNTUser { get; set; }
        public string LogHost { get; set; }
        public string LogProgram { get; set; }
        public int DocBookmarkCount { get; set; }
        public int DocCharactersCount { get; set; }
        public int DocFieldsCount { get; set; }
        public int DocParagraphsCount { get; set; }
        public int DocCustomPropertiesCount { get; set; }
        public int DocEndPageNumber { get; set; }
        public int DocSectionsCount { get; set; }
        public int DocSentencesCount { get; set; }
        public int DocTablesCount { get; set; }
        public int DocWordsCount { get; set; }
    }

    #region  EditMaterials
    public class MaterialCategoryDetails
    {
        public string MaterialCategory { get; set; }
        public string MaterialCategoryDescription { get; set; }
        public string Material_ItemCodeFormat { get; set; }
        public string Product_ItemCodeFormat { get; set; }
        public string MaterialHasImprint { get; set; }
        public string MaterialHasPerUnitData { get; set; }
        public string MaterialHasBOMText { get; set; }

    }

    public class OpUnitMaterial
    {
        public int MaterialID { get; set; }
        public string MaterialTitle { get; set; }
        public string ItemCode_MaterialTitle { get; set; }
        public string MaterialTitle_ItemCode { get; set; }
        public string ItemCode { get; set; }
        public string MaterialCategory { get; set; }
        public string MICK_Material_Item_Code { get; set; }

    }

    public class MaterialUnit
    {
        public string MaterialUOM { get; set; }
    }
    public class MaterialUsageDetails
    {
        public string MaterialUsage { get; set; }
        public bool MaterialUsageAllowAmount { get; set; }
        public string MaterialUsageDescription { get; set; }

    }

    public class DocMaterial_SelectDetails
    {
        public string MaterialTitle { get; set; }
        public int MaterialID { get; set; }
        public string MaterialDisplay { get; set; }
        public string ItemCode { get; set; }
        public string MaterialAmountDisplay { get; set; }
        public string MaterialUOM { get; set; }
        public int MaterialSequence { get; set; }
        public string MaterialUsage { get; set; }
        public bool QuantitySufficient { get; set; }
        public string MaterialCategory { get; set; }
        public string MaterialImprint { get; set; }
        public string MaterialPerUnitData { get; set; }
        public string MaterialBOMText { get; set; }

    }

    #endregion

    public class SelctedProperty
    {
        public string PropertyName { get; set; }
        public string PropertyDescription { get; set; }
        public string PropertyDataType { get; set; }
        public string PropertyDataTypeName { get; set; }
        public int PropertyIndexMax { get; set; }
        public string PropertyUOMLookupType { get; set; }
        public string PropertyUsage { get; set; }
        public string PropertyLookupType { get; set; }
        public string PropertyCategory { get; set; }
        public int PropertyTextMaxLen { get; set; }
    }

    
    public class BuildDocumentNameDetail
    {
        public string ErrorReturn { get; set; }
        public string DocumentReturn { get; set; }
        public string BaseNameReturn { get; set; }
        public int ErrorNumber { get; set; }

    }

    public class BuildDocumentNameVersionDetail
    {
        public string ErrorReturn { get; set; }
        public int ErrorNumber { get; set; }
        public string DocumentReturn { get; set; }
        public string BaseNameReturn { get; set; }
        public int VersionMajor { get; set; }

    }
    public class FileCreateDetail
    {
        public string ErrorReturn { get; set; }
        public int ErrorNumber { get; set; }
        public string DocumentNameReturn { get; set; }
        public string DocumentBaseReturn { get; set; }
        public int DocumentIDReturn { get; set; }
        public string DocPasswordReadReturn { get; set; }
        public string DocPasswordWriteReturn { get; set; }
        public string DocPasswordProtectReturn { get; set; }

    }

    public class FileCopyVersDetail
    {
        public string DocumentReturn { get; set; }
        public string ErrorReturn { get; set; }

    }
    public class FileCopyDiff_Detail
    {
        public string DocumentReturn { get; set; }
        public string ErrorReturn { get; set; }

    }
}
