using MBRAuthoring.DataAccess;
using MBRAuthoring.Models;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
namespace MBRAuthoring.Business
{
    public enum TemplateNameFormat
    {
        tnfNameOnly = 5209,
        tnfPathOnly = 17832,
        tnfNameAndPath = 28516,
        tnfReview = 13309
    }

    public enum qRequest
    {
        qRequest_Display = 17932,
        qRequest_Validate = 29836,
        qRequest_Write = 23041,
        qRequest_Add = 13195,
        qRequest_Delete = 27431,
        qRequest_DBDownLoad = 8341
    }

    public enum qReturn
    {
        qReturn_Back = 28071,
        qReturn_Next = 25095,
        qReturn_Cancel = 16803,
        qReturn_Finish = 30821,
        qReturn_Collapse = 14209,
        qReturn_Valid = 24815,
        qReturn_Written = 9437
    }


    public enum QWriteWhat
    {
        qWriteCLR = 24183,
        qWriteMBRTable = 13091,
        qWriteMBRProperty = 29084,
        qWritePKG = 20048,
        qWriteCOM = 31672,
        qWriteDBUpdate = 4302,
        qWritePROCleanup = 27409,
        qWritePROMaterial = 27391,
        qWritePROList = 2271,
        qWriteWhatever = 18230
    }
    public enum StepTextType
    {
        qtAutoText = 7491,
        qtAutoTextNoKeep = 111468,
        qtAutoTextNoLine = 21465,
        qtContinue = 3897,
        qtCriticalBeforeStep = 14923,
        qtCriticalWithinStep = 9529,
        qtImmediate = 5304,
        qtMainDouble = 6241,
        qtMainSingle = 28794,
        qtMainStep = 8904,
        qtNormalSingle = 21374,
        qtNormalDouble = 27134,
        qtNormalDoubleNoKeep = 80192,
        qtNormalTriple = 24137,
        qtNoteBeforeStep = 7286,
        qtNoteWithinStep = 12208,
        qtMBRHeading = 5219,
        qtMBRStep = 13902,
        qtMBRSub = 29419,
        qtNoteWithinStepNoKeep = 13982,
        qtSubDouble = 8235,
        qtSubSingle = 11406,
        qtSubStep = 25702,
        qtSubStepNoKeep = 27170,
        qtSubSubDouble = 14694,
        qtSubSubSingle = 16439,
        qtSubSubStep = 9273,
        qtWarningBeforeStep = 26318,
        qtWarningWithinStep = 13842,
        qtWarningWithinStepNoKeep = 14293,
        qtZeroZero = 10101,
        qtBRReviewConfirm = 737
    }

    public enum NCWLocation
    {
        qNCWBeforeStep = 4921,
        qNCWWithinStep = 17142,
        qNCWWithinStepNoKeep = 31420
    }
    public enum QDisplay_FirstLast
    {
        QDisplay_First = 3402,
        QDisplay_Last = 21608
    }

    public struct QPS_Find
    {
        public string QFindText { get; set; }
        public string QFindFont { get; set; }
        public string QFindStyle { get; set; }
        public string QFindReplacementText { get; set; }
        public string QFindReplacementFont { get; set; }
        public string QFindReplacementStyle { get; set; }
        public bool QFindForward { get; set; }
        public int QFindWrap { get; set; }
        public bool QFindFormat { get; set; }
        public bool QFindMatchByte { get; set; }
        public bool QFindMatchCase { get; set; }
        public bool QFindMatchWholeWord { get; set; }
        public bool QFindMatchWildcards { get; set; }
        public bool QFindMatchSoundsLike { get; set; }
        public bool QFindMatchAllWordForms { get; set; }
    }
    public struct qpsEditMaterialParms
    {
        public int intEditMaterialRun;
        public int intEditMaterialValidated;
        public string[,] strEditMaterialSelected;
        public string strEditMaterialStandardYield;
        public string strEditMaterialStandardUOM;
        public string[] strEditMaterialYields;
        public string strEditMaterialYieldIgnore;
    }

    public struct qpsEditModuleParms
    {
        public int intEditModuleRun;
        public int intEditModuleValidated;
        public string[,] strEditModuleSelected;

    }

    public struct qpsEditPropertyParms
    {
        public int intEditPropertyRun;
        public int intEditPropertyValidated;
        public string[,] strEditPropertySelected;

    }

    class CommonService
    {
        const string myServiceName = "CommonService";
        public const string attachedTemplate = "AttachedTemplate";
        public static string templateFullName;
        public static string templateName;
        public static string templatePath;
        public static int intVarCount;
        public static int intVarLevel;
        public static bool blnAlreadyRun;
        public static Template docTemplate;
        public static StepVariable mvarStepVariable;
        public const int qRequest_Add = 13195;
        public const int qRequest_Delete = 27431;
        public const int qRequest_DBDownLoad = 8341;

        public const int gintDocLocation_LOCK = -1;
        public const int gintDocLocation_AUTH = 0;
        public const int gintDocLocation_APPR = 1;
        public const int gintDocLocation_ARCH = 2;
        public const int gintDocLocation_BLCK = 3;

        public const string cProductID = "ProductID";
        public const string cLotNum = "LotNbr";
        public const string cProductName = "ProductName";
        public const string cRecordType = "RecordType";

        public const int qdlColCount = 10;
        public static int mintSwapCount;

        public const int qdlSubMaterial_MULTI_Offset = 90;
        public const int qdlSubColMaterialImprint = 91;
        public const int qdlSubColMaterialPerUnitData = 92;
        public const int qdlSubColMaterialBOMText = 93;
        public const int qdlSubColMaterialCategory = 94;
        public const string Multi_Array_Separator = "\x01";
        public const string qdlMaterialAmount_qs_No_Amount = "q.s.";
        private static SwapPair[] amSwap = new SwapPair[10];
        public static string qdlMaterialAmount_qs_With_Amount = "q.s. to";

        public const int gintWordVersion2007 = 12;
        public static string gstrStepTextLatest;
        public static string gstrAutoTextLatest;
        public static string gstrWhichIsLatest;
        static bool NCWPrevious;
        public static bool gblnDebug;

        public const string gstrStyle_Normal = "Normal";
        public const string gstrStyle_MBR = "MBR";
        public const string gstrStyle_MBRNotice = "MBRNotice";
        public const string gstrStyle_MBRSteps = "MBRSteps";
        public const string gstrStyle_MBRSub = "MBRSub";
        public const string gstrStyle_MBRSubSub = "MBRSubSub";
        public const int qtWizardInitialization = 24061;
        public const int qtCheckedBy = 26041;
        public const string gstrStepText_MarkupStart = "[";
        public const string gstrStepText_MarkupEnd = "]";
        public const string gstrStepText_Bold = "!";
        public const string gstrStepText_Italic = "~";
        public const string gstrStepText_Underline = "_";
        //public const string gstrStepText_Superscript = ">";
        //public const string gstrStepText_Subscript = "<";
        public const string gstrStepText_Centered = "%";
        public const string gstrStepText_Color = "*";
        public const string gstrStepText_Color_Red = "R";
        public const string gstrStepText_Color_Green = "G";
        public const string gstrStepText_Color_Blue = "B";
        public const string gstrStepText_Color_Yellow = "Y";
        public const string gstrStepText_Color_Magenta = "M";
        public const string gstrStepText_Color_Cyan = "C";
        public const string gstrStepText_Fontsize = "@";
        public const string gstrStepText_Property = "&";
        public const string gstrStepText_Literal = "=";

        public static double dblTotalAutoText;
        public static int intCountAutoText;
        static Word.Range rngDebug = null;
        static string stepTextName_Prev;
        static int stepTextNameCount;
        public double dblTotalStepText;
        public int intCountStepText;
        MBRAuthDAL objMBRAuthDAL = new MBRAuthDAL();
        public const int qdpColPropertyValueRequired = 0;
        public const int qdpColPropertyName = 1;
        public const int qdpColPropertyIndex = 2;
        public const int qdpColPropertyInstance = 3;
        public const int qdpColPropertyValue = 4;
        public const int qdpColPropertyUOM = 5;
        public const int qdpColPropertyValueDisplay = 6;

        public static string gstrUserName;
        public static int gintUserID;
        public static bool gblnUserType_ADMIN;
        public static bool gblnUserType_REVIEWER;
        public const string userType_Admin = "ADMIN";
        public const string userType_REVIEWER = "REVIEWER";
        public const string cOrderNum = "OrderNum";
        public const string cDocumentTitle = "DocumentTitle";
        public const string cFileName = "FileName";
        public const string cChangeNum = "ChangeNum";
        public const string cPrintDate = "PrintDate";
        public const string cPrintTime = "PrintTime";
        public const string cPrintPages = "PrintPages";
        public const string cPage = "Page";
        public const string gstrIntroEnd = "IntroEnd";
        public const string cManufacturing_Product_Code = "Manufacturing_Product_Code";
        public const string cNDC_Code = "NDC_Code";
        public const string cProduct_Identification_Code = "Product_Identification_Code";
        public const string cRCS_Sales_Code = "RCS_Sales_Code";
        public const string cGlobal_Trade_ID_Number = "Global_Trade_ID_Number";
        public const string cPrintPages_ValueDefault = "n";
        public const string cChangeNum_ValueDefault = "TBD";
        // public static string gstrAppVersionMajor; //gstrAppVersionMajor --> In ModChange history, Nut for now declates in commonservice

        public const int gintCOPYRIGHT = 169;
        public const int gintREGISTERED = 174;
        public const int gintDEGREES = 176;
        public const int gintGAMMA = 181;
        private const string gstrStepText_Superscript = ">";
        private const string gstrStepText_Subscript = "<";

        private static Word.Application wordApplication;

        public static string Get_Template_Name(Word.Application wordApp, TemplateNameFormat templateNameType)
        {
            const string myMethodName = "Get_Template_Name";
            //string strTemplateFullCandidate;
            //string strTemplateNameCandidate;
            //string strTemplatePathCandidate;

            Word.Templates templates = wordApp.Templates;
            try
            {
                if (templateNameType == TemplateNameFormat.tnfReview)
                {
                    return DatabaseHelper.AppStatus_Select(attachedTemplate);
                }

               
                    foreach (Word.Template tpt in templates)
                    {
                        if (tpt.Name.ToUpper().Contains("MBRAUTHORING.DOTX"))
                        {
                            if (String.IsNullOrEmpty(templateFullName))
                            {
                                templateFullName = tpt.FullName;
                                templateName = tpt.Name;
                                templatePath = tpt.Path;
                                docTemplate = tpt;
                            }
                            else
                            {
                                CommonService.Error_Handler("0", "Multiple Templates: '" + templateFullName + "' and '" + tpt.FullName + "'", myServiceName, myMethodName, false);
                                return "UNKNOWN";
                            }
                        }

                    }
               
                switch (templateNameType)
                {
                    case TemplateNameFormat.tnfNameOnly:
                        return templateName;
                    case TemplateNameFormat.tnfPathOnly:
                        return templatePath;
                    case TemplateNameFormat.tnfNameAndPath:
                        return templateFullName;
                    default:
                        CommonService.Error_Handler("0", "Invalid parameter='" + templateNameType + "'; expecting Name, Path, or FullName.", myServiceName, myMethodName, false);
                        return "INVALIDPARAMETER";
                }
                //string strTemplateFullCandidate;
                //DatabaseHelper.AppStatus_Select(attachedTemplate);

            }
            catch (Exception ex)
            {

                if (ex.Data["errNum"] != null)
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), myServiceName, myMethodName, false);
                }
                else
                {
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, myServiceName, myMethodName, false);
                }

            }
            return templateFullName;
        }
               

        public static void Error_Handler(string plngErr, string pstrErr, string pstrForm, string pstrProc, bool pblnLogToDatabase = true, int pintErl = 0)
        {
            string strMessage;
            string strProc;
            Globals.ThisDocument.ActiveDocument.Application.System.Cursor = WdCursorType.wdCursorNormal;

            strMessage = "An error has occurred. Please delete any text this wizard may have " + "\r\n" + "entered into your document and try again. If that does not work, " + "\r\n" + "please call the system administrator." + "\r\n" + "\r\n" + "Please have the following information available when you call:" + "\r\n" + "\r\n" + "User: " + "\t\t" + Environment.UserName + "\r\n" + "Date: " + "\t\t" + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt") + "\r\n" + "File: " + "\t\t" + Globals.ThisDocument.ActiveDocument.Name + "\r\n" + "Wizard: " + "\t\t" + pstrForm + "\r\n" + "Procedure: " + "\t" + pstrProc + "\r\n" + "Error Number: " + "\t" + plngErr + "\r\n" + "Error: " + "\t\t" + pstrErr;

            strProc = pstrProc;
            if (pintErl != 0) // It is a line num. kind of ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' '));
            {
                strMessage = strMessage + "\r\n" + "Line: " + "\t" + "\t" + pintErl;
                strProc = strProc + "." + pintErl;
            }
            MessageBox.Show(strMessage, "Error");

            if (pblnLogToDatabase)
            {
                DatabaseHelper.ErrorLog_Insert(plngErr.ToString(), pstrErr, pstrForm, strProc);

            }
        }

        public static string StrFormatter(string strIn, string strNullConvert = "~", string strTabConvert = "\t", int MaxLength = -1)
        {
            string strChar;
            long lngScan;
            string strOut = "";
            string strOutMaximum = "";

            for (lngScan = 0; lngScan < strIn.Length; lngScan++)
            {
                strChar = strIn.Substring((int)lngScan, 1);

                switch (strChar)
                {
                    case "'":
                        strOut += "''";
                        break;
                    case "\0":
                        strOut += strNullConvert;
                        break;
                    case "\t":
                        strOut += strTabConvert;
                        break;
                    default:
                        strOut += strChar;
                        break;
                }

                if (lngScan < MaxLength || MaxLength == -1)
                {
                    strOutMaximum = strOut;
                }
            }

            if (MaxLength == -1)
            {
                return strOut;
            }
            else
            {
                return strOutMaximum;
            }
        }

        public static void Style_Copy_All(Document openedDocument, IDocumentManager docMgr = null)
        {
            Get_Style(openedDocument, "MBR", docMgr);
            Get_Style(openedDocument, "MBRNotice", docMgr);
            Get_Style(openedDocument, "MBRSteps", docMgr);
            Get_Style(openedDocument, "MBRSub", docMgr);
            Get_Style(openedDocument, "MBRSubSub", docMgr);
        }

        public static string TranslatePassword(object varPassword)
        {
            string strPassword = string.Empty;

            if (varPassword == null)
            {
                return string.Empty;
            }

            strPassword = varPassword.ToString();

            if (string.IsNullOrEmpty(strPassword) || strPassword == " ")
            {
                return string.Empty;
            }

            return strPassword;
        }
               

        public static Style Get_Style(Document docCurrent, string styleName, IDocumentManager docMgr = null)  // NewCode3
        {
            const string MyProcName = "Get_Style";
            Template tptSource = null;
            string strDocumentName = docCurrent.Name;
            string strPasswordProtection = string.Empty;
            bool blnProtect = false;

            try
            {
                return docCurrent.Styles[styleName];
            }
            catch (Exception)
            {
                try
                {
                    tptSource = Get_Template_ForStyles(docCurrent.Application);

                    if (docCurrent.ProtectionType == WdProtectionType.wdNoProtection)
                    {
                        blnProtect = false;
                    }
                    else
                    {
                        blnProtect = true;
                        strPasswordProtection = docMgr.Document_SelectPasswordProtect(docCurrent.Name);
                        DocumentProtection_Off(docCurrent, strPasswordProtection);
                    }

                    docCurrent.Application.OrganizerCopy(
                        Source: tptSource.FullName,
                        Destination: docCurrent.FullName,
                        Name: styleName,
                        Object: WdOrganizerObject.wdOrganizerObjectStyles);

                    if (blnProtect)
                        DocumentProtection_On(docCurrent, strPasswordProtection);

                    return docCurrent.Styles[styleName];
                }
                catch (Exception)
                {
                    if (blnProtect)
                        DocumentProtection_On(docCurrent, strPasswordProtection);

                    MessageBox.Show(
                        $"Unable to find Style '{styleName}', using 'Normal'",
                        MyProcName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);

                    return docCurrent.Styles["Normal"];
                }
            }
        }

        public static Template Get_Template_ForStyles(Word.Application wordApp)  // NewCode3
        {
            Template template = null;

            foreach (Template temp in wordApp.Templates)
            {
                if (temp.FullName.Contains("MBR"))
                {
                    template = temp;
                    break;
                }
            }

            return template;
        }

        public static void DocumentProtection_On(Word.Document pdocProtectedDocument, string pstrProtectionPassword, int pintProtectionType = (int)Word.WdProtectionType.wdAllowOnlyFormFields, string pstrDocumentName = " ")
        {
            if (!DocumentProtection_Required(pdocProtectedDocument, pstrDocumentName))
            {
                return;
            }
            if (pdocProtectedDocument.ProtectionType != Word.WdProtectionType.wdNoProtection)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41060";
                ex.Data["errDesc"] = "'" + pdocProtectedDocument.Name + "' ALREADY PROTECTED";
                throw ex;
            }
            pdocProtectedDocument.Protect((Word.WdProtectionType)pintProtectionType, false, pstrProtectionPassword, false, false); //010625 TypeMismatch
        }

        public static void DocumentProtection_Off(Word.Document pdocProtectedDocument, string pstrProtectionPassword, string pstrDocumentName = "")
        {
            string strDocumentName;

            if (pstrDocumentName == "")
            {
                strDocumentName = pdocProtectedDocument.Name;
            }
            else
            {
                strDocumentName = pstrDocumentName;
            }
            if (!DocumentProtection_Required(pdocProtectedDocument, strDocumentName))
            {
                return;
            }
            if (pdocProtectedDocument.ProtectionType == Word.WdProtectionType.wdNoProtection)
            {
                return;
            }
            pdocProtectedDocument.Unprotect(pstrProtectionPassword);

        }

        private static bool DocumentProtection_Required(Word.Document pdocProtectedDocument, string pstrDocumentName = "")
        {
            const String MyProcName = "DocumentProtection_Required";
            return true;
        }

        public static void DocumentProtection_Assure(Word.Document pdocProtectedDocument, string pstrProtectionPassword)
        {
            const string MyProcName = "DocumentProtection_Assure";
            // This logic Assures that DocumentProtection is ON for a document.
            //This Protection includes a Password, which prevents the user from easily turning off the protection.
            if (!DocumentProtection_Required(pdocProtectedDocument)) // NO protection work needed
            {
                return;
            }

            if (string.IsNullOrEmpty(pstrProtectionPassword.Trim())) //NO PASSWORD --- cannot protect the doc
            {
                return;
            }

            // CHECK whether the document is currently Protected
            if (pdocProtectedDocument.ProtectionType == Word.WdProtectionType.wdNoProtection)
            {
                pdocProtectedDocument.Protect(Word.WdProtectionType.wdAllowOnlyFormFields, false, pstrProtectionPassword);

            }
            else if (pdocProtectedDocument.ProtectionType == Word.WdProtectionType.wdAllowOnlyFormFields)
            {
                //OK: wdAllowOnlyFormFields is the desired protection
                // no implementation on VB
            }
            else
            {

                var ex = new Exception();
                ex.Data["errNum"] = "41061";
                ex.Data["errDesc"] = myServiceName + "." + " DocumentProtection_Required - '" + pdocProtectedDocument.Name + "' has unexpected ProtectionType = '" + pdocProtectedDocument.ProtectionType + "'";
                throw ex;
            }


        }


        internal static void Find_Setup(Find fndProcessingSteps)
        {
            fndProcessingSteps.ClearFormatting();
            fndProcessingSteps.Replacement.ClearFormatting();
            fndProcessingSteps.Replacement.Text = "";
            fndProcessingSteps.Forward = true;
            fndProcessingSteps.Wrap = WdFindWrap.wdFindStop;
            fndProcessingSteps.Format = fndProcessingSteps.MatchByte = fndProcessingSteps.MatchCase = fndProcessingSteps.MatchWholeWord = fndProcessingSteps.MatchWildcards = fndProcessingSteps.MatchSoundsLike = fndProcessingSteps.MatchAllWordForms = false;
        }

        internal static void Find_Restore(Find fndProcessingSteps, QPS_Find qpfFindSaveArea)
        {
            fndProcessingSteps.Text = qpfFindSaveArea.QFindText;
            if (!string.IsNullOrEmpty(qpfFindSaveArea.QFindFont)) { fndProcessingSteps.Font.Name = qpfFindSaveArea.QFindFont; }
            if (!string.IsNullOrEmpty(qpfFindSaveArea.QFindStyle)) { fndProcessingSteps.set_Style(qpfFindSaveArea.QFindStyle); }
            fndProcessingSteps.Replacement.Text = qpfFindSaveArea.QFindReplacementText;
            if (!string.IsNullOrEmpty(qpfFindSaveArea.QFindReplacementFont)) { fndProcessingSteps.Replacement.Font.Name = qpfFindSaveArea.QFindReplacementFont; }
            if (!string.IsNullOrEmpty(qpfFindSaveArea.QFindReplacementStyle)) { fndProcessingSteps.Replacement.set_Style(qpfFindSaveArea.QFindReplacementStyle); }
            fndProcessingSteps.Forward = qpfFindSaveArea.QFindForward;
            fndProcessingSteps.Wrap = (WdFindWrap)qpfFindSaveArea.QFindWrap;
            fndProcessingSteps.Format = qpfFindSaveArea.QFindFormat;
            fndProcessingSteps.MatchByte = qpfFindSaveArea.QFindFormat;
            fndProcessingSteps.MatchCase = qpfFindSaveArea.QFindMatchCase;
            fndProcessingSteps.MatchWholeWord = qpfFindSaveArea.QFindMatchWholeWord;
            fndProcessingSteps.MatchWildcards = qpfFindSaveArea.QFindMatchWildcards;
            fndProcessingSteps.MatchSoundsLike = qpfFindSaveArea.QFindMatchSoundsLike;
            fndProcessingSteps.MatchAllWordForms = qpfFindSaveArea.QFindMatchAllWordForms;

        }

        public static QPS_Find Find_Save(Find fndProcessingSteps)
        {
            QPS_Find qpf = new QPS_Find
            {
                QFindText = fndProcessingSteps.Text,
                QFindFont = fndProcessingSteps.Font.Name,
                QFindStyle = fndProcessingSteps.get_Style(),
                QFindReplacementText = fndProcessingSteps.Text = "",
                QFindReplacementFont = fndProcessingSteps.Font.Name,
                QFindReplacementStyle = fndProcessingSteps.Replacement.get_Style(),
                QFindForward = fndProcessingSteps.Forward,
                QFindWrap = ((int)fndProcessingSteps.Wrap),
                QFindFormat = fndProcessingSteps.Format,
                QFindMatchByte = fndProcessingSteps.MatchByte,
                QFindMatchCase = fndProcessingSteps.MatchCase,
                QFindMatchWholeWord = fndProcessingSteps.MatchWholeWord,
                QFindMatchWildcards = fndProcessingSteps.MatchWildcards,
                QFindMatchSoundsLike = fndProcessingSteps.MatchSoundsLike,
                QFindMatchAllWordForms = fndProcessingSteps.MatchAllWordForms
            };
            return qpf;
        }


        public static bool SwapInvolvement(string swapName)
        {
            int intScan;
            return amSwap.Take(mintSwapCount).Any(sp => sp.SPFrom == swapName || sp.SPTo == swapName);
        }

        public static void SwapRequest(string strSwapFrom, string strSwapTo)
        {
            const string MyProcName = "SwapRequest";
            if (SwapInvolvement(strSwapFrom))
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41030";
                ex.Data["errDesc"] = "'" + strSwapFrom + "' ALREADY IN SWAP";
                throw ex;
                return;
            }
            if (SwapInvolvement(strSwapTo))
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41030";
                ex.Data["errDesc"] = "'" + strSwapTo + "' ALREADY IN SWAP";
                throw ex;
                return;
            }
            if (mintSwapCount >= amSwap.Length)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41030";
                ex.Data["errDesc"] = "TOO MANY SWAPS";
                throw ex;
                return;
            }

            amSwap[mintSwapCount] = new SwapPair
            {
                SPFrom = strSwapFrom,
                SPTo = strSwapTo

            };
            mintSwapCount = mintSwapCount + 1;

        }

        public static void SwapClear()
        {
            mintSwapCount = 0;
        }

        public static void SwapExecute(Document activeDocument)
        {
            Word.Range range = activeDocument.Range();

            for (int intScan = 1; intScan <= mintSwapCount; intScan++)
            {
                Find_Setup(range.Find);
                range.Find.Text = amSwap[intScan - 1].SPFrom;
                range.Find.Replacement.Text = amSwap[intScan - 1].SPTo;
                range.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            }

        }

        public bool IsPropertyRequired(string PropertyInstance, System.Data.DataTable dtReqProperty)
        {
            bool IsPropertyRequired = false;
            const string MyProcName = "IsPropertyRequired";
            try
            {

                if (dtReqProperty.Rows.Count == 0)
                {
                    IsPropertyRequired = false;
                    return IsPropertyRequired;
                }

                for (int i = 0; i <= dtReqProperty.Rows.Count - 1; i++)
                {

                    if (PropertyInstance == dtReqProperty.Rows[i]["PropertyInstance"].ToString())
                    {
                        IsPropertyRequired = true;
                        MessageBox.Show("'" + PropertyInstance + "' is a Required Property." + "\r\n" + "REQUIRED Properties cannot be deleted.", "Required Property", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return IsPropertyRequired;
                    }

                }
            }
            catch
            {

                //ToDo MsgBox("Error (" + Information.Err.Number + "):" + Information.Err.Description, Constants.vbCritical, MyCompName + "." + MyProcName);
            }

            return IsPropertyRequired;

        }

        public static Word.Table Table_Find_MaterialMBR(Word.Document pdocCurrent)
        {
            return Table_Find(pdocCurrent, "Material List");

        }

        public static Word.Table Table_Find_Property(Word.Document pdocCurrent)
        {
            Word.Table tblFindProperty = Table_Find(pdocCurrent, "Property List");



            if (tblFindProperty == null)
            {
                tblFindProperty = Table_Find(pdocCurrent, "Property Name", "Property Value");
            }

            return tblFindProperty;
        }

        internal static void Table_Delete(Table ptblToDelete, Document pdocCurrent, bool pblnSelect = true, string pstrProtectionPassword = "")
        {
            bool blnTrackRevisions;
            bool blnDocumentProtected;
            if (pdocCurrent.ProtectionType == WdProtectionType.wdNoProtection)
            {
                blnDocumentProtected = false;
            }
            else
            {
                DocumentProtection_Off(pdocCurrent, pstrProtectionPassword);
                blnDocumentProtected = true;
            }
            if (pblnSelect)
            {
                ptblToDelete.Select();
            }
            if (pdocCurrent.TrackRevisions)
            {
                pdocCurrent.TrackRevisions = false;
                blnTrackRevisions = true;
            }
            else
            {
                blnTrackRevisions = false;
            }
            ptblToDelete.Delete();
            if (blnDocumentProtected)
            {
                DocumentProtection_On(pdocCurrent, pstrProtectionPassword);
            }
        }

        
        public static Table Table_Find(Document activeDocument, string cell1Contents, string cell2Contents = "", string cell3Contents = "")
        {
            foreach (Table table in activeDocument.Tables)
            {
                if (IsTableMatch(table, cell1Contents, cell2Contents, cell3Contents))
                {
                    return table;
                }
            }
            return null;
        }

        private static bool IsTableMatch(Table table, string cell1Contents, string cell2Contents, string cell3Contents)
        {
            if (table.Rows.Count == 0 || table.Columns.Count == 0)
            {
                return false;
            }

            string cell1Text = Table_CellText(table, 1, 1);
            if (cell1Text != cell1Contents)
            {
                return false;
            }

            if (!string.IsNullOrEmpty(cell2Contents))
            {
                string cell2Text = Table_CellText(table, 1, 2);
                if (cell2Text != cell2Contents)
                {
                    return false;
                }
            }

            if (!string.IsNullOrEmpty(cell3Contents))
            {
                string cell3Text = Table_CellText(table, 1, 3);
                if (cell3Text != cell3Contents)
                {
                    return false;
                }
            }

            return true;
        }

        private static string GetCellText(Table table, int row, int column)
        {
            try
            {
                Cell cell = table.Cell(row, column);
                string text = cell.Range.Text;
                return text.Substring(0, text.Length - 2);
            }
            catch
            {
                return "UNMATCHABLE!";
            }
        }

        public static string Table_CellText(Table table, int RowIndex, int ColumnIndex)
        {
            //string strCellText;
            //strCellText = table.Cell(RowIndex, ColumnIndex).Range.Text;
            //strCellText = strCellText.Substring(0, strCellText.Length - 2);
            //return strCellText;
            try
            {
                Cell cell = table.Cell(RowIndex, ColumnIndex);
                string text = cell.Range.Text;
                return text.Substring(0, text.Length - 2);
            }
            catch
            {
                return "UNMATCHABLE!";
            }
        }

        public static void AutoText_Insert(string autotextEntryName, ref int tableDataRowFirst, int tableDataRows = -1, bool defineHeader = false, Word.Range rangeAutoText = null, Word.Table tableAutoText = null, float indent = -1, StepTextType stepTextType = StepTextType.qtAutoText, bool IspropertyDefine = false, Word.Application wordApp = null, Document document = null)
        {
            const string MyProcName = "AutoText_Insert";
            Word.Template tptAutoText = null;
            Word.Range range = null;
            Word.Table table;
            Word.Find find;
            Word.Field field;
            DateTime timeStart;
            int loopLimiter;
            int rowsCurrent;
            int rowsToAdd;
            QPS_Find qfdSave;
            int rangeStart = 0;
            int rangeEnd;
            string foundText;
            int retryCount = 0;
            string strVarName;
            string strVarValueExt = "";
            bool blnBold = false;
            int intSymbol = 0;
            timeStart = DateTime.Now;
            double dblElapsed;

            if (tableDataRows != -1 & (tableDataRows < 1))
            {
                CommonService.Error_Handler("1001", "Invalid TableDataRows=" + tableDataRows, myServiceName, MyProcName);
                return;
            }
            if (string.IsNullOrEmpty(autotextEntryName) && tableDataRows < 1)
            {
                CommonService.Error_Handler("1001", "Called AutoTextEntryName='" + autotextEntryName + "' and TableDataRows=" + tableDataRows, myServiceName, MyProcName);
                return;
            }
            wordApplication = wordApp;
            Word.Selection selection = wordApplication.Selection;
            if (selection.Information[WdInformation.wdWithInTable])
            {
                CommonService.Error_Handler("1001", "Called Within Table", myServiceName, MyProcName);
                return;
            }
            if (defineHeader)
            {
                CommonService.Error_Handler("1001", "DefineHeader code not written!", myServiceName, MyProcName);
                return;
            }
            CommonService.gstrAutoTextLatest = autotextEntryName;
            gstrWhichIsLatest = "AutoText";
            if (stepTextType > 0)
            {
                StepTextSpacing(stepTextType, locator: MyProcName);
            }
            if (!string.IsNullOrEmpty(autotextEntryName))
            {

                tptAutoText = Get_Template_ForAutoText();
                range = AutoText_Insert_Entry(tptAutoText, autotextEntryName);
                Paragraph_KeepTogether(range);
                rangeAutoText = range;
                if (selection.Information[WdInformation.wdWithInTable])
                {
                    CommonService.Error_Handler("1001", "After AutoText.Insert, Selection is WithinTable", myServiceName, MyProcName);
                    return;
                }
            }
            else
            {
                rangeAutoText = null;
            }
            if (defineHeader)
            {
                loopLimiter = 5;
                while (!selection.Information[WdInformation.wdWithInTable] || loopLimiter == 0)
                {
                    selection.MoveLeft(WdUnits.wdCharacter, 1);
                    loopLimiter--;
                }
                if (loopLimiter == 0)
                {
                    CommonService.Error_Handler("1001", "(DefineHeader) No table before Selection", myServiceName, MyProcName);
                    return;
                }
                rowsCurrent = selection.Information[WdInformation.wdMaximumNumberOfRows];
                if (rowsCurrent < 2)
                {
                    CommonService.Error_Handler("1001", "Cannot DefineHeader for RowCount=" + rowsCurrent, myServiceName, MyProcName);
                    return;

                }
                selection.MoveUp(WdUnits.wdLine, 1);
                loopLimiter = rowsCurrent;
                while (selection.Information[WdInformation.wdStartOfRangeRowNumber] == 1 || loopLimiter == 0)
                {
                    selection.MoveUp(WdUnits.wdLine, 1, WdMovementType.wdExtend);
                    loopLimiter--;
                }
                if (loopLimiter == 0)
                {
                    CommonService.Error_Handler("1001", "(DefineHeader) Before Row 1", myServiceName, MyProcName);
                    return;
                }
                if (!selection.Information[WdInformation.wdWithInTable])
                {
                    CommonService.Error_Handler("1001", "(DefineHeader) Before Row 1", myServiceName, MyProcName);
                    return;
                }
                selection.Rows.HeadingFormat = 1;
                selection.Collapse();
                Table_Exit();
            }
            if (tableDataRows == -1)
            {
                tableDataRowFirst = -1;
            }
            else
            {
                if (range.Tables.Count == 0)
                {
                    CommonService.Error_Handler("1001", "AutoText " + autotextEntryName + "does NOT contain a TABLE!", myServiceName, MyProcName);
                    return;
                }
                table = range.Tables[1];
                tableDataRowFirst = table.Rows.Count;
                if (tableDataRowFirst != 1)
                {
                    rowsToAdd = tableDataRows - 1;
                    for (int i = rowsToAdd; i >= 1; i--)
                    {
                        table.Rows.Add();
                    }
                    Table_Exit();
                }

            }
            if (indent != -1)
            {
                table = range.Tables[1];

                table.Rows.SetLeftIndent(ConvertInchesToPoints(indent), RulerStyle: WdRulerStyle.wdAdjustNone);
                //Todo:
                //  table.Rows.SetLeftIndent(InchesToPoints(psngIndent), wdAdjustNone);
            }
            qfdSave = Find_Save(range.Find);
            rangeStart = range.Start;
            rangeEnd = range.End;

            //Need to check
            do
            {
                range.Start = rangeStart;
                range.End = rangeStart;
                find = range.Find;
                Find_Setup(find);
                //Todo:
                //fnd.Text = "[" & gstrStepText_MarkupStart & "]*[" & gstrStepText_MarkupEnd & "]"
                find.MatchWildcards = true;
                if (!find.Execute())
                {
                    break;
                }
                foundText = range.Text;
                if (foundText.StartsWith(gstrStepText_MarkupStart) || foundText.EndsWith(gstrStepText_MarkupStart))
                {
                    if (retryCount < 5)
                    {
                        retryCount = retryCount + 1;
                        //GoTo Retry;
                    }
                }
                if (range.Start > rangeEnd || range.End > rangeEnd)
                {
                    break;
                }
                strVarName = foundText.Substring(1, foundText.Length - 1);
                StepVarRetrieveExt(strVarName, strVarValueExt, blnBold, intSymbol);
                range.Text = strVarValueExt;
                if (blnBold)
                {
                    range.Font.Bold = 1;
                }
                rangeStart = range.End;
                rangeEnd = rangeEnd - foundText.Length + strVarValueExt.Length;

            } while (true);

            rangeAutoText.Start = rangeStart;
            rangeAutoText.End = rangeEnd;
            if (rangeAutoText.Tables.Count > 0)
            {
                tableAutoText = rangeAutoText.Tables[1];
            }
            if (IspropertyDefine)
            {
                foreach (Word.Field fld in rangeAutoText.Fields)
                {

                    Field_PropertyDefine(document, fld.Code.Text);
                    fld.Update();
                    Field_ApplyShading(fld);
                }
            }
            Find_Restore(find, qfdSave);
            dblElapsed = (DateTime.Now - timeStart).TotalSeconds;
            dblTotalAutoText = dblTotalAutoText + dblElapsed;
            intCountAutoText = intCountAutoText + 1;

        }

        public static void StepVarSymbol(string pstrVarName, string pstrVarValueExt, int pintSymbol)
        {
            const string MyProcName = "StepVarSymbol";
            if (pstrVarName.Length > 0 && pstrVarName[0] != '$')
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41083";
                ex.Data["errDesc"] = "ERROR: Expecting leading '$' in '" + pstrVarName + "'";
                throw ex;
            }

            string strVarName = pstrVarName.Substring(1);

            if (int.TryParse(strVarName, out pintSymbol))
            {
                pstrVarValueExt = "";
            }
            else
            {
                switch (strVarName.ToUpper())
                {
                    case "LAMBDA":
                        pstrVarValueExt = "";
                        pintSymbol = -3988;
                        break;
                    case "MICRO":
                        pstrVarValueExt = "";
                        pintSymbol = -3987;
                        break;
                    case "DEGREE":
                        pstrVarValueExt = ((char)gintDEGREES).ToString();
                        pintSymbol = 0;
                        break;
                    case "GAMMA":
                        pstrVarValueExt = ((char)gintGAMMA).ToString();
                        pintSymbol = 0;
                        break;
                    case "COPYRIGHT":
                        pstrVarValueExt = ((char)gintCOPYRIGHT).ToString();
                        pintSymbol = 0;
                        break;
                    case "REGISTERED":
                        pstrVarValueExt = ((char)gintREGISTERED).ToString();
                        pintSymbol = 0;
                        break;
                    default:
                        pstrVarValueExt = "*UNKNOWN: " + pstrVarName + "*";
                        pintSymbol = 0;
                        break;
                }
            }
        }

        public static void StepVarRetrieveExt(string pstrVarName, string pstrVarValueExt, bool pblnVarBold, int pintVarSymbol)
        {
            const string MyProcName = "StepVarRetrieveExt";

            try
            {
                mvarStepVariable.Name = pstrVarName;
                StepVarManager.ManageVariable("get", mvarStepVariable);

                pstrVarValueExt = StepVarValueExt(mvarStepVariable.Value, mvarStepVariable.IsScalable, mvarStepVariable.IsReplaceable);
                pblnVarBold = mvarStepVariable.IsBold;
                pintVarSymbol = mvarStepVariable.Symbol;
            }
            catch
            {
                pstrVarValueExt = "? " + pstrVarName + " UNKNOWN?";
                pblnVarBold = false;
            }
        }

        public static string StepVarValueExt(string pstrVarValue, bool pblnScalable, bool pblnReplaceable)
        {
            if (pblnReplaceable)
            {
                return "[" + pstrVarValue + "]";
            }
            else if (pblnScalable)
            {
                return "{" + pstrVarValue + "}";
            }
            else
            {
                return pstrVarValue;
            }
        }
        private static Template Get_Template_ForAutoText()
        {
            Word.Template temp = null;
            foreach (Word.Template template in wordApplication.Templates)
            {
                if (template.FullName.Contains("MBR"))
                {
                    return temp = template;
                }
            }
            return temp;
        }

        private static Range AutoText_Insert_Entry(Word.Template tptAutoText, string autotextEntryName)
        {
            const string MyProcName = "AutoText_Insert_Entry";
            try
            {
                int version;
                string appVersion = wordApplication.Version;
                version = int.Parse(appVersion.Substring(0, appVersion.IndexOf('.')));
                if (version >= CommonService.gintWordVersion2007)
                {
                    return BuildingBlock(tptAutoText, autotextEntryName);

                }
                else
                {
                    return AutoTextEntries(tptAutoText, autotextEntryName);
                }
            }
            catch (COMException e)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "15941";
                ex.Data["errDesc"] = "Autotext '" + autotextEntryName + "' not found in " + tptAutoText.Name;
                throw ex;
            }
            catch (Exception e)
            {
                throw e;
            }



        }

        private static Range AutoTextEntries(Template tptAutoText, string autotextEntryName)
        {
            return tptAutoText.AutoTextEntries[autotextEntryName].Insert(wordApplication.Selection.Range, true);
        }

        public static string ErrorProcLatest(string pstrProcName)
        {
            return pstrProcName + "(StepText=" + gstrStepTextLatest + ",AutoText=" + gstrAutoTextLatest + ",Latest=" + gstrWhichIsLatest + ")";
        }

        private static Range BuildingBlock(Word.Template ptptAutoText, string autotextEntryName)
        {
            try
            {
                //rng = wordApplication.Selection.Range;
                BuildingBlock entry = ptptAutoText.BuildingBlockEntries.Item(autotextEntryName);
                return entry.Insert(wordApplication.Selection.Range, true);
            }
            catch
            {
                return wordApplication.Selection.Range;
            }
        }

        public static void StepTextSpacing(StepTextType psttStepTextType, bool pblnCalledByNotice = false, string locator = "")
        {
            const string MyProcName = "StepTextSpacing";
            bool blnNCWPrevious = false;
            bool blnNCWCurrent = false;
            int intBlankLinesNeededBefore = -1;
            int intBlankLinesFoundBefore = 0;
            int intParaToBeAddedBefore = 0;
            int intParaToBeDeletedBefore = 0;
            int intBlankLinesNeededAfter = -1;
            int intBlankLinesFoundAfter = 0;
            int intParaToBeAddedAfter = 0;
            int intParaToBeDeletedAfter = 0;
            bool blnSelectionAtParaStart = false;
            bool blnKeepBefore = true;
            bool blnKeepAfter = false;
            string strStyleName = "";
            dynamic strPrevNonBlankParaStyle;
            Word.Paragraph parScan;
            Word.Range rngScan;
            Word.Range rngSelection;
            int intSectionIndexInitial;
            int intSectionIndexScan;
            DateTime datStart = DateTime.Now;
            double dblElapsed;
            double dblTotalStepTextSpacing = 0;
            int intCountStepTextSpacing = 0;

            if (pblnCalledByNotice)
            {
                switch (psttStepTextType)
                {
                    case StepTextType.qtWarningBeforeStep:
                        break;
                    case StepTextType.qtWarningWithinStep:
                        break;
                    case StepTextType.qtCriticalBeforeStep:
                        break;
                    case StepTextType.qtCriticalWithinStep:
                        break;
                    case StepTextType.qtNoteBeforeStep:
                        break;
                    case StepTextType.qtNoteWithinStep:
                        break;
                    default:
                        var ex = new Exception();
                        ex.Data["errNum"] = "41060";
                        ex.Data["errDesc"] = "UNEXPECTED STEPTEXTTYPE=" + psttStepTextType + " FOR 'Notice'";
                        throw ex;
                        break;

                }
                intBlankLinesNeededAfter = 1;
            }
            Table_Exit();
            switch (psttStepTextType)
            {
                case StepTextType.qtMainStep:
                    if (NCWPrevious)
                    {
                        intBlankLinesNeededBefore = 1;
                        blnKeepBefore = true;
                    }
                    else
                    {
                        intBlankLinesNeededBefore = 2;
                        blnKeepBefore = false;
                    }
                    strStyleName = gstrStyle_MBRSteps;
                    break;
                case StepTextType.qtMainDouble:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBR;
                    break;
                case StepTextType.qtMainSingle:
                    intBlankLinesNeededBefore = 0;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBR;
                    break;
                case StepTextType.qtSubStep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBRSteps;
                    break;
                case StepTextType.qtSubStepNoKeep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_MBRSteps;
                    break;
                case StepTextType.qtSubSingle:
                    intBlankLinesNeededBefore = 0;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBRSub;
                    break;
                case StepTextType.qtSubDouble:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBRSub;
                    break;
                case StepTextType.qtSubSubSingle:
                    intBlankLinesNeededBefore = 0;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBRSubSub;
                    break;
                case StepTextType.qtSubSubStep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBRSteps;
                    break;
                case StepTextType.qtSubSubDouble:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBRSubSub;
                    break;
                case (StepTextType)qtWizardInitialization:
                    intBlankLinesNeededBefore = 2;
                    intBlankLinesNeededAfter = 2;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_Normal;
                    break;
                case (StepTextType)qtCheckedBy:
                    intBlankLinesNeededAfter = 2;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_Normal;
                    break;
                case StepTextType.qtAutoText:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case StepTextType.qtAutoTextNoKeep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_Normal;
                    break;
                case StepTextType.qtAutoTextNoLine:
                    intBlankLinesNeededBefore = 0;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_Normal;
                    break;
                case StepTextType.qtWarningBeforeStep:
                    intBlankLinesNeededBefore = 2;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_MBRNotice;
                    blnNCWCurrent = true;
                    break;
                case StepTextType.qtWarningWithinStep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBRNotice;
                    blnNCWCurrent = true;
                    break;
                case StepTextType.qtWarningWithinStepNoKeep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_MBRNotice;
                    blnNCWCurrent = true;
                    break;
                case StepTextType.qtCriticalBeforeStep:
                    intBlankLinesNeededBefore = 2;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_MBRNotice;
                    blnNCWCurrent = true;
                    break;
                case StepTextType.qtCriticalWithinStep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBRNotice;
                    blnNCWCurrent = true;
                    break;
                case StepTextType.qtNoteBeforeStep:
                    intBlankLinesNeededBefore = 2;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_MBRNotice;
                    blnNCWCurrent = true;
                    break;
                case StepTextType.qtNoteWithinStep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBRNotice;
                    blnNCWCurrent = true;
                    break;
                case StepTextType.qtNoteWithinStepNoKeep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_MBRNotice;
                    blnNCWCurrent = true;
                    break;
                case StepTextType.qtNormalSingle:
                    intBlankLinesNeededBefore = 0;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case StepTextType.qtNormalDouble:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case StepTextType.qtNormalDoubleNoKeep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_Normal;
                    break;
                case StepTextType.qtNormalTriple:
                    intBlankLinesNeededBefore = 2;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case StepTextType.qtMBRStep:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = false;
                    strStyleName = gstrStyle_MBRSteps;
                    break;
                case StepTextType.qtMBRSub:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_MBRSteps;
                    break;
                case StepTextType.qtZeroZero:
                    intBlankLinesNeededBefore = 0;
                    intBlankLinesNeededAfter = 0;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case StepTextType.qtBRReviewConfirm:
                    intBlankLinesNeededBefore = 1;
                    intBlankLinesNeededAfter = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case 0:
                    intBlankLinesNeededBefore = 0;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case (StepTextType)1:
                    intBlankLinesNeededBefore = 1;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case (StepTextType)2:
                    intBlankLinesNeededBefore = 2;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case (StepTextType)3:
                    intBlankLinesNeededBefore = 3;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case (StepTextType)4:
                    intBlankLinesNeededBefore = 4;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case (StepTextType)5:
                    intBlankLinesNeededBefore = 5;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                case (StepTextType)6:
                    intBlankLinesNeededBefore = 6;
                    blnKeepBefore = true;
                    strStyleName = gstrStyle_Normal;
                    break;
                default:
                    var ex = new Exception();
                    ex.Data["errNum"] = "41089";
                    ex.Data["errDesc"] = "'" + psttStepTextType + "' invalid parameter for " + MyProcName;
                    throw ex;
                    break;
            }

            if (intBlankLinesNeededBefore < 0 && intBlankLinesNeededAfter < 0)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "87034";
                ex.Data["errDesc"] = "No parameter was specified";
                throw ex;
            }

            strPrevNonBlankParaStyle = "";

            if (intBlankLinesNeededBefore >= 0)
            {
                intBlankLinesFoundBefore = 0;
                parScan = wordApplication.Selection.Paragraphs[1];
                intSectionIndexInitial = parScan.Range.Sections[1].Index;
                blnSelectionAtParaStart = false;

                while (true)
                {
                    if (parScan == null) break;

                    rngScan = parScan.Range;
                    intSectionIndexScan = rngScan.Sections[1].Index;

                    if (intSectionIndexScan != intSectionIndexInitial ||
                        (rngScan.Sections[1].ProtectedForForms && wordApplication.ActiveDocument.ProtectionType != Word.WdProtectionType.wdNoProtection) ||
                        rngScan.Text.Length > 1)
                    {
                        rngSelection = wordApplication.Selection.Range;
                        if (rngSelection.Start == rngScan.Start)
                        {
                            blnSelectionAtParaStart = true;
                            parScan = parScan.Previous();
                        }
                        else
                        {
                            strPrevNonBlankParaStyle = (string)parScan.get_Style().NameLocal;
                            break;
                        }
                    }
                    else if (rngScan.Information[Word.WdInformation.wdWithInTable])
                    {
                        strPrevNonBlankParaStyle = (string)parScan.get_Style().NameLocal;
                        break;
                    }
                    else
                    {
                        intBlankLinesFoundBefore++;
                        parScan = parScan.Previous();
                    }
                }

                if (intBlankLinesNeededBefore >= intBlankLinesFoundBefore)
                {
                    intParaToBeAddedBefore = intBlankLinesNeededBefore - intBlankLinesFoundBefore + 1;
                    if (blnSelectionAtParaStart) intParaToBeAddedBefore--;

                    if (intParaToBeAddedBefore == 1)
                    {
                        Paragraph_KeepTogether(wordApplication.Selection.Range, blnKeepBefore);
                    }
                    QPS_Paragraph(Keep: blnKeepBefore, Count: intParaToBeAddedBefore, Locator: MyProcName + "(Bef)" + locator, style: gstrStyle_Normal);
                }
                else
                {
                    intParaToBeDeletedBefore = intBlankLinesFoundBefore - intBlankLinesNeededBefore - 1;
                    if (blnSelectionAtParaStart) intParaToBeDeletedBefore++;

                    for (int i = intParaToBeDeletedBefore; i >= 1; i--)
                    {
                        wordApplication.Selection.TypeBackspace();
                    }

                    parScan = wordApplication.Selection.Paragraphs[1];
                    parScan = parScan.Previous();
                    if (parScan != null)
                    {
                        Paragraph_KeepTogether(parScan.Range, blnKeepBefore);
                    }
                }

                if (psttStepTextType == StepTextType.qtMainStep)
                {
                    parScan = wordApplication.Selection.Paragraphs[1].Previous();
                    if (parScan.KeepWithNext == 1)
                    {
                        if (!blnNCWPrevious)
                        {
                            parScan.KeepWithNext = 0;
                        }
                    }
                    else
                    {
                        if (blnNCWPrevious)
                        {
                            parScan.KeepWithNext = 1;
                        }
                    }
                }
            }

            if (intBlankLinesNeededAfter >= 0)
            {
                parScan = wordApplication.Selection.Paragraphs[1].Next();
                while (true)
                {
                    if (parScan == null) break;

                    rngScan = parScan.Range;
                    intSectionIndexScan = rngScan.Sections[1].Index;

                    if (rngScan.Sections[1].ProtectedForForms && wordApplication.ActiveDocument.ProtectionType != Word.WdProtectionType.wdNoProtection)
                    {
                        break;
                    }
                    else if (parScan.Range.Text.Length > 1)
                    {
                        break;
                    }
                    else
                    {
                        intBlankLinesFoundAfter++;
                        parScan = parScan.Next();
                    }
                }

                if (intBlankLinesNeededAfter > intBlankLinesFoundAfter)
                {
                    intParaToBeAddedAfter = intBlankLinesNeededAfter - intBlankLinesFoundAfter;
                    QPS_Paragraph(Keep: blnKeepAfter, Count: intParaToBeAddedAfter, Locator: MyProcName + "(Aft)" + locator, style: gstrStyle_Normal);
                    wordApplication.Selection.MoveUp(Word.WdUnits.wdLine, intParaToBeAddedAfter);
                }
                else if (intBlankLinesNeededAfter < intBlankLinesFoundAfter)
                {
                    intParaToBeDeletedAfter = intBlankLinesFoundAfter - intBlankLinesNeededAfter;
                    for (int i = intParaToBeDeletedAfter; i >= 1; i--)
                    {
                        wordApplication.Selection.Delete();
                    }
                }
            }

            wordApplication.Selection.Paragraphs[1].set_Style(Get_Style(wordApplication.ActiveDocument, strStyleName));
            Paragraph_KeepTogether(wordApplication.Selection.Range, true);
            blnNCWPrevious = blnNCWCurrent;
            dblElapsed = (DateTime.Now - datStart).TotalSeconds;
            dblTotalStepTextSpacing += dblElapsed;
            intCountStepTextSpacing++;
        }
             

        private static void QPS_Paragraph(bool Keep, int Count = 1, string Locator = "", string style = "", bool TrailingBlank = false)
        {
            string strDebugMessage;
            Word.Selection selection = wordApplication.Selection;

            for (int loopCount = 1; loopCount <= Count; loopCount++)
            {
                selection.TypeParagraph();
                if (selection.Font.Bold == 1)
                {
                    selection.Font.Bold = 0;
                }
                if (!string.IsNullOrEmpty(style))
                {
                    selection.Paragraphs.set_Style(Get_Style(wordApplication.ActiveDocument, style));
                }
                Paragraph_KeepTogether(selection.Range, Keep);
                if (gblnDebug)
                {
                    if (Keep)
                    {
                        strDebugMessage = "K=T";
                    }
                    else
                    {
                        strDebugMessage = "K=F";
                    }
                    if (!string.IsNullOrEmpty(Locator))
                    {
                        strDebugMessage = strDebugMessage = " " + Locator;
                    }
                    if (!string.IsNullOrEmpty(style))
                    {
                        strDebugMessage = strDebugMessage = " " + style;
                    }
                    if (Count != 1)
                    {
                        strDebugMessage = strDebugMessage = " " + loopCount + " of " + Count;
                    }
                    DebugMark(strDebugMessage);

                }
            }
            if (TrailingBlank)
            {
                selection.TypeText("   ");
            }
        }


        private static void Paragraph_KeepTogether(Range range, bool keep = true)
        {
            if (range != null && range.Paragraphs.Count > 0)
            {


                ParagraphFormat paragraphFormat = range.ParagraphFormat;
                paragraphFormat.WidowControl = -1;
                paragraphFormat.KeepWithNext = keep ? -1 : 0;
                paragraphFormat.KeepTogether = -1;

            }
            else
            {
                throw new ArgumentException("The provided range does not contain any paragraphs.");
            }
        }
        //private static void Paragraph_KeepTogether(Range range, bool keep = true)
        //{
        //    ParagraphFormat paragraphFormat = range.ParagraphFormat;

        //    int a = range.Sections.Count;
        //    int b = range.Tables.Count;

        //    paragraphFormat.WidowControl = true ? 1 : 0;
        //    paragraphFormat.KeepWithNext = keep ? 1 : 0;
        //    paragraphFormat.KeepTogether = 1;
        //    paragraphFormat.PageBreakBefore = 0;
        //}

        public static void Table_Exit(Word.Range rangeStart = null)
        {
            int subLoopLimiter;
            Word.Range rangeWork;
            if (rangeStart != null)
            {
                rangeStart.Select();
            }
            subLoopLimiter = 2000;
            while (wordApplication.Selection.Information[WdInformation.wdWithInTable] && subLoopLimiter > 0)
            {
                wordApplication.Selection.MoveDown(WdUnits.wdLine, 1);
                subLoopLimiter--;
            }
            if (subLoopLimiter == 0)
            {
                CommonService.Error_Handler("1001", "Unable to MoveDown out of Table", myServiceName, "Table_Exit");

            }
        }
        private static void DebugMark(string strDebugMessage)
        {
            if (!gblnDebug)
            {
                return;
            }
            if (gblnDebug == null)
            {
                rngDebug = wordApplication.ActiveDocument.Range();
                rngDebug.Collapse(WdCollapseDirection.wdCollapseEnd);
                rngDebug.Text = "DEBUG " + DateTime.Now + "\r\n" + "Page.Line  Message..." + "\r\n";
                rngDebug.Collapse(WdCollapseDirection.wdCollapseEnd);
            }
            rngDebug.Text = wordApplication.Selection.Information[WdInformation.wdActiveEndPageNumber] + "." + wordApplication.Selection.Information[WdInformation.wdFirstCharacterLineNumber] + " " + strDebugMessage + "\r\n";
            rngDebug.Collapse(WdCollapseDirection.wdCollapseEnd);
        }

        // merge
        public static void TypeStepText(StepTextType stepTextType, string stepTextName, string stepName = "", bool boldAll = false, bool useNameAsText = false, bool keep = true, bool keepMainStep = false) //added document param to access.
        {
            const string myProcName = "TypeStepText";
            //string strStepTextName_Prev; // static
            int intStepTextNameCount; // static
            bool blnBoldAll = boldAll;
            string strStepText;
            DateTime timeStart = DateTime.Now;
            double dblElapsed;

            if (stepTextName == stepTextName_Prev)
            {
                CommonService.gstrStepTextLatest = stepTextName + " #" + stepTextNameCount;
                stepTextNameCount++;
            }
            else
            {
                stepTextName_Prev = stepTextName;
                gstrStepTextLatest = stepTextName;
                stepTextNameCount = 1;
            }
            gstrWhichIsLatest = "StepText";
            Globals.ThisDocument.ActiveDocument.Application.StatusBar = "GENERATING: " + gstrStepTextLatest;
            if (useNameAsText)
            {
                strStepText = stepTextName;
            }
            else
            {
                strStepText = Get_Step_Text(stepTextName);
                if (string.IsNullOrEmpty(strStepText)) { Globals.ThisDocument.buildCommandBar = false; return; }
            }
            if (stepTextType == StepTextType.qtMainStep)
            {
                if (string.IsNullOrEmpty(stepName))
                {
                    var ex = new Exception();
                    ex.Data["errNum"] = "41091";
                    ex.Data["errDesc"] = "StepName required for MAIN Step.";
                    throw ex;
                    return;
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(stepName))
                {
                    var ex = new Exception();
                    ex.Data["errNum"] = "41092";
                    ex.Data["errDesc"] = "StepName '" + stepName + "' not allowed for non-MAIN Step.";
                    throw ex;
                    return;
                }
            }

            Document document = Globals.ThisDocument.ActiveDocument;
            wordApplication = document.Application;
            Selection selection = document.Application.Selection;
            if (selection.Text.Length > 1)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41093";
                ex.Data["errDesc"] = "TypeStepText given invalid Selection.Text '" + selection.Text + "'";
                throw ex;
                return;
            }
            switch (stepTextType)
            {
                case StepTextType.qtMainStep:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    selection.Font.Bold = 1;
                    selection.TypeText(stepName + ":");
                    selection.Font.Bold = 0;
                    selection.TypeText(" ");
                    Field_UpdateAll();
                    break;
                case StepTextType.qtMainDouble:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    break;
                case StepTextType.qtMainSingle:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    break;
                case StepTextType.qtSubStep:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    selection.Range.ListFormat.ListIndent();
                    break;
                case StepTextType.qtSubStepNoKeep:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    selection.Range.ListFormat.ListIndent();
                    break;
                case StepTextType.qtSubDouble:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    break;
                case StepTextType.qtSubSingle:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    break;
                case StepTextType.qtSubSubStep:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    selection.Range.ListFormat.ListIndent();
                    break;
                case StepTextType.qtSubSubDouble:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    break;
                case StepTextType.qtSubSubSingle:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    break;
                case StepTextType.qtImmediate:
                    break;
                case StepTextType.qtContinue:
                    TypeContinue(strStepText);
                    break;
                case StepTextType.qtNormalSingle:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    break;
                case StepTextType.qtNormalDouble:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    break;
                case StepTextType.qtNormalDoubleNoKeep:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    break;
                case StepTextType.qtNormalTriple:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    break;
                case StepTextType.qtMBRHeading:
                    selection.InsertBreak(WdBreakType.wdPageBreak);
                    Table table = document.Tables.Add(selection.Range, 1, 1);
                    table.Columns[1].SetWidth(document.Application.InchesToPoints(6.7f), WdRulerStyle.wdAdjustNone);
                    table.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    table.Borders[WdBorderType.wdBorderTop].LineWidth = WdLineWidth.wdLineWidth225pt;
                    table.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    table.Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth225pt;
                    table.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
                    table.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;

                    Cell cell = table.Cell(1, 1);
                    cell.Range.Text = "    ";
                    cell.Range.Select();
                    selection.Collapse(WdCollapseDirection.wdCollapseStart);

                    if (!selection.Information[WdInformation.wdWithInTable])
                    {
                        var e = new Exception();
                        e.Data["errNum"] = "41477";
                        e.Data["errDesc"] = "MBRHeading is NOT in TABLE";
                        throw e;
                    }
                    selection.set_Style(Get_Style(document, gstrStyle_MBRSteps));
                    selection.Tables[1].Cell(1, 1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    break;
                case StepTextType.qtMBRStep:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    selection.Range.ListFormat.ListIndent();
                    break;
                case StepTextType.qtMBRSub:
                    StepTextSpacing(stepTextType, locator: myProcName);
                    selection.Range.ListFormat.ListIndent();
                    selection.Range.ListFormat.ListIndent();
                    break;
                default:
                    var ex = new Exception();
                    ex.Data["errNum"] = "41089";
                    ex.Data["errDesc"] = "'" + stepTextType + "' invalid parameter for " + myProcName;
                    throw ex;
                    break;


            }

            TypeStepString(strStepText, boldAll);
            //Todo:
            //timeElapsed = DateTime.Now - timeStart;
            //dbltotalStepText = dblTotalStepText + dblElapsed;
            // intCountStepText = intCountStepText + 1; // not req
        }


        private static bool blnInitialized = false;
        private static List<TblStepText> arrStepText = new List<TblStepText>();
        private static int intStepText = 0;

        public static string Get_Step_Text(string pstrStepName, bool blnFound = true)
        {
            string strStepNameUCase = pstrStepName.ToUpper();

            if (!blnInitialized)
            {
                int intStepScan = DatabaseHelper.StepText_SelectCOUNT();
                if (intStepScan == 0) { return ""; }
                arrStepText = new List<TblStepText>(intStepScan);
                intStepText = 0;
                blnInitialized = true;
            }

            for (int intStepScan = intStepText; intStepScan >= 1; intStepScan--)
            {
                if (intStepScan - 1 < arrStepText.Count && arrStepText[intStepScan - 1].StepName == strStepNameUCase)
                {
                    return arrStepText[intStepScan - 1].StepText;
                }
            }

            List<TblStepText> arsStepText = DatabaseHelper.StepText_Select(strStepNameUCase);
            if (arsStepText.Count > 0)
            {
                foreach (var stepText in arsStepText)
                {
                    intStepText++;
                    var stepData = new TblStepText
                    {
                        StepName = stepText.StepName.ToUpper(),
                        StepText = stepText.StepText,
                        IsFound = "T"
                    };
                    arsStepText.Add(stepData);

                    if (stepData.StepName == strStepNameUCase)
                    {
                        return stepData.StepText;
                    }
                }
                blnFound = true;
            }
            else
            {
                intStepText++;
                TblStepText stepData = new TblStepText
                {
                    StepName = strStepNameUCase,
                    StepText = "STEPTEXT <" + strStepNameUCase + "> NOT FOUND",
                    IsFound = "F"
                };
                arsStepText.Add(stepData);
                blnFound = false;
            }
            return arrStepText[intStepText - 1].StepText;
        }

        private static void TypeContinue(string strStepText)
        {
            string prevChars;
            string prevChars1;
            string prevChars2;
            string strNextChar1;
            string strNextChar2;
            int intBlanksProvided;
            int intBlanksNeeded;
            int intDifference;

            Word.Range rangPrevChars = wordApplication.Selection.Range;
            rangPrevChars.Start = rangPrevChars.End - 5;
            prevChars = rangPrevChars.Text;
            prevChars1 = prevChars.Substring(prevChars.Length - 1, 1);
            prevChars2 = prevChars.Substring(prevChars.Length - 2, 1);

            if (strStepText.Length > 0)
            {
                strNextChar1 = prevChars.Substring(strStepText.Length - 1, 1);

                if (strStepText.Length > 1)
                {
                    strNextChar2 = strStepText.Substring(1, 1);
                }
                else
                {
                    strNextChar2 = "";
                }
            }
            else
            {
                strNextChar1 = "";
                strNextChar2 = "";
            }
            if (string.IsNullOrWhiteSpace(strNextChar1))
            {
                if (string.IsNullOrWhiteSpace(strNextChar2))
                {
                    intBlanksProvided = 2;

                }
                else
                {
                    intBlanksProvided = 1;

                }
            }
            else
            {
                intBlanksProvided = 0;

            }
            if (prevChars1 == "." || prevChars1 == ":")
            {
                intBlanksNeeded = 2;
            }
            else if (string.IsNullOrWhiteSpace(prevChars1))
            {
                if (prevChars2 == "." || prevChars2 == ":")
                {
                    intBlanksNeeded = 1;

                }
                else if (string.IsNullOrWhiteSpace(prevChars2))
                {
                    intBlanksNeeded = 0;
                }
                else
                {
                    intBlanksNeeded = 0;
                }
            }
            else
            {
                if (strNextChar1 == ".")
                {
                    intBlanksNeeded = 0;
                }
                else if (strNextChar1 == ";")
                {
                    intBlanksNeeded = 0;
                }
                else if (strNextChar1 == ":")
                {
                    intBlanksNeeded = 0;
                }
                else
                {
                    intBlanksNeeded = 1;
                }
            }

            if (intBlanksNeeded > intBlanksProvided)
            {
                intDifference = intBlanksNeeded - intBlanksProvided;
                string spaces = new string(' ', intDifference);

                wordApplication.Selection.TypeText(spaces);
            }
            else if (intBlanksNeeded < intBlanksProvided)
            {
                intDifference = intBlanksProvided - intBlanksNeeded;
                strStepText = strStepText.Substring(strStepText.Length - intDifference);

            }
        }

        private static void TypeStepString(string strIn, bool pblnBoldAll)
        {
            const string MyProcName = "TypeStepString";

            int intIn = strIn.Length;
            int intA;
            string strChar;
            string strCommon = "";

            for (intA = 1; intA <= intIn; intA++)
            {
                strChar = strIn.Substring(intA - 1, 1);

                if (strChar == gstrStepText_MarkupStart)
                {
                    WordCommonText(strCommon, pblnBoldAll);
                    StepVarSubstitution(strIn, ref intA, pblnBoldAll: pblnBoldAll);
                }
                else if (strChar == "\r" || strChar == "\n")
                {
                    if (strChar == "\n")
                    {
                        wordApplication.Selection.TypeParagraph();
                    }
                    continue;
                }
                else
                {
                    strCommon += strChar;
                }
            }
            WordCommonText(strCommon, pblnBoldAll);
        }

        //private static void TypeStepString(string strIn, bool pblnBoldAll)
        //{
        //    const string MyProcName = "TypeStepString";

        //    int intIn = strIn.Length;
        //    int intA;
        //    string strChar;
        //    string strCommon = "";

        //    for (intA = 1; intA <= intIn; intA++)
        //    {
        //        strChar = strIn.Substring(intA - 1, 1);

        //        if (strChar == gstrStepText_MarkupStart)
        //        {
        //            WordCommonText(strCommon, pblnBoldAll);
        //            StepVarSubstitution(strIn, ref intA, pblnBoldAll: pblnBoldAll);
        //        }
        //        else
        //        {
        //            strCommon += strChar;
        //        }
        //    }

        //    WordCommonText(strCommon, pblnBoldAll);
        //}

        private static string StepVarSubstitution(string strIn, ref int pintScan, bool pblnBoldAll, bool pblnReturnString = false)
        {
            const string MyProcName = "StepVarSubstitution";
            int intScanOrig = pintScan;
            string strChar;
            string strVarName = "";
            string strVarValueExt = "";
            int intIn = strIn.Length + 1;
            int intSymbol = 0;
            bool blnBoldable = false;
            bool blnVarBold = false;
            bool blnVarCenter = false;
            bool blnVarUnderline = false;
            bool blnVarItalic = false;
            string strVarSuperSub = "";
            bool blnVarLiteral = false;
            bool blnVarProperty = false;
            int intFontSize = -1;
            string strColorCode = "";
            long lngColorCode = 0;
            int intColorCodeRED = 0;
            int intColorCodeGREEN = 0;
            int intColorCodeBLUE = 0;
            int intColorCodeLength = 0;
            bool blnMoreFormatting = true;
            string strQuotechar = " ";
            int intBracketdepth = 0;

            strChar = strIn.Substring(pintScan - 1, 1); // Adjusting for zero-based index
            if (strChar != gstrStepText_MarkupStart)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41082";
                ex.Data["errDesc"] = $"ERROR: Expecting '[variable]', detected '{strIn.Substring(pintScan - 1, Math.Min(10, strIn.Length - pintScan + 1))}...'";
                throw ex;
            }

            pintScan++;
            for (pintScan = pintScan; pintScan < intIn; pintScan++)
            {
                strChar = strIn.Substring(pintScan - 1, 1); // Adjusting for zero-based index
                if (strChar == "'" || strChar == "\"")
                {
                    if (strChar == strQuotechar)
                    {
                        strQuotechar = " ";
                    }
                    else if (string.IsNullOrWhiteSpace(strQuotechar))
                    {
                        strQuotechar = strChar;
                    }
                }

                if (strChar == gstrStepText_MarkupStart)
                {
                    if (!string.IsNullOrWhiteSpace(strQuotechar))
                    {
                        strVarName += strChar;
                    }
                    else
                    {
                        strVarName += strChar;
                        intBracketdepth++;
                    }
                }
                else if (strChar == gstrStepText_MarkupEnd)
                {
                    if (!string.IsNullOrWhiteSpace(strQuotechar))
                    {
                        strVarName += strChar;
                    }
                    else if (intBracketdepth > 1)
                    {
                        strVarName += strChar;
                        intBracketdepth--;
                    }
                    else
                    {
                        break;
                    }
                }
                else
                {
                    strVarName += strChar;
                }
            }

            if (strChar != gstrStepText_MarkupEnd)
            {
                var ex = new Exception();
                ex.Data["errNum"] = "41083";
                ex.Data["errDesc"] = $"ERROR: did not detect end-of-variable ']' in '{strIn.Substring(intScanOrig - 1)}'";
                throw ex;
            }

            // Reset formatting variables
            blnVarBold = blnVarCenter = blnVarUnderline = blnVarItalic = false;
            strVarSuperSub = "";
            blnVarLiteral = blnVarProperty = false;
            intFontSize = -1;
            strColorCode = "";
            lngColorCode = 0;
            blnMoreFormatting = true;

            while (blnMoreFormatting)
            {
                strChar = strVarName.Substring(0, 1);
                switch (strChar)
                {
                    case gstrStepText_Subscript:
                        if (!string.IsNullOrEmpty(strVarSuperSub))
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MULTIPLE SUBSCRIPT ({gstrStepText_Superscript}{gstrStepText_Subscript}) FORMATTING";
                            throw ex;
                        }
                        strVarSuperSub = gstrStepText_Subscript;
                        break;

                    case gstrStepText_Superscript:
                        if (!string.IsNullOrEmpty(strVarSuperSub))
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MULTIPLE SUPERSCRIPT ({gstrStepText_Superscript}{gstrStepText_Subscript}) FORMATTING";
                            throw ex;
                        }
                        strVarSuperSub = gstrStepText_Superscript;
                        break;

                    case gstrStepText_Bold:
                        if (blnVarBold)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MULTIPLE BOLD ({gstrStepText_Bold}) FORMATTING";
                            throw ex;
                        }
                        blnVarBold = true;
                        break;

                    case gstrStepText_Underline:
                        if (blnVarUnderline)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MULTIPLE UNDERLINE ({gstrStepText_Underline}) FORMATTING";
                            throw ex;
                        }
                        blnVarUnderline = true;
                        break;

                    case gstrStepText_Centered:
                        if (blnVarCenter)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MULTIPLE CENTERING ({gstrStepText_Centered}) FORMATTING";
                            throw ex;
                        }
                        blnVarCenter = true;
                        break;

                    case gstrStepText_Italic:
                        if (blnVarItalic)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MULTIPLE ITALIC ({gstrStepText_Italic}) FORMATTING";
                            throw ex;
                        }
                        blnVarItalic = true;
                        break;

                    case gstrStepText_Color:
                        if (!string.IsNullOrEmpty(strColorCode))
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MULTIPLE COLOR ({gstrStepText_Color}) FORMATTING";
                            throw ex;
                        }
                        if (strVarName.Length < 4)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MISSING COLORCODE AFTER ({gstrStepText_Color})";
                            throw ex;
                        }

                        intColorCodeLength = strVarName.Length < 2 ? 2 : 4;
                        strColorCode = strVarName.Substring(0, intColorCodeLength);

                        if (intColorCodeLength == 4)
                        {
                            if (!int.TryParse(strColorCode.Substring(1, 3), out _))
                            {
                                intColorCodeLength = 2;
                                strColorCode = strColorCode.Substring(0, intColorCodeLength);
                            }
                        }

                        if (intColorCodeLength == 2)
                        {
                            switch (strColorCode.Substring(1, 1).ToUpper())
                            {
                                case gstrStepText_Color_Red:
                                    //strColorCode = gstrStepText_Color + "800";
                                    lngColorCode = (long)WdColor.wdColorRed;
                                    break;
                                case gstrStepText_Color_Green:
                                    //strColorCode = gstrStepText_Color + "080";
                                    lngColorCode = (long)WdColor.wdColorGreen;
                                    break;
                                case gstrStepText_Color_Blue:
                                    //strColorCode = gstrStepText_Color + "008";
                                    lngColorCode = (long)WdColor.wdColorBlue;
                                    break;
                                case gstrStepText_Color_Yellow:
                                    //strColorCode = gstrStepText_Color + "880";
                                    lngColorCode = (long)WdColor.wdColorYellow;
                                    break;
                                case gstrStepText_Color_Magenta:
                                    //strColorCode = gstrStepText_Color + "808";
                                    lngColorCode = (long)WdColor.wdColorDarkRed;
                                    break;
                                case gstrStepText_Color_Cyan:
                                    //strColorCode = gstrStepText_Color + "088";
                                    lngColorCode = (long)WdColor.wdColorTeal;
                                    break;
                                default:
                                    var ex = new Exception();
                                    ex.Data["errNum"] = "41983";
                                    ex.Data["errDesc"] = $"INVALID COLORCODE AFTER ({gstrStepText_Color})";
                                    throw ex;
                            }
                        }


                        strVarName = strVarName.Substring(intColorCodeLength - 1);
                        break;

                    case gstrStepText_Literal:
                        if (blnVarLiteral)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MULTIPLE LITERAL ({gstrStepText_Literal}) FORMATTING";
                            throw ex;
                        }
                        if (blnVarProperty)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"CONFLICTING PROPERTY({gstrStepText_Property})/LITERAL({gstrStepText_Literal}) FORMATTING";
                            throw ex;
                        }
                        blnVarLiteral = true;
                        break;

                    case gstrStepText_Property:
                        if (blnVarProperty)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MULTIPLE PROPERTY ({gstrStepText_Property}) FORMATTING";
                            throw ex;
                        }
                        if (blnVarLiteral)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"CONFLICTING PROPERTY({gstrStepText_Property})/LITERAL({gstrStepText_Literal}) FORMATTING";
                            throw ex;
                        }
                        blnVarProperty = true;
                        break;

                    case gstrStepText_Fontsize:
                        if (intFontSize != -1)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MULTIPLE FONTSIZE ({gstrStepText_Fontsize}) SPECIFICATIONS";
                            throw ex;
                        }
                        if (strVarName.Length < 3)
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"MISSING FONTSIZE AFTER ({gstrStepText_Fontsize})";
                            throw ex;
                        }
                        if (!int.TryParse(strVarName.Substring(1, 2), out intFontSize))
                        {
                            var ex = new Exception();
                            ex.Data["errNum"] = "41983";
                            ex.Data["errDesc"] = $"NONNUMERIC FONTSIZE '{strVarName.Substring(1, 2)}' AFTER ({gstrStepText_Fontsize})";
                            throw ex;
                        }
                        strVarName = strVarName.Substring(2);
                        break;

                    default:
                        blnMoreFormatting = false;
                        break;
                }

                if (blnMoreFormatting)
                {
                    strVarName = strVarName.Substring(1);
                }
            }

            if (blnVarProperty)
            {
                strVarValueExt = strVarName;
                blnBoldable = false;
            }
            else if (blnVarLiteral)
            {
                strVarValueExt = strVarName;
                blnBoldable = false;
            }
            else if ((strVarName.StartsWith("\"") && strVarName.EndsWith("\"")) || (strVarName.StartsWith("'") && strVarName.EndsWith("'")))
            {
                strVarValueExt = strVarName.Substring(1, strVarName.Length - 2);
                blnBoldable = false;
            }
            else if (strVarName.StartsWith("$"))
            {
                StepVarSymbol(strVarName, strVarValueExt, intSymbol);
            }
            else
            {
                StepVarRetrieveExt(strVarName, strVarValueExt, blnBoldable, intSymbol);
            }

            if (blnVarBold)
            {
                blnBoldable = true;
            }
            if (pblnBoldAll)
            {
                blnBoldable = true;
            }

            if (pblnReturnString)
            {
                return strVarValueExt;
            }
            else
            {
                WordVariableText(strVarValueExt, blnBoldable, blnVarCenter, intSymbol, strVarSuperSub, blnVarUnderline, blnVarItalic, lngColorCode, blnVarProperty, intFontSize);
                return null;
            }
        }

        private static void WordCommonText(string text, bool boldAll)
        {
            const string myProcName = "WordCommonText";

            if (!string.IsNullOrEmpty(text.Trim()))
            {
                if (wordApplication.Selection.Font.Bold == 1 && !boldAll)
                {
                    wordApplication.Selection.Font.Bold = 0;
                }

                if (wordApplication.Selection.Font.Bold == 0 && boldAll)
                {
                    wordApplication.Selection.Font.Bold = 1;
                }
            }
            wordApplication.Selection.TypeText(text);
        }

        private static void Field_UpdateAll()
        {
            foreach (Field field in wordApplication.ActiveDocument.Fields)
            {
                field.Update();
            }
        }

        public static void WordVariableText(string pstrText, bool pblnBoldNeeded, bool pblnAlignCenter, int pintSymbol, string pstrSuperSub, bool pblnVarUnderline, bool pblnVarItalic, long plngColor, bool pblnFieldProperty, int pintFontSize)
        {
            const string MyProcName = "WordVariableText";
            long lngColorSave = 0;

            if (pblnBoldNeeded)
            {
                if (wordApplication.Selection.Font.Bold == 0)
                {
                    wordApplication.Selection.Font.Bold = 1;
                }
            }
            else
            {
                if (wordApplication.Selection.Font.Bold == 1)
                {
                    wordApplication.Selection.Font.Bold = 0;
                }
            }

            if (pblnVarUnderline)
            {
                if (wordApplication.Selection.Font.Underline == WdUnderline.wdUnderlineNone)
                {
                    wordApplication.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;
                }
            }
            else
            {
                if (wordApplication.Selection.Font.Underline != WdUnderline.wdUnderlineNone)
                {
                    wordApplication.Selection.Font.Underline = WdUnderline.wdUnderlineNone;
                }
            }

            if (pblnVarItalic)
            {
                if (wordApplication.Selection.Font.Italic == 0)
                {
                    wordApplication.Selection.Font.Italic = 1;
                }
            }
            else
            {
                if (wordApplication.Selection.Font.Italic == 1)
                {
                    wordApplication.Selection.Font.Italic = 0;
                }
            }

            switch (pstrSuperSub)
            {
                case "":
                    break;
                case gstrStepText_Subscript:
                    wordApplication.Selection.Font.Subscript = 1;
                    break;
                case gstrStepText_Superscript:
                    wordApplication.Selection.Font.Superscript = 1;
                    break;
                default:
                    var ex = new Exception();
                    ex.Data["errNum"] = "41070";
                    ex.Data["errDesc"] = $"Unexpected SuperSub='{pstrSuperSub}' in {MyProcName}";
                    throw ex;
            }

            if (pintFontSize > 0)
            {
                wordApplication.Selection.Font.Size = pintFontSize;
            }

            if (pblnAlignCenter)
            {
                wordApplication.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
            else
            {
                if (wordApplication.Selection.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphLeft)
                {
                    wordApplication.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                }
            }

            if (plngColor != (long)WdColor.wdColorAutomatic)
            {
                lngColorSave = (long)wordApplication.Selection.Font.Color;
                wordApplication.Selection.Font.Color = (WdColor)plngColor;
            }

            if (pblnFieldProperty)
            {
                Field_InsertProperty(wordApplication.Selection.Range, pstrText);
            }
            else
            {
                wordApplication.Selection.TypeText(pstrText);
            }

            if (pintSymbol != 0)
            {
                wordApplication.Selection.InsertSymbol(CharacterNumber: pintSymbol, Font: "Symbol", Unicode: true);
            }

            if (plngColor != (long)WdColor.wdColorAutomatic)
            {
                wordApplication.Selection.Font.Color = (WdColor)lngColorSave;
            }

            switch (pstrSuperSub)
            {
                case "":
                    break;
                case gstrStepText_Subscript:
                    wordApplication.Selection.Font.Subscript = 0;
                    break;
                case gstrStepText_Superscript:
                    wordApplication.Selection.Font.Superscript = 0;
                    break;
                default:
                    var ex = new Exception();
                    ex.Data["errNum"] = "41071";
                    ex.Data["errDesc"] = $"Unexpected SuperSub='{pstrSuperSub}' in {MyProcName}";
                    throw ex;
            }
        }

        //private static void Wizard_Terminate(Word.Document pdocDocument, bool pblnCalledWizard, string pstrProtectionPassword)
        //{
        //    try
        //    {

        //        if (!pblnCalledWizard)
        //        {
        //            _formManager.Wizard_BookMarkRemoval(pdocDocument);
        //            CommonService.DocumentProtection_Assure(pdocDocument, pstrProtectionPassword);
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //}
        public static void StepVarClear()
        {
            StepVarManager.ManageVariable("CLR", mvarStepVariable);
        }

        public static void PageSetup(Range pageRange)
        {
            PageSetup pageSetup = pageRange.PageSetup;

            pageSetup.TopMargin = ConvertInchesToPoints(1);
            pageSetup.BottomMargin = ConvertInchesToPoints(1);
            pageSetup.LeftMargin = ConvertInchesToPoints(0.7f);
            pageSetup.RightMargin = ConvertInchesToPoints(0.5f);
            pageSetup.LineNumbering.Active = 0; // false
            pageSetup.Orientation = WdOrientation.wdOrientPortrait;
            pageSetup.Gutter = ConvertInchesToPoints(0);
            pageSetup.PageWidth = ConvertInchesToPoints(8.5f);
            pageSetup.PageHeight = ConvertInchesToPoints(11);
            pageSetup.FirstPageTray = WdPaperTray.wdPrinterDefaultBin;
            pageSetup.OtherPagesTray = WdPaperTray.wdPrinterDefaultBin;
            pageSetup.SectionStart = WdSectionStart.wdSectionNewPage;
            pageSetup.OddAndEvenPagesHeaderFooter = 0; // false
            pageSetup.DifferentFirstPageHeaderFooter = 0; // false
            pageSetup.VerticalAlignment = WdVerticalAlignment.wdAlignVerticalTop; //WdAlignVertical.wdAlignVerticalTop;
            pageSetup.SuppressEndnotes = 0; // false
            pageSetup.MirrorMargins = 0; // false
        }



        public static void InsertHeaderMBR(int sectionIndex, string lotCaption, Document activeDocument)
        {
            const string MyProcName = "InsertHeaderMBR";
            Range rng = activeDocument.Sections[sectionIndex].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

            rng.Select();
            rng.Font.Size = 4;

            Table tbl = rng.Tables.Add(rng, 4, 2);
            tbl.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

            tbl.Range.Font.Bold = 0;
            tbl.Range.Font.Size = 12;
            tbl.Columns[1].SetWidth(ConvertInchesToPoints(3.35f), WdRulerStyle.wdAdjustNone);
            tbl.Columns[2].SetWidth(ConvertInchesToPoints(4f), WdRulerStyle.wdAdjustNone);
            tbl.Rows.SpaceBetweenColumns = ConvertInchesToPoints(0.15f);
            tbl.Cell(1, 1).Merge(tbl.Cell(1, 2));

            InsertHeaderCellBuilder(tbl.Cell(1, 1), "", cRecordType, false, 8, WdParagraphAlignment.wdAlignParagraphLeft);
            InsertHeaderCellBuilder(tbl.Cell(2, 1), "Product ID: ", cProductID, true, 16, WdParagraphAlignment.wdAlignParagraphLeft);
            InsertHeaderCellBuilder(tbl.Cell(2, 2), lotCaption, cLotNum, true, 16, WdParagraphAlignment.wdAlignParagraphRight);
            tbl.Cell(3, 1).Merge(tbl.Cell(3, 2));
            InsertHeaderCellBuilder(tbl.Cell(3, 1), "Product Name: ", cProductName, false, 12, WdParagraphAlignment.wdAlignParagraphLeft);
            tbl.Cell(4, 1).Range.Font.Size = 12;
            tbl.Cell(4, 1).Range.ParagraphFormat.SpaceBefore = 0.0f; //  
            tbl.Cell(4, 1).Range.ParagraphFormat.SpaceAfter = 0.0f;  //  
            InsertHeaderCellBuilder(tbl.Cell(4, 2), "Order Number: ", cOrderNum, false, 12, WdParagraphAlignment.wdAlignParagraphRight);
            SetBorders(tbl, "Header"); //  
        }

        private static void SetBorders(Table tbl, string type) //  
        {
            // Set all borders to none
            tbl.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
            tbl.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
            tbl.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
            tbl.Borders[WdBorderType.wdBorderHorizontal].LineStyle = WdLineStyle.wdLineStyleNone;
            tbl.Borders[WdBorderType.wdBorderVertical].LineStyle = WdLineStyle.wdLineStyleNone;

            //var bottomBorder = tbl.Borders[WdBorderType.wdBorderBottom]; bottomBorder.LineStyle = WdLineStyle.wdLineStyleSingle; bottomBorder.LineWidth = WdLineWidth.wdLineWidth300pt;
            //// LineWidth in points (300pt = 3pt)
            //bottomBorder.Color = WdColor.wdColorAutomatic;

            //  changes begin
            if (type == "Header")
            {
                var bottomBorder = tbl.Borders[WdBorderType.wdBorderBottom]; 
                bottomBorder.LineStyle = WdLineStyle.wdLineStyleSingle; 
                bottomBorder.LineWidth = WdLineWidth.wdLineWidth300pt;
                bottomBorder.Color = WdColor.wdColorAutomatic;
            }

            if (type == "Footer")
            {
                var topBorder = tbl.Borders[WdBorderType.wdBorderTop]; topBorder.LineStyle = WdLineStyle.wdLineStyleSingle; topBorder.LineWidth = WdLineWidth.wdLineWidth300pt;
                topBorder.Color = WdColor.wdColorAutomatic;
            }
        }

        private static void InsertHeaderCellBuilder(Cell cel, string cellCaption, string cellPropertyName, bool cellPropertyBold, int cellPropertyFontSize, WdParagraphAlignment cellParagraphAlignment)
        {
            Range rng = cel.Range;
            rng.Text = "";
            rng.Font.Name = "Times New Roman"; // 2301
            rng.Font.Bold = cellPropertyBold ? 1 : 0;
            rng.Font.Size = cellPropertyFontSize;

            Field_InsertProperty(rng, cellPropertyName);

            rng = cel.Range;
            rng.Collapse(WdCollapseDirection.wdCollapseStart);
            rng.Text = cellCaption;
            rng.Font.Size = cellPropertyFontSize;
            rng.Font.Bold = 0;
            rng.Font.Name = "Times New Roman"; // 2301
            rng.ParagraphFormat.Alignment = cellParagraphAlignment;

            rng.ParagraphFormat.SpaceBefore = 0.0f; //  
            rng.ParagraphFormat.SpaceAfter = 0.0f;  //  
        }

        public static void InsertFooterMBR(int sectionIndex, bool introFooter, bool restartNumbering, Document activeDocument)
        {
            const string MyProcName = "InsertFooterMBR";
            Table tbl;
            Range rng;

            IntroPageNumbering(sectionIndex, restartNumbering);
            rng = activeDocument.Sections[sectionIndex].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

            // Delete existing tables in the footer
            while (rng.Tables.Count > 0)
            {
                rng.Tables[1].Delete();
            }

            rng.Select();
            rng.Font.Size = 1;

            tbl = rng.Tables.Add(rng, 3, 4);
            rng.Font.Name = "Times New Roman"; //  
            tbl.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

            // Set font and column widths
            tbl.Range.Font.Bold = 1;
            tbl.Range.Font.Size = 11;
            tbl.Columns[1].SetWidth(ConvertInchesToPoints(1.4f), WdRulerStyle.wdAdjustNone);
            tbl.Columns[2].SetWidth(ConvertInchesToPoints(2.35f), WdRulerStyle.wdAdjustNone);
            tbl.Columns[3].SetWidth(ConvertInchesToPoints(2.25f), WdRulerStyle.wdAdjustNone);
            tbl.Columns[4].SetWidth(ConvertInchesToPoints(1.35f), WdRulerStyle.wdAdjustNone);
            tbl.Rows.SpaceBetweenColumns = ConvertInchesToPoints(0.15f);

            tbl.Cell(1, 2).Merge(tbl.Cell(1, 4));
            tbl.Cell(1, 1).Range.Text = "Filename:";
            Field_InsertProperty(tbl.Cell(1, 2).Range, cDocumentTitle);
            tbl.Cell(1, 2).Range.InsertBefore(", ");
            Field_InsertProperty(tbl.Cell(1, 2).Range, cFileName);

            rng.ParagraphFormat.SpaceBefore = 0.0f;  //  
            rng.ParagraphFormat.SpaceAfter = 0.0f;  //  

            tbl.Cell(2, 1).Range.Text = "Change Number:";
            Field_InsertProperty(tbl.Cell(2, 2).Range, cChangeNum);

            rng = tbl.Cell(2, 3).Range;
            rng.Text = "";
            Field_InsertProperty(rng, cPrintTime);

            rng = tbl.Cell(2, 3).Range;
            rng.Collapse(WdCollapseDirection.wdCollapseStart);
            rng.Text = " ";
            rng.Collapse(WdCollapseDirection.wdCollapseStart);
            Field_InsertProperty(rng, cPrintDate);

            rng = tbl.Cell(2, 4).Range;
            rng.Text = "";
            if (!introFooter)
            {
                Field_InsertProperty(rng, cPrintPages);
                rng.Text = " of ";
            }
            Field_InsertProperty(rng, cPage);

            rng = tbl.Cell(2, 4).Range;
            rng.Select();
            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            rng.Collapse(WdCollapseDirection.wdCollapseStart);
            rng.Text = introFooter ? "Intro Page: " : "Page: ";

            rng.ParagraphFormat.SpaceBefore = 0.0f;  //  
            rng.ParagraphFormat.SpaceAfter = 0.0f;   //  

            rng = tbl.Rows[3].Range;
            rng.Cells.Merge();
            rng.Text = " ";
            rng.Collapse(WdCollapseDirection.wdCollapseStart);
            rng.Select();
            rng.Collapse(WdCollapseDirection.wdCollapseStart);

            TypeStepText(StepTextType.qtImmediate, "PNU_Confidential");
            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            // Set borders
            SetBorders(tbl, "Footer"); // 
        }

        public static void IntroPageNumbering(int sectionIndex, bool restartNumbering)
        {
            var footer = Globals.ThisDocument.ActiveDocument.Sections[sectionIndex].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary]; //220125
            footer.LinkToPrevious = !restartNumbering;

            var pageNumbers = footer.PageNumbers;
            pageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic;
            pageNumbers.HeadingLevelForChapter = 0;
            pageNumbers.IncludeChapterNumber = false;
            pageNumbers.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen;
            pageNumbers.RestartNumberingAtSection = restartNumbering;

            pageNumbers.StartingNumber = restartNumbering ? 1 : 0;
        }

        public static void Field_InsertProperty(Range where, string propertyName) 
        {
            bool checkProperty = false;
            bool PropertExists = false;
            string propertyNameUpper = propertyName.ToUpper();
            Document docCurrent = where.Document;
            Field myField;

            where.Collapse(WdCollapseDirection.wdCollapseStart);

            switch (propertyNameUpper)
            {
                case "FILENAME":
                    myField = docCurrent.Fields.Add(where, WdFieldType.wdFieldFileName, true);
                    checkProperty = false;
                    break;
                case "PAGE":
                    myField = docCurrent.Fields.Add(where, WdFieldType.wdFieldPage, true);
                    checkProperty = false;
                    break;
                default:
                    myField = docCurrent.Fields.Add(where, WdFieldType.wdFieldDocProperty, propertyName, true);
                    checkProperty = true;
                    break;
            }

            Field_ApplyShading(myField);

            if (checkProperty)
            {
                foreach (var Property in docCurrent.CustomDocumentProperties)
                {

                    if (Property.Name.ToUpper() == propertyNameUpper)
                    {
                        PropertExists = true;
                        break;
                    }
                }

                if (!PropertExists)
                {
                    //MessageBox.Show("WARNING: Field references Property  " + propertyName + "which DOES NOT EXIST in Document" + docCurrent.Name + "Property will be added to Document.", "Field added", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DocumentProperty_PropertyAdd(DocumentPropertyName: propertyName, Bracketed: true, document: docCurrent);

                }

            }
        }



        public static void Field_ApplyShading(Field fld)
        {
            fld.Result.Font.Shading.Texture = WdTextureIndex.wdTexture20Percent;
            fld.Result.Font.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic;
            fld.Result.Font.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
        }
        public static void RecordingBlock(string Caption, double TextRowHeight = 0.25, double DataRowHeight = 0.5)
        {
            const string MyProcName = "RecordingBlock";
            Word.Table tblRecording;
            int intRowCount;
            const int cintXxByColumn = 2;
            const int cintDoneByTextRow = 1;
            const int cintDoneByDataRow = 2;
            const int cintChkdByTextRow = 3;
            const int cintChkdByDataRow = 4;
            // Word.Application wordApplication = null;

            Word.Selection selection = wordApplication.Selection;

            Table_Exit();
            StepTextSpacing(StepTextType.qtAutoText);
            intRowCount = cintDoneByDataRow;
            tblRecording = wordApplication.ActiveDocument.Tables.Add(Range: selection.Range, NumRows: intRowCount, NumColumns: 2, DefaultTableBehavior: WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior: WdAutoFitBehavior.wdAutoFitFixed);
            tblRecording.Rows.AllowBreakAcrossPages = 0;
            tblRecording.Range.Paragraphs.KeepWithNext = -1;
            tblRecording.Rows[intRowCount].Range.Paragraphs.KeepWithNext = -1;
            tblRecording.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
            tblRecording.Columns[1].PreferredWidth = ConvertInchesToPoints(5f);
            tblRecording.Columns[2].PreferredWidth = ConvertInchesToPoints(1.5f);
            tblRecording.Rows.HeightRule = WdRowHeightRule.wdRowHeightAtLeast;
            tblRecording.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom;
            tblRecording.Rows[cintDoneByTextRow].Height = ConvertInchesToPoints((float)TextRowHeight);
            tblRecording.Rows[cintDoneByDataRow].Height = ConvertInchesToPoints((float)DataRowHeight);

            tblRecording.Borders[WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblRecording.Borders[WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblRecording.Borders[WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblRecording.Borders[WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblRecording.Borders[WdBorderType.wdBorderHorizontal].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblRecording.Borders[WdBorderType.wdBorderVertical].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblRecording.Borders.Shadow = false;


            Table_CellBorder(Table: tblRecording, Row: cintDoneByDataRow, Column: cintXxByColumn);
            tblRecording.Cell(Row: cintDoneByTextRow, Column: cintXxByColumn).Range.Text = Caption;
            Table_Exit();
        }

        public static void SignAndDate(bool CheckedBy, bool SecondCheck = false, string DoneByCaption = "Done By", string CheckedByCaption = "Checked By", string SecondCheckCaption = "Second Check", string DateText = "Date", string AppliesTo = "", double TextRowHeight = 0.25, double DataRowHeight = 0.5)
        {
            const string MyProcName = "SignAndDate";
            Word.Table tblDBCB;
            int intRowCount;
            const int cintAppliesToColumn = 1;
            const int cintXxByColumn = 2;
            const int cintDateColumn = 3;
            const int cintDoneByTextRow = 1;
            const int cintDoneByDataRow = 2;
            const int cintChkdByTextRow = 3;
            const int cintChkdByDataRow = 4;
            const int cintSecondTextRow = 5;
            const int cintSecondDataRow = 6;
            // Word.Application wordApplication = null;

            Word.Selection selection = wordApplication.Selection;
            if (SecondCheck & !CheckedBy)
            {
                //Information.Err.Raise(Number: 43501, Source: MyProcName, Description: "CANNOT HAVE SECOND-CHECK WITHOUT CHECKED-BY");
                return;
            }

            Table_Exit();
            StepTextSpacing(StepTextType.qtAutoText);
            if (CheckedBy)
            {
                if (SecondCheck)
                {
                    intRowCount = cintSecondDataRow;
                }
                else
                {
                    intRowCount = cintChkdByDataRow;
                }
            }
            else
            {
                intRowCount = cintDoneByDataRow;
            }

            tblDBCB = wordApplication.ActiveDocument.Tables.Add(Range: selection.Range, NumRows: intRowCount, NumColumns: 3, DefaultTableBehavior: WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior: WdAutoFitBehavior.wdAutoFitFixed);
            tblDBCB.Rows.AllowBreakAcrossPages = 0;
            tblDBCB.Range.Paragraphs.KeepWithNext = -1;
            tblDBCB.Rows[intRowCount].Range.Paragraphs.KeepWithNext = -1;
            tblDBCB.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
            tblDBCB.Columns[1].PreferredWidth = ConvertInchesToPoints(3f);
            tblDBCB.Columns[2].PreferredWidth = ConvertInchesToPoints(2f);
            tblDBCB.Columns[3].PreferredWidth = ConvertInchesToPoints(1f);
            tblDBCB.Rows.HeightRule = WdRowHeightRule.wdRowHeightAtLeast;
            tblDBCB.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom;
            tblDBCB.Rows[cintDoneByTextRow].Height = ConvertInchesToPoints((float)TextRowHeight);
            tblDBCB.Rows[cintDoneByDataRow].Height = ConvertInchesToPoints((float)DataRowHeight);

            if (CheckedBy)
            {
                tblDBCB.Rows[cintChkdByTextRow].Height = ConvertInchesToPoints((float)TextRowHeight);
                tblDBCB.Rows[cintChkdByDataRow].Height = ConvertInchesToPoints((float)DataRowHeight);
            }

            if (SecondCheck)
            {

                tblDBCB.Rows[cintSecondTextRow].Height = ConvertInchesToPoints((float)TextRowHeight);
                tblDBCB.Rows[cintSecondDataRow].Height = ConvertInchesToPoints((float)DataRowHeight);
            }


            tblDBCB.Borders[WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblDBCB.Borders[WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblDBCB.Borders[WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblDBCB.Borders[WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblDBCB.Borders[WdBorderType.wdBorderHorizontal].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblDBCB.Borders[WdBorderType.wdBorderVertical].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tblDBCB.Borders.Shadow = false;

            Table_CellBorder(Table: tblDBCB, Row: cintDoneByDataRow, Column: cintXxByColumn);
            Table_CellBorder(Table: tblDBCB, Row: cintDoneByDataRow, Column: cintDateColumn);
            if (CheckedBy)
            {
                Table_CellBorder(Table: tblDBCB, Row: cintChkdByDataRow, Column: cintXxByColumn);
                Table_CellBorder(Table: tblDBCB, Row: cintChkdByDataRow, Column: cintDateColumn);
            }

            if (SecondCheck)
            {
                Table_CellBorder(Table: tblDBCB, Row: cintSecondDataRow, Column: cintXxByColumn);
                Table_CellBorder(Table: tblDBCB, Row: cintSecondDataRow, Column: cintDateColumn);
            }


            tblDBCB.Cell(Row: cintDoneByDataRow, Column: cintAppliesToColumn).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            tblDBCB.Cell(Row: cintDoneByDataRow, Column: cintAppliesToColumn).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tblDBCB.Cell(Row: cintDoneByDataRow, Column: cintAppliesToColumn).Range.Font.Bold = 1;
            tblDBCB.Cell(Row: cintDoneByDataRow, Column: cintAppliesToColumn).Range.Text = AppliesTo;
            tblDBCB.Cell(Row: cintDoneByTextRow, Column: cintXxByColumn).Range.Text = DoneByCaption;
            tblDBCB.Cell(Row: cintDoneByTextRow, Column: cintDateColumn).Range.Text = DateText;
            if (CheckedBy)
            {
                tblDBCB.Cell(Row: cintChkdByTextRow, Column: cintXxByColumn).Range.Text = CheckedByCaption;
                tblDBCB.Cell(Row: cintChkdByTextRow, Column: cintDateColumn).Range.Text = DateText;
            }

            if (SecondCheck)
            {
                tblDBCB.Cell(Row: cintSecondTextRow, Column: cintXxByColumn).Range.Text = SecondCheckCaption;
                tblDBCB.Cell(Row: cintSecondTextRow, Column: cintDateColumn).Range.Text = DateText;
            }
        }

        public static void Table_CellBorder(Table Table, int Row, int Column)
        {
            {
                Table.Cell(Row, Column).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                Table.Cell(Row, Column).Range.Borders[WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                Table.Cell(Row, Column).Range.Borders[WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                Table.Cell(Row, Column).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;


            }
        }

        public static void Field_PropertyDefine(Document activedocument, string fld)
        {
            const string MyProcName = "Field_PropertyDefine";
            Word.Field fldScan;
            DocumentProperty Prop;
            Word.CustomProperties DocProps;
            string strFieldCode;
            int intSlashPos;
            int intStartPos;
            int intSpacePos;
            bool blnExists = false;
            intStartPos = fld.IndexOf(" DOCPROPERTY ", 1) + 13;
            if (intStartPos == 0)
            {
                //Todo
                //Information.Err.Raise(3456, MyProcName, "MISSING DOCPROPERTY IN '" + fld + "'");
                return;
            }

            intSlashPos = fld.IndexOf(@" \", 1);
            if (intSlashPos > 0)
            {
                strFieldCode = fld.Substring(intStartPos, intSlashPos - intStartPos);
            }
            else
            {
                strFieldCode = fld.Substring(intStartPos);
            }
            foreach (var property in wordApplication.ActiveDocument.CustomDocumentProperties)
            {
                if (property.Name == strFieldCode)
                {
                    blnExists = true;
                }
            }

            if (blnExists)
            {
                return;
            }
            else
            {
                DocumentProperty_PropertyAdd(DocumentPropertyName: strFieldCode, Bracketed: true, document: activedocument);
            }
        }

        public static void DocumentProperty_PropertyAdd(string DocumentPropertyName, bool Bracketed = false, string DocumentPropertyValue = "", Document document = null)
        {

            const string MyProcName = "DocumentProperty_Add";
            string PropertyValue = string.Empty;
            //DocumentPropertyName = "Test1";

            if (Bracketed)
            {

                if (DocumentPropertyValue != "")
                {
                    //Todo
                    //Information.Err.Raise(Number: 41561, Source: MyProcName, Description: "CANNOT SPECIFY BOTH BRACKETED AND A VALUE");


                }
                PropertyValue = "[" + DocumentPropertyName + "]";
            }
            else
            {
                PropertyValue = DocumentPropertyValue;

            }

            object properties = document.CustomDocumentProperties;
            Type propertiesType = properties.GetType();
            object[] DocCustomProperty = { DocumentPropertyName, false, MsoDocProperties.msoPropertyTypeString, PropertyValue };
            propertiesType.InvokeMember("Add", BindingFlags.InvokeMethod, null, properties, DocCustomProperty);
            //document.Save();
            return;

        }
        public static void DocumentProperty_Add(DocumentProperties DocProperties, string DocumentPropertyName, bool Bracketed = false, string DocumentPropertyValue = "")
        {

            const string MyProcName = "DocumentProperty_Add";
            string PropertyValue = string.Empty;
            //DocumentPropertyName = "Test";

            if (Bracketed)
            {

                if (DocumentPropertyValue != "")
                {
                    //Todo
                    //Information.Err.Raise(Number: 41561, Source: MyProcName, Description: "CANNOT SPECIFY BOTH BRACKETED AND A VALUE");


                }
                PropertyValue = "[" + DocumentPropertyName + "]";
            }
            else
            {
                PropertyValue = DocumentPropertyValue;

            }

            /*doc.CustomDocumentProperties.Add(DocumentPropertyName, false,
                                                      MsoDocProperties.msoPropertyTypeString,
                                                      PropertyValue);*/

            DocProperties.Add(DocumentPropertyName, LinkToContent: false, MsoDocProperties.msoPropertyTypeString, Value: PropertyValue);//160125

            /*object properties = document.CustomDocumentProperties;
            Type propertiesType = DocProperties.GetType();
            object[] DocCustomProperty = { DocumentPropertyName, false, MsoDocProperties.msoPropertyTypeString, PropertyValue };
            propertiesType.InvokeMember("Add", BindingFlags.InvokeMethod, null, (object)DocProperties, DocCustomProperty);*/

            return;

        }
        
        public static void DocumentProperty_Update(DocumentProperties pdpcDocProperties, string pstrDocumentPropertyName, string pstrDocumentPropertyValue)
        {
            pdpcDocProperties[pstrDocumentPropertyName].Value = pstrDocumentPropertyValue;
            
        }

        public static void DocumentProperty_Delete(DocumentProperties pdpcDocProperties, string pstrDocumentPropertyName)
        {

            pdpcDocProperties[pstrDocumentPropertyName].Delete();
            
            return;

        }

        private static float ConvertInchesToPoints(float inches)
        {
            return inches * 72; // 1 inch = 72 points
        }

        public static void ViewPage_Main(Word.Application wordApp)
        {
            while (wordApp.ActiveWindow.Panes.Count > 1)
            {
                wordApp.ActiveWindow.Panes[2].Close();
            }

            var activePane = wordApp.ActiveWindow.ActivePane.View;
            activePane.Type = WdViewType.wdPrintView;
            activePane.SeekView = WdSeekView.wdSeekMainDocument;
        }

        public static void Bookmark_IntroEnd_Add(Range rangeHere)
        {
            if (rangeHere != null)
            {
                rangeHere.Bookmarks.Add(gstrIntroEnd, rangeHere);
            }
        }

        public static void SectionArray_Build(Document activeDocument, ref List<bool> pablnSection) //220125 // 2301
        {
            pablnSection = new List<bool>(new bool[activeDocument.Sections.Count]);

            for (int intSectionIndex = 0; intSectionIndex < pablnSection.Count; intSectionIndex++)
            {
                pablnSection[intSectionIndex] = false;
            }
           
        }

        public static void DocumentProtection_Create(List<bool> pablnSectionProtected, string pstrProtectionPassword, int pintBookmarksDeleted, Document activeDocument)
        {
            const string MyProcName = "DocumentProtection_Create";
            Section wscSection;
            Bookmark bmk;
            Range rng;
            int intBookmarksDeleted = 0;
            int intSectionCount;
            int intSectionIndex;

            List<TblLookupData> lookupDataList = DatabaseHelper.LookupData_Select("BookMark", "1");

            if (lookupDataList == null)
            {  // NewCode2
                var ex = new Exception();
                ex.Data["errNum"] = "0000";
                ex.Data["errDesc"] = "Issue on BookMark LookupData.";
                throw ex;
            }

            if (activeDocument.ProtectionType != WdProtectionType.wdNoProtection)
            {
                //throw new Exception();
                var ex = new Exception();
                ex.Data["errNum"] = "41060";
                ex.Data["errDesc"] = "'" + activeDocument.Name + "' ALREADY PROTECTED";
                throw ex;
            }

            foreach (Section section in activeDocument.Sections)
            {
                section.ProtectedForForms = false;
            }

            intBookmarksDeleted = 0;

            foreach (Bookmark bookmark in activeDocument.Bookmarks)
            {
                rng = bookmark.Range;

                if (pablnSectionProtected[(rng.Sections[1].Index)-1]) // 2301
                {
                    if (!BookMarkIsReserved(lookupDataList, bookmark.Name))
                    {
                        bookmark.Delete();
                        intBookmarksDeleted++;
                    }
                }
            }

            pintBookmarksDeleted = intBookmarksDeleted;

            intSectionCount = activeDocument.Sections.Count < pablnSectionProtected.Count ?
                activeDocument.Sections.Count : pablnSectionProtected.Count;

            for (intSectionIndex = 1; intSectionIndex <= intSectionCount; intSectionIndex++)
            {
                activeDocument.Sections[intSectionIndex].ProtectedForForms = pablnSectionProtected[intSectionIndex - 1];
            }

            

        }

        public static bool BookMarkIsReserved(List<TblLookupData> prstBookmarks, string bookmarkName)
        {
            string strBookmarkName = bookmarkName.ToUpper();
            int intBookmarkLen = strBookmarkName.Length;
            string strBookmarkPrefix;
            string strBookmarkSuffix;

            if (intBookmarkLen > 2)
            {
                strBookmarkPrefix = strBookmarkName.Substring(0, intBookmarkLen - 2);
                strBookmarkSuffix = strBookmarkName.Substring(intBookmarkLen - 2, 2);
            }
            else
            {
                strBookmarkPrefix = "";
                strBookmarkSuffix = "";
            }

            bool blnBNUReserved = false;

            if (prstBookmarks.Count > 0)
            {
                foreach (var bookmark in prstBookmarks)
                {
                    string strBookmarkLookup = bookmark.LookupText.ToUpper();
                    int intBookmarkLookupLen = strBookmarkLookup.Length;
                    string strBookmarkLookupPrefix;
                    string strBookmarkLookupSuffix;

                    if (intBookmarkLookupLen > 2)
                    {
                        strBookmarkLookupPrefix = strBookmarkLookup.Substring(0, intBookmarkLookupLen - 2);
                        strBookmarkLookupSuffix = strBookmarkLookup.Substring(intBookmarkLookupLen - 2, 2);
                    }
                    else
                    {
                        strBookmarkLookupPrefix = "XX";
                        strBookmarkLookupSuffix = "XX";
                    }

                    if (strBookmarkLookupSuffix == "++")
                    {
                        if (intBookmarkLookupLen == intBookmarkLen)
                        {
                            if (strBookmarkLookupPrefix == strBookmarkPrefix && CommonService.IsNumeric(strBookmarkSuffix))
                            {
                                blnBNUReserved = true;
                                break;
                            }
                        }
                    }
                    else if (strBookmarkLookup == strBookmarkName)
                    {
                        blnBNUReserved = true;
                        break;
                    }
                }
            }
            return blnBNUReserved;
        }


        private static bool IsNumeric(string value)
        {
            return double.TryParse(value, out _);
        }

        public static bool QReturn_QGet(string compName, qReturn qReturn)
        {

            const string MyProcName = "QReturn_QGet";
            bool blnRunQget;
            switch (qReturn)
            {
                case qReturn.qReturn_Cancel:
                case qReturn.qReturn_Written:
                    blnRunQget = false;
                    break;
                case qReturn.qReturn_Back:
                case qReturn.qReturn_Next:
                case qReturn.qReturn_Collapse:
                case qReturn.qReturn_Valid:
                case qReturn.qReturn_Finish:
                    blnRunQget = true;
                    break;

                default:
                    MessageBox.Show("INVALID QReturn=" + qReturn + " from " + compName, MyProcName, MessageBoxButtons.OK);
                    blnRunQget = false;
                    break;

            }

            return blnRunQget;


        }

        public static void InsertFooterMBR_PageNumber(Table tbl, bool blnIntroFooter) 
        {
            Range rng = tbl.Cell(2, 4).Range;
            rng.Text = "";
            if (!blnIntroFooter)
            {
                Field_InsertProperty(rng, cPrintPages);
                rng.Text = " of ";
            }
            Field_InsertProperty(rng, cPage);
            rng = tbl.Cell(2, 4).Range;
            rng.Select();
            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            rng.Collapse(WdCollapseDirection.wdCollapseStart);
            if (blnIntroFooter)
            {
                rng.Text = "Intro Page: ";

            }
            else
            {
                rng.Text = "Page: ";
            }
        }

        public static void ListBox_To_Array(DataGridView dataGrid, string[,] strArray)
        {
            const string MyProcName = "ListBox_To_Array";

            int rowCount = dataGrid.Rows.Count;
            int columnCount = dataGrid.Columns.Count;

            if (columnCount < 1)
            {
                MessageBox.Show("Cannot process ListBox '" + dataGrid.Name + "' with " + columnCount + " Columns", myServiceName + "." + MyProcName);
                strArray = null; // Return null if processing cannot proceed
                return;
            }

            // array to hold the values from the DataGrid
            strArray = new string[rowCount, columnCount];

            for (int row = 0; row < rowCount; row++)
            {
                for (int col = 0; col < columnCount; col++)
                {
                    strArray[row, col] = dataGrid.Rows[row].Cells[col].Value?.ToString();
                }
            }
        }

        public static Word.Table Table_Find_ModuleMBR(Word.Document pdocCurrent)
        {
            return Table_Find(pdocCurrent, "Module List");

        }

        public static Word.Table Table_Find_PropertyMBR(Word.Document pdocCurrent)
        {
            return Table_Find(pdocCurrent, "Property List");

        }

    }

    class SwapPair
    {
        public string SPFrom { get; set; }
        public string SPTo { get; set; }
    }

    public class StepVariable
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public bool IsReplaceable { get; set; }
        public bool IsScalable { get; set; }
        public bool IsBold { get; set; }
        public int Symbol { get; set; }
    }

    public static class StepVarManager
    {
        private const int MaxVariables = 200;
        private const string ProcedureName = "StepVarManager";

        private static readonly List<StepVariable> _stepVariables = new List<StepVariable>();
        private static int _variableCount;
        private static int _variableLevel;
        private static bool _hasRun;

        public static void ManageVariable(string action, StepVariable variable)
        {
            if (!_hasRun)
            {
                _hasRun = true;
                _variableLevel = 0;
            }

            string variableName = variable?.Name?.ToUpper() ?? string.Empty;

            switch (action.ToUpper())
            {
                case "C":
                case "CLR":
                    ClearVariables();
                    break;

                case "G":
                case "GET":
                    GetVariable(ref variable, variableName);
                    break;

                case "P":
                case "PUT":
                    PutVariable(variable, variableName);
                    break;

                case "R":
                case "REP":
                    ReplaceVariable(variable, variableName);
                    break;

                case "D":
                case "DEL":
                    DeleteVariable(variableName);
                    break;

                case "POP":
                    PopVariable();
                    break;

                case "PSH":
                    PushVariable();
                    break;

                default:
                    throw new ApplicationException($"Action '{action}' is invalid in {ProcedureName}");
            }
        }

        private static void ClearVariables()
        {
            for (int i = 0; i < _variableCount; i++)
            {
                _stepVariables[i] = new StepVariable
                {
                    IsBold = false,
                    Name = string.Empty,
                    IsReplaceable = false,
                    IsScalable = false,
                    Symbol = 0,
                    Value = string.Empty
                };
            }
            _variableCount = 0;
        }

        private static void GetVariable(ref StepVariable variable, string variableName)
        {
            for (int i = 0; i < _variableCount; i++)
            {
                if (_stepVariables[i].Name == variableName)
                {
                    variable = _stepVariables[i];
                    return;
                }
            }

            variable.Value = $"*[{variable.Name}]NOTFOUND*";
            variable.Name = string.Empty;
            variable.IsBold = false;
            variable.IsReplaceable = false;
            variable.IsScalable = false;
        }

        private static void PutVariable(StepVariable variable, string variableName)
        {
            for (int i = 0; i < _variableCount; i++)
            {
                if (_stepVariables[i].Name == variableName)
                {
                    throw new ApplicationException($"'{variableName}' already exists for 'PUT' in {ProcedureName}");
                }
            }

            if (_variableCount >= MaxVariables)
            {
                throw new ApplicationException($"Too many StepVariables in {ProcedureName}");
            }

            _stepVariables.Add(variable);
            _variableCount++;
            _stepVariables.Sort((x, y) => x.Name.CompareTo(y.Name));
        }

        private static void ReplaceVariable(StepVariable variable, string variableName)
        {
            for (int i = 0; i < _variableCount; i++)
            {
                if (_stepVariables[i].Name == variableName)
                {
                    _stepVariables[i] = variable;
                    return;
                }
            }

            throw new ApplicationException($"'{variableName}' not found for 'REP' in {ProcedureName}");
        }

        private static void DeleteVariable(string variableName)
        {
            bool found = false;

            for (int i = 0; i < _variableCount; i++)
            {
                if (_stepVariables[i].Name != variableName)
                {
                    _stepVariables[found ? i - 1 : i] = _stepVariables[i];
                }
                else
                {
                    found = true;
                }
            }

            if (found)
            {
                _variableCount--;
                _stepVariables.RemoveAt(_variableCount);
            }
            else
            {
                throw new ApplicationException($"'{variableName}' not found for 'DEL' in {ProcedureName}");
            }
        }

        private static void PopVariable()
        {
            if (_variableLevel > 0)
            {
                _variableLevel--;
            }
            else
            {
                throw new ApplicationException($"Cannot POP from VarLevel {_variableLevel} in {ProcedureName}");
            }
        }

        private static void PushVariable()
        {
            if (_variableLevel < MaxVariables)
            {
                _variableLevel++;
            }
            else
            {
                throw new ApplicationException($"Cannot PUSH from VarLevel {_variableLevel} in {ProcedureName}");
            }
        }

        public static void Form_Push(Form form)
        {
            Form_PushPop(form, false);
        }

        public static void Form_Pop(Form form)
        {
            Form_PushPop(form, true);
        }

        private static void Form_PushPop(Form frm, bool IsVisible)
        {
            const string MyProcName = "Form_PushPop";
            System.Windows.Forms.Control ctl;
            string strControlName;
            System.Windows.Forms.Label lbl;

            if (IsVisible)
            {
                StepVarPop();
            }
            else
            {
                StepVarPush();
            }

            frm.Controls["mpgWizardPage"].Visible = IsVisible;
            frm.Controls["pnlCmdButtons"].Visible = IsVisible;

            foreach (Control ctrl in frm.Controls["panel1"].Controls)
            {
                strControlName = ctrl.Name;

                if (strControlName.Contains("shp"))
                {
                    if (!(ctrl is Label))
                    {
                        MessageBox.Show("ERROR: " + strControlName + " IS NOT A 'LABEL'", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    lbl = (Label)ctrl;
                    if (IsVisible)
                    {
                        if (lbl.BackColor == Color.Blue)//COLOR_CALLING)
                        {
                            lbl.BackColor = Color.Lime;//Color.Green;//COLOR_ACTIVE;
                            break;
                        }

                    }
                    else
                    {
                        if (lbl.BackColor == Color.Lime) //Color.Green)//COLOR_ACTIVE
                        {
                            lbl.BackColor = Color.Blue;// COLOR_CALLING; 
                            break;
                        }
                    }

                }
            }
        }

        public static void StepVarPop()
        {
            StepVariable qsvX = null;
            ManageVariable("POP", qsvX);
        }
        public static void StepVarPush()
        {
            StepVariable qsvX = null;
            ManageVariable("PSH", qsvX);
        }
    }

    public class CustomException : Exception
    {
        public int ErrNum { get; }
        public string ErrDesc { get; }

        public CustomException(int errNum, string errDesc)
        {
            ErrNum = errNum;
            ErrDesc = errDesc;
        }


    }
}
