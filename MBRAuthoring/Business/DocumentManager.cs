using MBRAuthoring.DataAccess;
using MBRAuthoring.Models;
using MBRAuthoring.Presentation;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static MBRAuthoring.ThisDocument;
using Application = Microsoft.Office.Interop.Word.Application;

namespace MBRAuthoring.Business
{
    public interface IDocumentManager
    {
        Document OpenDocument(string filePath, string readPassword, string writePassword, String fileName = "");
        DocumentInfo Document_SelectQPSObject(string docName, DocumentInfo documentInfo, bool blnRefresh = false);
        DocumentSelectAttribute Document_SelectAttributes(string DocumentName, string DocumentWorkCenter = "", int DocumentWorkCenterID = 0); 
        bool DocEqualsLastSave(string docName); //This is in modMain()-->But we need to access wordapp and document inside this method. So written here.
        void Document_Has(string DocumentName, ref string DocIsModule, ref string DocHasModule, ref string DocHasMaterial, ref string DocHasProduct, ref string DocHasProperty, ref string DocHasWorkCenter, ref string DocHasWorkCenterAlt, ref string DocPermitsQuantitySufficient);
        bool ValidDocument(string docName);
        string PendingUpdateAny(string docName);
        string Document_SelectPasswordProtect(string pstrDocName);

    }
    public class DocumentManager : IDocumentManager
    {
        Document doc = null;
        Application wordApp = new Application();
        const string MyProcName = "Document_SelectQPSObject";
        List<DocumentInfo> docsInfo = new List<DocumentInfo>();
        ClsDocument clsDocument = null;

        public Document OpenDocument(string filePath, string readPassword, string writePassword, String fileName = "")
        {
            try
            {
                wordApp.Visible = true;
                doc = wordApp.Documents.Open(FileName: filePath, ConfirmConversions: false, ReadOnly: false, AddToRecentFiles: false, PasswordDocument: readPassword, WritePasswordDocument: writePassword, PasswordTemplate: "", Revert: false, Format: WdOpenFormat.wdOpenFormatAuto);
                wordApp.DocumentBeforeClose += Document_BeforeClose;
                wordApp.DocumentBeforeSave += Document_BeforeSave;
                doc.Application.Activate(); // NewCode
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                // Release COM objects to avoid memory leaks
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
            return doc;
        }

        public void Document_BeforeSave(Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            if (SaveAsUI)
            {
                SaveManager.Document_Save_FileAndDatabase(doc);

                if (false) // FileSaveAs in VBA - but won't trigger just in case we need it in future
                {
                    string strFileName = string.Empty;
                    try
                    {
                        if (!CommonService.gblnUserType_REVIEWER)
                        {
                            MessageBox.Show("This MBR document cannot be saved under a new file name.", "MBR - SaveAs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                        strFileName = @"M:\" + doc.Name;
                        while (true)
                        {
                            strFileName = InputBox.ShowDialog("Enter Target path and File for SaveAs:", "MBR SaveAs for REVIEWER", docName: strFileName);
                            if (string.IsNullOrEmpty(strFileName))
                            {

                                MessageBox.Show("Document was NOT Saved", "Save As", MessageBoxButtons.OK);
                                break;
                            }
                            try
                            {
                                doc.SaveAs2(FileName: strFileName, Password: "", WritePassword: "");
                                break;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("ERROR - 'SaveAs' Failed:\n\n" + "Error Number: " + ex.HResult + "\n" + "Error Desription: " + ex.Message, "MBR 'SaveAs' Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
        }
        public void Document_BeforeClose(Document doc, ref bool Cancel)
        {
            Template tptAttached;
            int intScaledValues;
            string strDBTransactionName;
            string strMessageAddendum;
            try
            {
                strDBTransactionName = DatabaseHelper.Transaction_Active();
                if (strDBTransactionName == doc.Name)
                {
                    strMessageAddendum = "(DATABASE UPDATES are pending, and will be committed.)";
                }
                else if (string.IsNullOrEmpty(strDBTransactionName))
                {
                    strMessageAddendum = "(NO Database Updates are pending.)";
                }
                else
                {
                    strMessageAddendum = "(NO Database Updates are pending for this Document.)";
                }
                if (!doc.Saved)
                {
                    IntPtr hwnd = (IntPtr)doc.Windows[1].Hwnd;
                    IWin32Window win32Window = new WindowWrapper(hwnd);

                    DialogResult result = MessageBox.Show(win32Window, $"Save Changes made to {doc.Name} ?\n\n{strMessageAddendum}", "MBR Document Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question);  // NewCode
                                                                                                                                                                                                               //, MessageBoxDefaultButton.Button1, handle);
                                                                                                                                                                                                               //Globals.ThisDocument.Set_Parent_AND_ShowDialog(formEditMaterials, ActiveDocument);
                    if (result == DialogResult.Yes)
                    {
                        SaveManager.Document_Save_FileAndDatabase(doc);
                    }
                    else
                    {
                        if (strDBTransactionName == doc.Name)
                        {
                            DatabaseHelper.Transaction_Rollback(strDBTransactionName);
                        }
                    }
                }
                tptAttached = doc.get_AttachedTemplate();
                tptAttached.Saved = true;
                doc.Saved = true;
                ReleaseDocumentInfo(doc.Name);
                if (Globals.ThisDocument.docDetails.ContainsKey(doc.Name)) // 2301
                {
                    Globals.ThisDocument.docDetails.Remove(doc.Name);
                }
            }
            finally
            {
                if (doc != null)
                {
                    Marshal.ReleaseComObject(doc);
                    doc = null;
                }
                if (wordApp != null)
                {
                    Marshal.ReleaseComObject(wordApp);
                    wordApp = null;
                }
            }
        }

        private void ReleaseDocumentInfo(string docName)
        {
            try
            {
                docsInfo.RemoveAll(d => d.GetClsDocument.DocumentName == docName);
            }
            catch (Exception ex)
            {
                throw ex;  // NewCode1
            }
        }

        public DocumentInfo Document_SelectQPSObject(string docName, DocumentInfo documentInfo = null, bool blnRefresh = false)
        {

            try
            {
                List<ClsDocument> clsdocsInfo = new List<ClsDocument>();
                clsDocument = clsdocsInfo.Find(d => d.DocumentName == docName);
                if (clsDocument is null)
                {
                    clsDocument = new ClsDocument();
                    documentInfo = new DocumentInfo();

                    if (!DatabaseHelper.Read(docName, documentInfo))
                    {
                        clsDocument = null;
                        documentInfo = null;
                        var ex = new Exception();
                        ex.Data["errNum"] = "42042";
                        ex.Data["errDesc"] = "modDocInfo.Document_SelectQPSObject - Unable to read document information from database for '" + docName + "'";
                        throw ex;
                    }
                    else
                    {
                        //todo m_colDocs.Add qdcDocument, pstrDocName
                        docsInfo.Add(documentInfo);
                    }

                }
                else if (blnRefresh)
                {
                    if (!DatabaseHelper.Read(docName, documentInfo))
                    {
                        clsDocument = null;
                        documentInfo = null;
                        var ex = new Exception();
                        ex.Data["errNum"] = "42042";
                        ex.Data["errDesc"] = "modDocInfo.Document_SelectQPSObject - Unable to read document information from database for '" + docName + "'";
                        throw ex;
                    }

                }

            }

            catch (Exception ex)
            {
                throw ex;  // NewCode1
            }
            return documentInfo;

            //finally
            //{
            //    // Release COM objects to avoid memory leaks
            //    // System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            //}

        }

        public DocumentSelectAttribute Document_SelectAttributes(string DocumentName, string DocumentWorkCenter = "", int DocumentWorkCenterID = 0) //(DocumentSelectAttribute documentSelectAttribute)
        {
            const string MyProcName = "Document_SelectAttributes";
            DocumentInfo documentInfo = new DocumentInfo();
            DocumentSelectAttribute selectAttribute = new DocumentSelectAttribute();
            try
            {  // NewCode1
                documentInfo = Document_SelectQPSObject(DocumentName, documentInfo);
                selectAttribute.DocumentName = documentInfo.GetClsDocument.DocumentName;
                selectAttribute.DocumentID = documentInfo.GetClsDocument.DocID;
                selectAttribute.DocumentOpUnit = documentInfo.GetTblOpUnit.OpUnit;
                selectAttribute.DocumentOpUnitID = documentInfo.GetClsDocument.OpUnitID;
                selectAttribute.DocumentOpUnitType = documentInfo.GetTblOpUnit.OpUnitType;
                selectAttribute.DocumentDocType = documentInfo.GetDocumentType.DocType;
                selectAttribute.DocumentDocTypeAbb = documentInfo.GetDocumentType.DocTypeAbb;
                selectAttribute.DocumentDocTypeID = documentInfo.GetClsDocument.DocTypeID;
                selectAttribute.DocumentProtectionPassword = documentInfo.GetClsDocument.DocPasswordProtect;
                selectAttribute.DocumentTrackChangesReq = documentInfo.GetClsDocument.mblnTrackChangesReq;
                //DocumentWorkCenter = clsDocument.WorkCenterID;

            }
            catch (Exception ex)
            {
                throw ex;  // NewCode1
            }
            return selectAttribute;
        }


        public bool DocEqualsLastSave(string docName)
        {

            string strChanges;
            string Differences;
            bool blnRestoreRevView;
            bool docEqualsLastSave;

            blnRestoreRevView = wordApp.ActiveWindow.View.ShowRevisionsAndComments;
            wordApp.ActiveWindow.View.ShowRevisionsAndComments = false;
            wordApp.ActiveWindow.View.RevisionsView = WdRevisionsView.wdRevisionsViewFinal;
            docEqualsLastSave = true;
            blnRestoreRevView = wordApp.ActiveWindow.View.ShowRevisionsAndComments;
            var result = Document_SelectAttributes(docName);

            const string MyProcName = "DocEqualsLastSave";
            DocImageLog imageLogs = DatabaseHelper.DocImageLog_SelectLast(result.DocumentID);
            if (imageLogs == null)
            {
                return docEqualsLastSave = true;
            }
            Differences = "";
            if (imageLogs.DocFieldsCount > 0)
            {
                if (doc.Fields.Count != imageLogs.DocFieldsCount)
                {
                    Differences = Differences + "\r\n" + "Fields: " + "\t" + "were " + imageLogs.DocFieldsCount + ", but are now " + doc.Fields.Count;

                }

            }
            if (imageLogs.DocParagraphsCount > 0)
            {
                if (doc.Paragraphs.Count != imageLogs.DocParagraphsCount)
                {
                    Differences = Differences + "\r\n" + "Paragraphs: " + "\t" + "were " + imageLogs.DocParagraphsCount + ", but are now " + doc.Paragraphs.Count;

                }

            }
            if (imageLogs.DocCustomPropertiesCount > 0)
            {
                if (doc.CustomDocumentProperties.Count != imageLogs.DocCustomPropertiesCount)
                {
                    Differences = Differences + "\r\n" + "Doc Properties: " + "\t" + "were " + imageLogs.DocCustomPropertiesCount + ", but are now " + doc.CustomDocumentProperties.Count;

                }
            }
            
            if (imageLogs.DocEndPageNumber > 0)
            {
                if (doc.Range().Information[Microsoft.Office.Interop.Word.WdInformation.wdNumberOfPagesInDocument] != imageLogs.DocEndPageNumber)
                {
                    Differences = Differences + "\r\n" + "Pages: " + '\t' + "were " + imageLogs.DocEndPageNumber + ", but are now " + doc.Range().Information[Microsoft.Office.Interop.Word.WdInformation.wdNumberOfPagesInDocument];  // NewCode1
                }

            }
            if (imageLogs.DocSectionsCount > 0)
            {
                if (doc.Sections.Count != imageLogs.DocSectionsCount)
                {
                    Differences = Differences + "\r\n" + "Sections: " + "\t" + "were " + imageLogs.DocSectionsCount + ", but are now " + doc.Sections.Count;
                }
            }

            if (imageLogs.DocSentencesCount > 0)
            {
                if (doc.Sentences.Count != imageLogs.DocSentencesCount)
                {
                    Differences = Differences + "\r\n" + "Sentences: " + "\t" + "were " + imageLogs.DocSentencesCount + ", but are now " + doc.Sentences.Count;

                }
            }

            if (imageLogs.DocTablesCount > 0)
            {
                if (doc.Tables.Count != imageLogs.DocTablesCount)
                {
                    Differences = Differences + "\r\n" + "Tables: " + "\t" + "were " + imageLogs.DocTablesCount + ", but are now " + doc.Tables.Count;  // NewCode1
                }
            }

            if (imageLogs.DocWordsCount > 0)
            {
                if (doc.Words.Count != imageLogs.DocWordsCount)
                {
                    Differences = Differences + "\r\n" + "Words: " + "\t" + "were " + imageLogs.DocWordsCount + ", but are now " + doc.Words.Count;  // NewCode1

                }
            }
            if (imageLogs.DocCharactersCount > 0)
            {
                if (doc.Characters.Count != imageLogs.DocCharactersCount)
                {
                    Differences = Differences + "\r\n" + "Characters: " + "\t" + "were " + imageLogs.DocCharactersCount + ", but are now " + doc.Characters.Count;

                }
            }
            //Differences = "";
            if (!string.IsNullOrEmpty(Differences))
            {
                strChanges = docName.ToUpper() + "\r\n" + "\r\n" + "has changed since the last MBR_Authoring Save on " + "\r\n" + "\r\n" + imageLogs.LogDate + "\r\n" + "\r\n" + "Count Differences: " + "\r\n" + Differences + "\r\n" + "\r\n" + "Do you want to continue with this Document?";

                IntPtr hwnd = (IntPtr)Globals.ThisDocument.ActiveDocument.Windows[1].Hwnd;
                IWin32Window win32Window = new WindowWrapper(hwnd);

                DialogResult dialogResult = MessageBox.Show(win32Window, strChanges, "ERROR: Document Invalid", MessageBoxButtons.YesNo, MessageBoxIcon.Error);  // NewCode

                if (dialogResult == DialogResult.Yes)
                {
                    return docEqualsLastSave = true;
                }
                else
                {
                    return docEqualsLastSave = false;
                }
            }
            doc.ActiveWindow.View.ShowRevisionsAndComments = blnRestoreRevView;
            return docEqualsLastSave;
        }
        public void Document_Has(string DocumentName, ref string DocIsModule, ref string DocHasModule, ref string DocHasMaterial, ref string DocHasProduct, ref string DocHasProperty, ref string DocHasWorkCenter, ref string DocHasWorkCenterAlt, ref string DocPermitsQuantitySufficient)
        {
            const string MyProcName = "Document_Has";
            DocumentInfo documentInfo = new DocumentInfo();
            try
            {
                documentInfo = Document_SelectQPSObject(DocumentName, documentInfo);  // NewCode1
                DocIsModule = documentInfo.GetOpUnitDocType.DocIsModule;
                DocHasModule = documentInfo.GetOpUnitDocType.DocHasModule;
                DocHasMaterial = documentInfo.GetOpUnitDocType.DocHasMaterial;
                DocHasProduct = documentInfo.GetOpUnitDocType.DocHasProduct;
                DocHasProperty = documentInfo.GetOpUnitDocType.DocHasProperty;
                DocPermitsQuantitySufficient = documentInfo.GetOpUnitDocType.DocPermitsQuantitySufficient;

            }
            catch (Exception ex)
            {
                throw ex;  // NewCode1
            }


        }

        public bool ValidDocument(string docName)
        {
            const string MyProcName = "ValidDocument";
            try
            {
                DocumentInfo documentInfo = new DocumentInfo();
                documentInfo = Document_SelectQPSObject(docName);
                if (documentInfo is null)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;  // NewCode1
                return false;
            }

        }
        public string PendingUpdateAny(string docName)//need to check
        {
            const string MyProcName = "PendingUpdateAny";
            string pendingUpdateAny = "";
            DocumentInfo documentInfo = new DocumentInfo();
            try  // NewCode1
            {
            documentInfo = Document_SelectQPSObject(docName);
            if (documentInfo.GetClsDocument.mblnPendingUpdate)
            {
                pendingUpdateAny = documentInfo.GetClsDocument.DocumentName;
                }
            }
            catch(Exception ex)
            {
                throw ex;  // NewCode1
            }
            return pendingUpdateAny;
        }

        public string Document_SelectPasswordProtect(string pstrDocName)  // NewCode3
        {
            DocumentInfo qdcDocument = Document_SelectQPSObject(pstrDocName, null, true);
            return qdcDocument.GetClsDocument.DocPasswordProtect;
        }


    }
}


