using MBRAuthoring.DataAccess;
using MBRAuthoring.Models;
using MBRAuthoring.Presentation;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace MBRAuthoring.Business
{
    static class SaveManager
    {
        // Constants for retry logic- RPC error PANNEJ02
        private const uint RpcCallRejectedErrorCode = 0x80010001;
        private const int WordComRetryCount = 6;
        private const int WordComRetryDelayMs = 150;

        static IFormManager formManager = FormBase.formMangerInstance;
        static IDocumentManager _documentManager = null;// ServiceLocator.GetDocumentManager();
                                                        //public ICommandBarManager _commandBarManager = null;
        private static TblOpUnit _tblOpUnit;
        private static OpUnitDocType _opUnitDocType;
        const string MyCompName = "SaveManager";

        static SaveManager()
        {
            _documentManager = new DocumentManager();
        }
        public static void SetOpUnit(TblOpUnit tblOpUnit, OpUnitDocType opUnitDocType)
        {
            _tblOpUnit = tblOpUnit;
            _opUnitDocType = opUnitDocType;
        }

        public static void Document_Save_FileAndDatabase(Document pdocCurrent, bool pblnLockDocument = false, string pstrSaveAsFileName = "", string pstrDocPasswordRead = "", string pstrDocPasswordWrite = "", bool pblnCloseDocument = false, string docpasswordProtect = "", bool pblnRestoreCommandBar = true) //added docpasswordProtect,docId
        {
            const string MyProcName = "Document_Save_FileAndDatabase";            
            string strDocName;
            Template template;
            string strProtectionPassword = "";
            string strDocumentCmdBarName = "";

            long SavePosition = 0;
            int SaveRowIndex = 0;
            int SaveLeft = 0;
            bool SaveVisible = false;
            bool DocTrackChgsReq;
            //Todo //   qdcinstance clsDoc;
            bool blnRestoreRevView;
            int intVersion;
            string templatePath;
            string DocPasswordProtect = "";
            bool wasProtectedAtStart = false;
            try
            {//RPC error pannej02
                if (pdocCurrent == null)
                {
                    throw new ArgumentNullException(nameof(pdocCurrent), "Target document is null.");
                }

                //blnRestoreRevView = pdocCurrent.ActiveWindow.View.ShowRevisionsAndComments;
                blnRestoreRevView = ExecuteWordComWithRetry(() => pdocCurrent.ActiveWindow.View.ShowRevisionsAndComments);

                var documentSelect = _documentManager.Document_SelectQPSObject(pdocCurrent.Name, null, true); //   28012025
                if (documentSelect?.GetClsDocument == null)
                {
                    throw new InvalidOperationException("Unable to load document metadata from database.");
                }

                DocPasswordProtect = documentSelect.GetClsDocument.DocPasswordProtect;
                strProtectionPassword = DocPasswordProtect;
                //RPC error pannej02
                //wasProtectedAtStart = pdocCurrent.ProtectionType != WdProtectionType.wdNoProtection;
                wasProtectedAtStart = ExecuteWordComWithRetry(() => pdocCurrent.ProtectionType != WdProtectionType.wdNoProtection);

                if (wasProtectedAtStart)//   28012025
                {
                    CommonService.DocumentProtection_Off(pdocProtectedDocument: pdocCurrent, pstrProtectionPassword: DocPasswordProtect, pdocCurrent.Name);
                }
                //RPC error pannej02
                ExecuteWordComWithRetry(() =>
                {
                    pdocCurrent.ActiveWindow.View.ShowRevisionsAndComments = false;
                    pdocCurrent.ActiveWindow.View.RevisionsView = WdRevisionsView.wdRevisionsViewFinal;
                });

                if (pstrSaveAsFileName == "")
                {
                    strDocName = pdocCurrent.Name;

                }
                else
                {
                    strDocName = pstrSaveAsFileName;

                }
                if (strDocName.EndsWith(".dotx"))
                {
                    if (MessageBox.Show("OK to SAVE the Template '" + strDocName + "'?", "FileSave in " + /*gstrAppAcronym +*/  " Template", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        //RPC error pannej02
                        //pdocCurrent.Save();
                        ExecuteWordComWithRetry(() => pdocCurrent.Save());
                    }

                    return;
                }
                foreach (Template tpt in pdocCurrent.Application.Templates)
                {
                    if (tpt.Name.IndexOf("MBR", StringComparison.OrdinalIgnoreCase) >= 0) // 2301
                    {
                        tpt.Saved = true;
                    }
                }
                //RPC error pannej02
                //template = pdocCurrent.get_AttachedTemplate();//Need to check
                
                template = ExecuteWordComWithRetry(() => pdocCurrent.get_AttachedTemplate());//Need to check
                template.Saved = true;
                formManager.Wizard_BookMarkRemoval(pdocCurrent);

                if (pblnRestoreCommandBar)
                {
                    //RPC error pannej02
                    ExecuteWordComWithRetry(() =>
                    {
                        foreach (CommandBar command in pdocCurrent.CommandBars)
                        {
                            string commandName = command.Name;
                            if (!string.IsNullOrEmpty(commandName) && commandName.Length > 5)
                            {
                                if (commandName.StartsWith("QPSNT", StringComparison.OrdinalIgnoreCase))
                                {
                                    if (command.Visible)
                                    {
                                        strDocumentCmdBarName = commandName;
                                        SavePosition = (long)command.Position;
                                        SaveRowIndex = command.RowIndex;
                                        SaveLeft = command.Left;
                                        SaveVisible = command.Visible;
                                        break;
                                    }
                                    
                                }
                            }
                        }
                    });
                }
                
                //RPC error pannej02
                //string version = pdocCurrent.Application.Version;
                string version = ExecuteWordComWithRetry(() => pdocCurrent.Application.Version);
                intVersion = int.Parse(version.Substring(0, version.IndexOf('.')));
                if (intVersion >= CommonService.gintWordVersion2007)
                {

                    templatePath = CommonService.Get_Template_Name(pdocCurrent.Application, TemplateNameFormat.tnfReview) + "m";
                }
                else
                {
                    templatePath = CommonService.Get_Template_Name(pdocCurrent.Application, TemplateNameFormat.tnfReview);
                }
                //RPC error pannej02
                //Template saveFlowTemplate = pdocCurrent.get_AttachedTemplate();
                Template saveFlowTemplate = ExecuteWordComWithRetry(() => pdocCurrent.get_AttachedTemplate());
                string currentTemplateFullName = saveFlowTemplate?.FullName ?? string.Empty;
                bool reviewTemplateNeededForSave =
                    !string.IsNullOrEmpty(templatePath) &&
                    !string.Equals(currentTemplateFullName, templatePath, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(currentTemplateFullName, CommonService.templateFullName, StringComparison.OrdinalIgnoreCase);

                if (reviewTemplateNeededForSave)
                {
                    //RPC error pannej02
                    //pdocCurrent.set_AttachedTemplate(templatePath);
                    ExecuteWordComWithRetry(() => pdocCurrent.set_AttachedTemplate(templatePath));
                }

                if (pstrSaveAsFileName == "")  // NewCode
                {
                    //System.Threading.Tasks.Task.Run(() =>
                    //{
                    //pdocCurrent.Application.Options.SaveInterval = 0;

                    //RPC error pannej02
                    //pdocCurrent.Save();
                    ExecuteWordComWithRetry(() => pdocCurrent.Save());
                    //}); 

                }
                else
                {
                    //RPC error pannej02
                    //pdocCurrent.SaveAs(FileName: pstrSaveAsFileName, Password: pstrDocPasswordRead, WritePassword: pstrDocPasswordWrite);
                    ExecuteWordComWithRetry(() => pdocCurrent.SaveAs(FileName: pstrSaveAsFileName, Password: pstrDocPasswordRead, WritePassword: pstrDocPasswordWrite));
                }
                //RPC error pannej02
                //if (pdocCurrent.TrackRevisions == true)
                //{
                //    pdocCurrent.TrackRevisions = false;
                //}
                if (ExecuteWordComWithRetry(() => pdocCurrent.TrackRevisions) == true)
                {
                    ExecuteWordComWithRetry(() => pdocCurrent.TrackRevisions = false);
                }

                //var documentSelect = _documentManager.Document_SelectQPSObject(pdocCurrent.Name, null, true);   280125
                formManager.DocumentSaved_DB_Record(documentSelect.GetClsDocument.DocID, pdocCurrent);//(Document_SelectID(ActiveDocument.Name))
                DocTrackChgsReq = documentSelect.GetClsDocument.mblnTrackChangesReq;

                if (DocTrackChgsReq == true)
                {
                    //RPC error pannej02
                    //if (pdocCurrent.TrackRevisions == false)
                    //{
                    //    pdocCurrent.TrackRevisions = true;
                    //}
                    if (ExecuteWordComWithRetry(() => pdocCurrent.TrackRevisions) == false)
                    {
                        ExecuteWordComWithRetry(() => pdocCurrent.TrackRevisions = true);
                    }
                }

                if (!pblnCloseDocument)
                {
                    templatePath = CommonService.templateFullName; // 2301
                    bool templateRefreshNeeded = true;
                    //RPC error pannej02
                    //Template currentTemplate = pdocCurrent.get_AttachedTemplate();
                    Template currentTemplate = ExecuteWordComWithRetry(() => pdocCurrent.get_AttachedTemplate());
                    if (currentTemplate != null && !string.IsNullOrEmpty(currentTemplate.FullName))
                    {
                        templateRefreshNeeded = !string.Equals(currentTemplate.FullName, templatePath, StringComparison.OrdinalIgnoreCase);
                    }

                    if (templateRefreshNeeded)
                    {
                        //RPC error pannej02
                        //pdocCurrent.set_AttachedTemplate(templatePath);
                        ExecuteWordComWithRetry(() => pdocCurrent.set_AttachedTemplate(templatePath));
                    }
                }
                if (DatabaseHelper.Transaction_Active() == strDocName)
                {
                    DatabaseHelper.Transaction_Commit(strDocName);

                }
                // documentSelect.ScaledValues = intScaledValues
                if (pblnLockDocument)
                {

                    documentSelect.GetClsDocument.DocLocation = CommonService.gintDocLocation_LOCK;
                }

                if (pstrDocPasswordWrite != "")
                {
                    documentSelect.GetClsDocument.DocPasswordWrite = pstrDocPasswordWrite;
                }
                //.SaveToDatabase() //Need to check
                if (pblnCloseDocument)
                {
                    //RPC error pannej02
                    //pdocCurrent.Close();
                    ExecuteWordComWithRetry(() => pdocCurrent.Close());
                }
                else
                {
                    if (pblnRestoreCommandBar && strDocumentCmdBarName != "")
                    {
                        Globals.ThisDocument.CommandBar_Visible(strDocumentCmdBarName, SavePosition, (MsoBarRow)SaveRowIndex, SaveLeft, SaveVisible);

                        bool needsCommandBarBuild = true;
                        //RPC error pannej02
                        ExecuteWordComWithRetry(() =>
                        {
                            foreach (CommandBar command in pdocCurrent.Application.CommandBars)
                            {
                                if (string.Equals(command.Name, strDocumentCmdBarName, StringComparison.Ordinal))
                                {
                                    needsCommandBarBuild = false;
                                    break;
                                }
                            }
                        });

                        if (needsCommandBarBuild)
                        {
                            Globals.ThisDocument.CommandBar_Build(_tblOpUnit.OpUnitType, _opUnitDocType.DocTypeAbb, _tblOpUnit.OpUnitID, _opUnitDocType.DocTypeID);
                        }
                    }
                    //RPC error pannej02
                    //pdocCurrent.Saved = true;
                    ExecuteWordComWithRetry(() => pdocCurrent.Saved = true);

                }
                // pdocCurrent.ActiveWindow.View.ShowRevisionsAndComments = blnRestoreRevView;// gettinr error: "This method or property is not available because the document is a protected document."}

                if (!pblnCloseDocument)
                {
                    //RPC error pannej02
                    //pdocCurrent.ActiveWindow.View.ShowRevisionsAndComments = blnRestoreRevView;// gettinr error: "This method or property is not available because the document is a protected document."}
                    ExecuteWordComWithRetry(() => pdocCurrent.ActiveWindow.View.ShowRevisionsAndComments = blnRestoreRevView);// gettinr error: "This method or property is not available because the document is a protected document."}
                    if (wasProtectedAtStart)
                    {
                        CommonService.DocumentProtection_On(pdocCurrent, DocPasswordProtect);
                    }
                }
                // open the doc and save the count it in DB*
            }

            catch (Exception ex)
            {
                if (ex.Data["errNum"] != null)  // NewCode1
                {
                    CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                }
                else
                {
                    var stackTrace = new StackTrace(ex, true);
                    var frame = stackTrace.GetFrame(0);
                    int lineNumber = frame?.GetFileLineNumber() ?? -1;
                    CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName, pintErl: lineNumber);
                }
            }

        }
        // Retry logic for Word COM calls to handle RPC call rejected errors
        //RPC error pannej02
        private static void ExecuteWordComWithRetry(Action action, int retries = WordComRetryCount, int delayMs = WordComRetryDelayMs)
        {
            for (int attempt = 0; attempt < retries; attempt++)
            {
                try
                {
                    action();
                    return;
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == RpcCallRejectedErrorCode)
                {
                    if (attempt == retries - 1)
                    {
                        throw;
                    }

                    Thread.Sleep(delayMs);
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
                catch (COMException ex) when ((uint)ex.ErrorCode == RpcCallRejectedErrorCode)
                {
                    if (attempt == retries - 1)
                    {
                        throw;
                    }

                    Thread.Sleep(delayMs);
                }
            }

            throw new InvalidOperationException("Failed to execute Word COM call after retries.");
        }

        

    }
}
