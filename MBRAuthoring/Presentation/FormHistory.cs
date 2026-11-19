using MBRAuthoring.Business;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace MBRAuthoring.Presentation
{
    public partial class FormHistory : FormBase
    {
        public IDocumentManager _documentManager = null; //ServiceLocator.GetDocumentManager();
        private Word.Document currentActiveDocument;

        public FormHistory()
        {
            InitializeComponent();
            _documentManager = new DocumentManager();

        }
        private void FormHistory_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            currentActiveDocument = Globals.ThisDocument.ActiveDocument;//currentDocument;
            this.Text = ModChangeHistory.gstrAppAcronym + " Show Document 'Save' History";
            var result = _documentManager.Document_SelectAttributes(currentActiveDocument.Name);
            Fill_DocImageLogRow(result.DocumentID);
            Cursor = Cursors.Default;
            currentActiveDocument.Application.StatusBar = "";

        }

        private void Fill_DocImageLogRow(int docId)
        {
            SetToolTip(grdImageLog, "grdImageLog: DocID=" + docId + " ImageLog");  // NewCode
            var resultSet = _formManager.Fill_DocImageLog(docId);
            grdImageLog.DataSource = resultSet;

            grdImageLog.Columns[0].HeaderText = "Date/Time";
            grdImageLog.Columns[0].Width = 130;
            grdImageLog.Columns[0].DefaultCellStyle.Format = "M/d/yyyy h:mm:ss tt";

            grdImageLog.Columns[1].HeaderText = "User";
            grdImageLog.Columns[1].Width = 90;

            grdImageLog.Columns[2].HeaderText = "Pages";
            grdImageLog.Columns[2].Width = 45;

            grdImageLog.Columns[3].HeaderText = "Sects";
            grdImageLog.Columns[3].Width = 45;

            grdImageLog.Columns[4].HeaderText = "Paras";
            grdImageLog.Columns[4].Width = 45;

            grdImageLog.Columns[5].HeaderText = "Words";
            grdImageLog.Columns[5].Width = 45;

            grdImageLog.Columns[6].HeaderText = "Chars";
            grdImageLog.Columns[6].Width = 45;

            grdImageLog.Columns[7].HeaderText = "Tables";
            grdImageLog.Columns[7].Width = 45;

            grdImageLog.Columns[8].HeaderText = "Props";
            grdImageLog.Columns[8].Width = 45;

            grdImageLog.Columns[9].HeaderText = "Fields";
            grdImageLog.Columns[9].Width = 45;

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();
        }


    }
}
