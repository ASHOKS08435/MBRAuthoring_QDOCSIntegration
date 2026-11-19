using MBRAuthoring.Business;
using MBRAuthoring.DataAccess;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace MBRAuthoring.Presentation
{
    public partial class FormAbout : FormBase
    {


        Word.Application WordApp = new Word.Application();
        public FormAbout()
        {
            InitializeComponent();
        }

        private void FormAbout_Load(object sender, EventArgs e)
        {
            const string MyProcName = "UserForm_Initialize";
            lblTitle.Text = "MBR Authoring";
            lblDate.Text = "16-APR-2025";
            lblVersion.Text = "2.0";
            lblDocumentPath.Text = DatabaseHelper.OpUnitDocType_DocumentPath_Select();
            lblTemplatePath.Text = CommonService.templateFullName;//CommonService.Get_Template_Name(Globals.ThisDocument.ActiveDocument.Application, TemplateNameFormat.tnfNameAndPath);
            lblUserType.Text = DatabaseHelper.UserData.User.Type;
            lblEnvironment.Text = DatabaseHelper.AppStatus_Select("Environment");
            lblDBServer.Text = DatabaseHelper.dbDetails.SERVERNAME;
            lblDBName.Text = DatabaseHelper.dbDetails.DBNAME;
            string strversion = DatabaseHelper.dbDetails.VERSION;
            lblDBVersion.Text = strversion.Split('\n')[0];
            lblWordVersion.Text = WordApp.Version + "  (Build " + WordApp.Build + ")";

        }


        private void cmdDBStats_Click(object sender, EventArgs e)
        {


            DialogResult result = MessageBox.Show("Database Connected at: " + DatabaseHelper.dbConnectionDetails.ConnectStartTime + "\r\n" +
                                    "Database Disconnect at: " + DatabaseHelper.dbConnectionDetails.ConnectEndTime + "\r\n" +
                                    "Timed Disconnect at: " + DatabaseHelper.gdatDatabaseConnectOnTime + "\r\n" +
                                       "Current Date/Time: " + DateTime.Now + "\r\n" +
                                       "Database Open's: " + DatabaseHelper.dbConnectionDetails.DatabaseConnections + "\r\n" +
                                        "Data Requests In: " + DatabaseHelper.dbConnectionDetails.DatabaseRequestsIn + "\r\n" +
                                        "Database Requests Out: " + DatabaseHelper.dbConnectionDetails.DatabaseRequestsOut + "\r\n" + "\r\n" +
                                        "       Reset Request Counters ? ", "Database Statistics", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                DatabaseHelper.dbConnectionDetails.DatabaseRequestsIn = 0;
                DatabaseHelper.dbConnectionDetails.DatabaseRequestsOut = 0;
            }
            else
            {

            }
        }

        private void cmdClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
