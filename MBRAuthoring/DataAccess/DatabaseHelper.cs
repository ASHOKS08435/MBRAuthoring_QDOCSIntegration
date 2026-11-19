using MBRAuthoring.Business;
using MBRAuthoring.Models;
using MBRAuthoring.Utilities;
using MBRAuthoring.Utilities.CustomAttributes;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Timers;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;

namespace MBRAuthoring.DataAccess
{
    public static class DatabaseHelper
    {
        //private static string connectionString;
        //private static string logFilePath;
        static string mstrDBTransactionName = "";
        const int DocID_ResetRequest = -47;
        const string cLookupTypeDefault = "LookupTypeDefault";
        const string cLookupText = "LookupText";
        const string cLookupAttributeDefinition = "LookupAttributeDefinition";
        const string cLookupAttributeValue = "LookupAttributeValue";
        public static UserDataResult UserData;
        public static DBDetails dbDetails;
        public static DBConnectionDetails dbConnectionDetails;
        private static SqlConnection macnDatabase;
        private static SqlTransaction transaction;
        const string MyCompName = "DatabaseHelper";
        static string mstrProvider =string.Empty;
        static string gstrDatabaseConnectFinish;
        static bool gblnDatabaseTimerIsSet=false;
        public static DateTime gdatDatabaseConnectOnTime;
        //public static bool gblnUserType_ADMIN;
        //public static bool gblnUserType_REVIEWER;
        //public const string userType_Admin = "ADMIN";
        //public const string userType_REVIEWER = "REVIEWER";


        //public DatabaseHelper(string connectionString, string logFilePath)
        //{
        //    this.connectionString = connectionString;
        //    this.logFilePath = logFilePath;
        //}

        public static void Transaction_Begin(string documentName)
        {

            if (!string.IsNullOrEmpty(mstrDBTransactionName))
            {
                MessageBox.Show("ERROR: TRANSACTION '" + mstrDBTransactionName + "' ALREADY ACTIVE; " + " CANNOT START TRANSACTION '" + documentName + "'", MyCompName + ". Transaction_Begin", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            Connection_Assure();
            transaction = macnDatabase.BeginTransaction();
            mstrDBTransactionName = documentName;
        }

        public static void Transaction_Commit(string DocumentName)
        {
            const string MyProcName = "Transaction_Commit";
            if (string.IsNullOrEmpty(mstrDBTransactionName))
            {
                MessageBox.Show("ERROR: NO TRANSACTION ACTIVE; " + " CANNOT COMMIT '" + DocumentName + "'", MyCompName + "." + MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (mstrDBTransactionName != DocumentName)
            {
                MessageBox.Show("ERROR: TRANSACTION IS '" + mstrDBTransactionName + "', NOT '" + DocumentName + "'", MyCompName + "." + MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            transaction.Commit();
            mstrDBTransactionName = "";
        }

        public static void Transaction_Rollback(string DocumentName)
        {
            const string MyProcName = "Transaction_Rollback";
            if (string.IsNullOrEmpty(mstrDBTransactionName)) { return; }

            if (mstrDBTransactionName != DocumentName)
            {
                MessageBox.Show("ERROR: TRANSACTION IS '" + mstrDBTransactionName + "', NOT '" + DocumentName + "'", MyCompName + "." + MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            transaction.Rollback();
            mstrDBTransactionName = "";
        }

        public static string Transaction_Active()
        {
            const string MyProcName = "Transaction_Active";
            return mstrDBTransactionName;
        }

        private static void Connection_Assure()
        {
            const string MyProcName = "Connection_Assure";
            try
            {
                if (macnDatabase == null)
                {
                    Connection_Open();
                }
                else
                {
                    if (macnDatabase.State != ConnectionState.Open)
                    {
                        Connection_Open();
                    }
                }
                Connection_TimerSet();
                return;
            }
            catch (Exception ex)
            {
                CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, "Connection_Assure", false);
            }
        }
        public static void Connection_Timer(object sender, ElapsedEventArgs e)
        {
            gblnDatabaseTimerIsSet = false;
            if (!string.IsNullOrEmpty(mstrDBTransactionName))
            {
                Connection_TimerSet();
                return;
            }
            if (!(macnDatabase == null))
            {
                macnDatabase.Close();
            }
            gstrDatabaseConnectFinish = DateTime.Now.ToString();
        }

        private static void Connection_TimerSet()
        {
            gblnDatabaseTimerIsSet = true;
            gdatDatabaseConnectOnTime = DateTime.Now.AddMinutes(15);  // NewCode
            TimeSpan timeToWait = gdatDatabaseConnectOnTime - DateTime.Now;
            Globals.ThisDocument.timer = new System.Timers.Timer(timeToWait.TotalMilliseconds);
            Globals.ThisDocument.timer.Elapsed += Connection_Timer;
        }

        public static void ErrorLog_Insert(string plngErrorNum, string pstrErrorDesc, string pstrErrorModule, string pstrErrorProc)
        {
            const string MyProcName = "ErrorLog_Insert";
            int intRecordsAffected = 0;
            string strUserName;
            string strErrorApp;
            string strErrorDocName = "[Unavailable]";
            try
            {
                dbConnectionDetails.DatabaseRequestsIn++;
                Connection_Assure();
                if (!string.IsNullOrEmpty(mstrDBTransactionName))
                {
                    MessageBox.Show("CAUTION: An Error Message is being logged to the Database" + Environment.NewLine + "\t" + "while Transaction '" + mstrDBTransactionName + "' is Active." + Environment.NewLine + Environment.NewLine + "\t" + "This Error Message WILL BE LOST IF the Transaction is discarded.", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                string spName = "SP_InsertErrorLog";
                strUserName = Environment.UserName;
                strErrorApp = ModChangeHistory.gstrAppAcronym + " Authoring " + ModChangeHistory.gstrAppVersionNumber;

                try
                {
                    var activeDoc = Globals.ThisDocument?.ActiveDocument;
                    if (activeDoc != null)
                    {
                        strErrorDocName = activeDoc.Name;
                    }
                }
                catch
                {
                    strErrorDocName = "[Unavailable]";
                }

                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {

                    using (SqlCommand command = new SqlCommand(spName, connection))
                    {
                        connection.Open();
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@userName", StrFormatter(strUserName));
                        command.Parameters.AddWithValue("@errApp", StrFormatter(strErrorApp));
                        command.Parameters.AddWithValue("@errNum", plngErrorNum);
                        command.Parameters.AddWithValue("@errDesc", StrFormatter(pstrErrorDesc));
                        command.Parameters.AddWithValue("@errModule", StrFormatter(pstrErrorModule));
                        command.Parameters.AddWithValue("@errProc", StrFormatter(pstrErrorProc));
                        command.Parameters.AddWithValue("@errDocName", StrFormatter(strErrorDocName));

                        SqlParameter returnValueParam = new SqlParameter();
                        returnValueParam.ParameterName = "@ReturnValue";
                        returnValueParam.SqlDbType = SqlDbType.Int;
                        returnValueParam.Direction = ParameterDirection.ReturnValue;
                        command.Parameters.Add(returnValueParam);

                        intRecordsAffected = command.ExecuteNonQuery();

                        connection.Close();

                        if (intRecordsAffected != 1)
                        {
                            MessageBox.Show("Unexpected RecordsAffected= " + intRecordsAffected, MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errLineNum = ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' '));
                MessageBox.Show("Error Number: " + "\t" + ex.HResult.ToString() + "\r\n" + "Line Number: " + "\t" + errLineNum + "\r\n" + "Error: " + "\t\t" + ex.Message, "Error on ErrorLog_Insert", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        public static (List<T>, Dictionary<string, object>) ExecuteStoredProcedure<T>(string storedProcedureName, SqlParameter[] parameters, Func<IDataReader, T> mapper, out Dictionary<string, object> outputValues, bool useTransaction = true)
        {
            List<T> result = new List<T>();
            outputValues = new Dictionary<string, object>();

            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                connection.Open();

                // Begin transaction if requested
                SqlTransaction transaction = null;
                if (useTransaction)
                {
                    transaction = connection.BeginTransaction();
                }

                try
                {
                    SqlCommand command = connection.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = storedProcedureName;
                    command.Transaction = transaction;

                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                    }

                    // Execute the command
                    if (IsStoredProcedureWithResultSet(storedProcedureName))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            try
                            {

                                //command.ExecuteNonQuery();
                                //string value = command.Parameters["@ErrorReturn"].Value.ToString();
                                if (reader.Read())
                                {
                                    while (reader.Read())
                                    {
                                        T entity = mapper(reader);
                                        result.Add(entity);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }

                        }

                    }
                    else
                    {
                        command.ExecuteNonQuery();
                    }

                    foreach (SqlParameter parameter in command.Parameters)
                    {
                        if (parameter.Direction == ParameterDirection.Output || parameter.Direction == ParameterDirection.InputOutput)
                        {
                            outputValues[parameter.ParameterName] = parameter.Value;
                        }
                    }

                    // Commit transaction if successful and if transaction is used
                    if (useTransaction)
                    {
                        transaction.Commit();
                    }

                    // Log success
                    Log($"Stored procedure '{storedProcedureName}' executed successfully.");
                }
                catch (Exception ex)
                {
                    // Rollback transaction if transaction is used
                    if (useTransaction)
                    {
                        transaction.Rollback();
                    }

                    // Log error
                    Log($"Error executing stored procedure '{storedProcedureName}': {ex.Message}");
                    throw; // Rethrow exception
                }
            }

            return (result, outputValues);
        }

        private static bool IsStoredProcedureWithResultSet(string storedProcedureName)
        {
            // You may implement a logic to determine whether the stored procedure returns a result set
            // For simplicity, let's assume all stored procedures starting with "Get" return result sets
            return storedProcedureName.StartsWith("SP_");
        }

        private static void Log(string message)
        {
            string logFilePath ="";
            try
            {
                using (StreamWriter writer = File.AppendText(logFilePath))
                {
                    writer.WriteLine($"{DateTime.Now} - {message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error logging: {ex.Message}");
            }
        }

        #region FileNew
        public static List<TblOpUnit> OpUnit_SelectALL()
        {
            int? userID = UserData != null ? UserData.User.UserID : 0;
            List<TblOpUnit> opUnitList = new List<TblOpUnit>();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {

                using (SqlCommand command = new SqlCommand("SP_OpUnitSelectALL", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("userID", userID);
                    try
                    {
                        connection.Open();
                        dbConnectionDetails.DatabaseRequestsIn++;

                        using (SqlDataReader reader = command.ExecuteReader())
                        {

                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    //TblOpUnit opUnit = new TblOpUnit();
                                    //var props = typeof(TblOpUnit).GetProperties();
                                    //foreach (var prop in props)
                                    //{
                                    //    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                    //    {
                                    //        prop.SetValue(opUnit, reader[prop.Name]);
                                    //    }
                                    //}

                                    TblOpUnit opUnit = new TblOpUnit
                                    {
                                        OpUnitID = Convert.ToInt32(reader["OpUnitID"]),
                                        OpUnit = reader["OpUnit"].ToString(),
                                        OpUnitType = reader["OpUnitType"].ToString(),
                                        OUBuilding = reader["OUBuilding"].ToString(),
                                        WorkCenterOpUnitID = Convert.ToInt32(reader["WorkCenterOpUnitID"]),
                                        MaterialOpUnitID = Convert.ToInt32(reader["MaterialOpUnitID"]),
                                    };

                                    //OpUnit opUnit = new OpUnit()
                                    //{
                                    //    OpUnitID = Convert.ToInt32(reader["OpUnitID"]),
                                    //    OpUnitName = reader["OpUnit"].ToString(),
                                    //    OpUnitType = reader["OpUnitType"].ToString(),
                                    //    OUBuilding = reader["OUBuilding"].ToString(),
                                    //    WorkCenterOpUnitID = Convert.ToInt32(reader["WorkCenterOpUnitID"]),
                                    //    MaterialOpUnitID = Convert.ToInt32(reader["MaterialOpUnitID"])
                                    //};

                                    opUnitList.Add(opUnit);

                                }
                            }
                            connection.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }

            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return opUnitList;

        }

        private static void RecordSet_Assure()
        {
            Connection_Assure();
        }
        public static int StepText_SelectCOUNT()
        {
            int stepTextCount = 0;
            dbConnectionDetails = new DBConnectionDetails();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            if (dbConnectionDetails.DatabaseConnections != 0)
            {

                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand("sp_StepTextCOUNT", connection))
                    {
                        try
                        {
                            command.CommandType = CommandType.StoredProcedure;
                            connection.Open();
                            stepTextCount = (int)command.ExecuteScalar();
                            connection.Close();
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return stepTextCount;
        }

        public static List<TblStepText> StepText_Select(string stepTextName)
        {

            List<TblStepText> stepTextList = new List<TblStepText>();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand command = new SqlCommand("SP_StepTextSelect", connection))
                {
                    try
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@stepTextName", stepTextName);
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    stepTextList.Add(new TblStepText
                                    {
                                        StepName = reader["StepName"].ToString(),
                                        StepText = reader["StepText"].ToString(),
                                    });
                                }
                            }
                            connection.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return stepTextList;
        }

        private static void Connection_Open()
        {
            const string MyProcName = "Connection_Open";
            try
            {
                string strTemplateFullName;
                string strUDLFullName;

                mstrDBTransactionName = "";
                if (macnDatabase == null)
                {
                    macnDatabase = new SqlConnection(AppConfig.ConnString);
                }
                // macnDatabase.CursorLocation = adUseClient; // not req

                strTemplateFullName = CommonService.Get_Template_Name(Globals.ThisDocument.ActiveDocument.Application, TemplateNameFormat.tnfNameAndPath);
                if (strTemplateFullName.Substring(strTemplateFullName.Length - 5).ToUpper() != ".DOTX")
                {
                    CommonService.Error_Handler("0", "Invalid extension '" + strTemplateFullName.Substring(strTemplateFullName.Length - 4) + "', expecting '.dotx' to include Add-ins", MyCompName, MyProcName, false);

                    return;
                }

                macnDatabase.Open();
                dbConnectionDetails.DatabaseConnections++;
                dbConnectionDetails.ConnectStartTime = DateTime.Now;
                dbConnectionDetails.DatabaseRequestsIn++;
                dbConnectionDetails.DatabaseRequestsOut++;
                mstrProvider = "UNKNOWN"; // Not req
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {

                    using (SqlCommand command = new SqlCommand("SP_ConnectionOpen", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    dbDetails = new DBDetails
                                    {
                                        SERVERNAME = reader["SERVERNAME"].ToString(),
                                        DBNAME = reader["DBNAME"].ToString(),
                                        VERSION = reader["VERSION"].ToString(),
                                        UserDetails = GetUserDetails()
                                    };
                                    UserData = dbDetails.UserDetails;
                                }
                            }
                            else
                            {
                                dbDetails = new DBDetails
                                {
                                    SERVERNAME = "UNKNOWN",
                                    DBNAME = "UNKNOWN",
                                    VERSION = "UNKNOWN"
                                };
                            }
                            connection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName, false);
            }
        }

        public static List<TblLookupData> LookupData_Select(string lookupType, string lookupDefault = "", string orderBy = "")
        {
            const string MyProcName = "LookupData_Select";
            bool blnLookupType_VALID = false;
            string strLookupAttributeDefinition = "";
            string strLookupDefaultText = "";
            string strDateSampleColumn = "";
            dbConnectionDetails.DatabaseRequestsIn++;
            var arsLookupType = LookupType_Select();
            //var lookupTypeItem = arsLookupType
            //    .FirstOrDefault(item => string.Equals(item.LookupType.Trim(), lookupType, StringComparison.OrdinalIgnoreCase));
            //Need to close app with do while of arsLookupType
            //if (lookupTypeItem != null)
            /*if (lookupTypeItem.LookupType.Trim().ToString().ToUpper() == lookupType.ToUpper())//070125
            {
                blnLookupType_VALID = true;
                strLookupAttributeDefinition = lookupTypeItem.LookupAttributeDefinition;
                strLookupDefaultText = lookupTypeItem.LookupText;
            }

            if (!blnLookupType_VALID)
            {
                MessageBox.Show("INVALID LookupType='" + lookupType + "'", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;  // NewCode2
            }*/

            foreach (var lookupTypeItem in arsLookupType)
            {
                if (string.Equals(lookupTypeItem.LookupType.Trim(), lookupType, StringComparison.OrdinalIgnoreCase))
                {
                    blnLookupType_VALID = true;
                    strLookupAttributeDefinition = lookupTypeItem.LookupAttributeDefinition;
                    strLookupDefaultText = lookupTypeItem.LookupText;
                    break;
                }
            }

            if (!blnLookupType_VALID)
            {
                MessageBox.Show("INVALID LookupType='" + lookupType + "'", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;  // NewCode2 throw exception
            }


            if (string.Equals(lookupType, "DATEFORMAT", StringComparison.OrdinalIgnoreCase))
            {
                strDateSampleColumn = ", dbo.MBR_DateAs(getdate(), LookupText)";
            }
            RecordSet_Assure();
            var sqlBuilder = new StringBuilder();
            sqlBuilder.Append("SELECT LookupText ");
            sqlBuilder.Append(strDateSampleColumn);
            sqlBuilder.Append(GenericUtility.Cma(GenericUtility.IsNullAs(cLookupAttributeValue, "' '")));
            sqlBuilder.Append(" FROM tblLookupData ");

            if (!string.IsNullOrEmpty(lookupType))
            { // check
                sqlBuilder.Append($" WHERE LookupType ='{lookupType}'");
            }

            if (!string.IsNullOrEmpty(orderBy))
            {
                sqlBuilder.Append($" ORDER BY {orderBy}");
            }

            var lookupResults = new List<TblLookupData>();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand command = new SqlCommand(sqlBuilder.ToString(), connection))
                {
                    command.Parameters.AddWithValue("@LookupType", lookupType);
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var lookupData = new TblLookupData
                            {
                                LookupText = reader["LookupText"].ToString(),
                            };
                            lookupResults.Add(lookupData);
                        }
                    }
                }
            }

            if (lookupResults.Count == 0)
            {
                MessageBox.Show("NO LOOKUP DATA IN DATABASE", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return lookupResults;
        }

        public static List<TblLookupType> LookupType_Select()
        {
            List<TblLookupType> lookupTypeList = new List<TblLookupType>();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            string query = "Select LT.LookupType" + GenericUtility.Cma(GenericUtility.IsNullAs(cLookupTypeDefault, "-1")) + GenericUtility.Cma(GenericUtility.IsNullAs(cLookupText, "''")) + Utilities.GenericUtility.Cma(Utilities.GenericUtility.IsNullAs(cLookupAttributeDefinition, "''")) + " from tblLookupType LT " + " LEFT OUTER JOIN tblLookupData LD " + " on LT.LookupTypeDefault = LD.LookupDataID ";
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    try
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    TblLookupType lookupType = new TblLookupType();
                                    var props = typeof(TblLookupType).GetProperties();
                                    foreach (var prop in props)
                                    {
                                        if (!reader.IsDBNull(reader.GetOrdinal(prop.Name.ToString())))
                                        {
                                            prop.SetValue(lookupType, reader[prop.Name].ToString());
                                        }
                                    }
                                    lookupTypeList.Add(lookupType);
                                }
                            }
                            connection.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return lookupTypeList;
        }

        public static List<Material> OpUnitProduct_SelectALL(int OpUnitID, bool isProductOrderByNumber)
        {
            List<Material> materials = new List<Material>();
            //string query = "Select DISTINCT MaterialID, MaterialTitle , ItemCode_MaterialTitle as MaterialDisplay , ItemCode FROM vewProduct  WHERE OpUnitID= @OpUnitID  ORDER BY ";

            //query = isProductOrderByNumber ? query + " ItemCode" : query + " MaterialTitle";
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {

                using (SqlCommand command = new SqlCommand("SP_OpUnitProductSelectALL", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@OpUnitID", OpUnitID);
                    command.Parameters.AddWithValue("@OrderBy", isProductOrderByNumber);
                    try
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {

                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    Material material = new Material()
                                    {
                                        MaterialID = Convert.ToInt32(reader["MaterialID"]),
                                        MaterialTitle = reader["MaterialTitle"].ToString(),
                                        MaterialDisplay = reader["MaterialDisplay"].ToString(),
                                        ItemCode = reader["ItemCode"].ToString(),
                                    };
                                    //var props = typeof(Material).GetProperties();
                                    //foreach (var prop in props)
                                    //{
                                    //    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                    //    {
                                    //        prop.SetValue(material, reader[prop.Name]);
                                    //    }
                                    //}


                                    //OpUnitDocType opUnit = new OpUnitDocType()
                                    //{
                                    //    DocTypeID = Convert.ToInt32(reader["DocTypeID"]),
                                    //    DocType = reader["DocType"].ToString(),
                                    //    DocTypeAbb = reader["DocTypeAbb"].ToString(),
                                    //    DocumentPath = reader["DocumentPath"].ToString(),
                                    //    ProjectNameUsed = reader["ProjectNameUsed"].ToString(),
                                    //    DocIsModule = reader["DocIsModule"].ToString(),
                                    //};

                                    materials.Add(material);

                                }
                            }
                            connection.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }

            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return materials;

        }



        public static List<DocBaseNameDetails> DocBaseName_SelectAll(int docTypeID, bool isDocOrderByBaseName)
        {
            List<DocBaseNameDetails> docBaseNameDetailsList = new List<DocBaseNameDetails>();
            //  string query = "SELECT DocumentBaseName, DocumentTitle FROM tblDocBaseName  WHERE Include=1 and DocTypeID= @DocTypeID ORDER BY ";

            // query = isDocOrderByBaseName ? query + " DocumentBaseName" : query + " DocumentTitle";

            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {

                using (SqlCommand command = new SqlCommand("SP_DocBaseNameSelectAll", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@DocTypeID", docTypeID);
                    command.Parameters.AddWithValue("@IsDocOrderByBaseName", isDocOrderByBaseName);
                    try
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {

                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    DocBaseNameDetails docBaseNameDetails = new DocBaseNameDetails();
                                    var props = typeof(DocBaseNameDetails).GetProperties();
                                    foreach (var prop in props)
                                    {
                                        if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                        {
                                            prop.SetValue(docBaseNameDetails, reader[prop.Name]);
                                        }
                                    }


                                    //OpUnitDocType opUnit = new OpUnitDocType()
                                    //{
                                    //    DocTypeID = Convert.ToInt32(reader["DocTypeID"]),
                                    //    DocType = reader["DocType"].ToString(),
                                    //    DocTypeAbb = reader["DocTypeAbb"].ToString(),
                                    //    DocumentPath = reader["DocumentPath"].ToString(),
                                    //    ProjectNameUsed = reader["ProjectNameUsed"].ToString(),
                                    //    DocIsModule = reader["DocIsModule"].ToString(),
                                    //};

                                    docBaseNameDetailsList.Add(docBaseNameDetails);

                                }
                            }
                            connection.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }

            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return docBaseNameDetailsList;

        }


        public static List<OpUnitDocType> OpUnitDocType_Select(int OpUnitID)
        {

            List<OpUnitDocType> opUnitDocTypeList = new List<OpUnitDocType>();
            //string query = "Select RT.DocTypeID, DocType, DocTypeAbb, DocumentPath , ProjectNameUsed, DocIsModule, DocHasModule, DocHasMaterial, DocHasWorkCenter,DocHasWorkCenterAlt,DocHasProduct ,DocHasProperty ,DocPermitsQuantitySufficient, DocNameIncludes_OpUnit ,DocNameIncludes_DocTypeAbb ,DocNameIncludes_WorkCenter ,DocNameIncludes_VerMajor ,DocNameIncludes_VerMinor ,DocNameIncludes_FreeText ,DocNameIncludes_DocBaseName,DocNameIncludes_ProductIdent from tblDocType RT, tblOpUnitDocType OUDT where RT.DocTypeID = OUDT.DocTypeID and OUDT.OpUnitID= @OpUnitID and OUDT.Include=1  Order by DocType ";
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {

                using (SqlCommand command = new SqlCommand("SP_OpUnitDocTypeSelect", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@OpUnitID", OpUnitID);
                    try
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {

                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    //OpUnitDocType opUnitDocType = new OpUnitDocType();
                                    //var props = typeof(OpUnitDocType).GetProperties();
                                    //foreach (var prop in props)
                                    //{
                                    //    if (Attribute.IsDefined(prop, typeof(NonDBBackedAttribute)))
                                    //    {
                                    //        continue;
                                    //    }

                                    //    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                    //    {
                                    //        prop.SetValue(opUnitDocType, reader[prop.Name]);
                                    //    }
                                    //}


                                    OpUnitDocType opUnitDocType = new OpUnitDocType()
                                    {
                                        DocTypeID = Convert.ToInt32(reader["DocTypeID"]),
                                        DocType = reader["DocType"].ToString(),
                                        DocTypeAbb = reader["DocTypeAbb"].ToString(),
                                        DocumentPath = reader["DocumentPath"].ToString(),
                                        ProjectNameUsed = reader["ProjectNameUsed"].ToString(),
                                        DocIsModule = reader["DocIsModule"].ToString(),
                                        DocHasModule = reader["DocHasModule"].ToString(),
                                        DocHasMaterial = reader["DocHasMaterial"].ToString(),
                                        DocHasWorkCenter = reader["DocHasWorkCenter"].ToString(),
                                        DocHasWorkCenterAlt = reader["DocHasWorkCenterAlt"].ToString(),
                                        DocHasProduct = reader["DocHasProduct"].ToString(),
                                        DocHasProperty = reader["DocHasProperty"].ToString(),
                                        DocPermitsQuantitySufficient = reader["DocPermitsQuantitySufficient"].ToString(),
                                        DocNameIncludes_OpUnit = Convert.ToInt32(reader["DocNameIncludes_OpUnit"]),
                                        DocNameIncludes_DocTypeAbb = Convert.ToInt32(reader["DocNameIncludes_DocTypeAbb"]),
                                        DocNameIncludes_WorkCenter = Convert.ToInt32(reader["DocNameIncludes_WorkCenter"]),
                                        DocNameIncludes_VerMajor = Convert.ToInt32(reader["DocNameIncludes_VerMajor"]),
                                        DocNameIncludes_VerMinor = Convert.ToInt32(reader["DocNameIncludes_VerMinor"]),
                                        DocNameIncludes_FreeText = Convert.ToInt32(reader["DocNameIncludes_FreeText"]),
                                        DocNameIncludes_DocBaseName = Convert.ToInt32(reader["DocNameIncludes_DocBaseName"]),
                                        DocNameIncludes_ProductIdent = Convert.ToInt32(reader["DocNameIncludes_ProductIdent"])
                                    };

                                    opUnitDocTypeList.Add(opUnitDocType);

                                }
                            }
                            connection.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }

            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return opUnitDocTypeList;

            //return DatabaseHelper.ExecuteQuery(query);


        }

        public static bool DocProperty_Apply(int docID, int propertyIndex, string propertyName, string propertyValue, string propertyUOM, int materialID, int workCenterID)
        {
            const string myProcName = "DocProperty_Apply";
            dbConnectionDetails.DatabaseRequestsIn++;
            Connection_Assure();
            DocProperty_Reset();

            string stroreProcedure = "SP_DocPropertyApply";
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand(stroreProcedure, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@DocID", docID);
                        command.Parameters.AddWithValue("@PropertyName", propertyName);
                        command.Parameters.AddWithValue("@PropertyIndex", propertyIndex);
                        command.Parameters.AddWithValue("@PropertyValue", propertyValue);
                        // if(!string.IsNullOrEmpty(propertyUOM))
                        // {
                        //command.Parameters.AddWithValue("@PropertyUOM", propertyUOM);
                        command.Parameters.AddWithValue("@PropertyUOM", string.IsNullOrEmpty(propertyUOM) ? (object)DBNull.Value : propertyUOM);
                        //}
                        //else
                        //{
                        // command.Parameters.AddWithValue("@PropertyUOM", string.IsNullOrWhiteSpace(propertyUOM) ? null : propertyUOM);
                        // }
                        command.Parameters.AddWithValue("@MaterialID", materialID >= 0 ? materialID : (object)DBNull.Value);
                        command.Parameters.AddWithValue("@WorkCenterID", workCenterID >= 0 ? workCenterID : (object)DBNull.Value);
                        // command.Parameters.AddWithValue("@MaterialID", materialID);
                        // command.Parameters.AddWithValue("@WorkCenterID", workCenterID);

                        //command.Parameters.Add("@ErrorMsg", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output;

                        SqlParameter errMessageParam = new SqlParameter("@ErrorMsg", SqlDbType.VarChar, 200);
                        errMessageParam.Direction = ParameterDirection.Output;
                        command.Parameters.Add(errMessageParam);

                        SqlParameter propertyInstance = new SqlParameter("@PropertyInstance", SqlDbType.VarChar, 37)
                        {
                            Direction = ParameterDirection.Output,
                            Value = DBNull.Value
                        };
                        command.Parameters.Add(propertyInstance);

                        SqlParameter displayValueParam = new SqlParameter("@PropertyValueDisplay", SqlDbType.VarChar, 288)
                        {
                            Direction = ParameterDirection.Output,
                            Value = DBNull.Value
                        };
                        command.Parameters.Add(displayValueParam);
                        connection.Open();
                        command.ExecuteNonQuery();

                    }
                    dbConnectionDetails.DatabaseRequestsOut++;
                    return true;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + "\r\n" + stroreProcedure);
                return false;

            }
        }

        public static DataTable Document_Property_Select(string docIsModule, string propertyCategory)
        {
            DataTable properties = new DataTable();
            string modulePredicate;
            string systemPredicate;
            string propertySource;

            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            //if (docIsModule == "Y")
            //{
            //    modulePredicate = " or PropertyName in ('ModuleTitle','ModuleTableRows')  ";
            //}
            //else
            //{
            //    modulePredicate = " and PropertyName not in ('ModuleTitle','ModuleTableRows')  ";
            //}
            //systemPredicate = " or PropertyName in ('ChangeNum','LotNbr','OrderNum','PrintDate','PrintTime','ProductID','ProductName','Manufacturing_Product_Code','NDC_Code','Product_Identification_Code','Global_Trade_ID_Number','RCS_Sales_Code','PrintPages','RecordType','FirstLotUsage','DocumentTitle') or PropertyName LIKE 'MOD\\_%' ESCAPE '\\' or PropertyName LIKE 'MTL\\_%' ESCAPE '\\' ";
            //propertySource = "Select PropertyName,PropertyDescription,PropertyDataType,PropertyDataTypeName,PropertyIndexMax,ISNULL(PropertyUOMLookupType, '') AS PropertyUOMLookupType, PropertyUsage, ISNULL(PropertyLookupType, '<NoLookup>') AS PropertyLookupType, PropertyCategory, PropertyTextMaxLen from tblProperty WHERE Include =1  and (  PropertyUsage in ('A','B') " + modulePredicate + systemPredicate + " ) ";
            //if (propertyCategory != "")
            //{
            //    propertySource = propertySource + " AND PropertyCategory=" + propertyCategory;
            //}
            //propertySource = propertySource + " Order By PropertyName ";

            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                connection.Open();
                //updated SP name GTIN PANNEJ02
                using (SqlCommand command = new SqlCommand("SP_PropertySelect", connection))
                //using (SqlCommand command = new SqlCommand("SP_PropertySelect_NEW", connection))
                {
                    try
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@docIsModule", docIsModule);
                        command.Parameters.AddWithValue("@propertyCategory", propertyCategory);
                        properties.Load(command.ExecuteReader());
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return properties;
        }

        private static void DocProperty_Reset()
        {
            DocProperty_Select(DocID_ResetRequest, false);
        }

        //Todo: Need to check 
        private static void DocProperty_Select(int docID, bool IsModuleRecord)
        {
            int sDocId;
            if (docID == DocID_ResetRequest)
            {
                sDocId = docID;
                return;
            }
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            string storedProcedure = "usp_GetDocProperty";
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand(storedProcedure, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@DocID", docID);
                        connection.Open();
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            dbConnectionDetails.DatabaseRequestsOut++;
        }

        private static void InsertUserDetailsInAudit(string userNTID)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {

                    using (SqlCommand command = new SqlCommand("SP_InsertUserDetails", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@NTUserName", userNTID);
                        command.Parameters.AddWithValue("@NTDomain", Environment.UserDomainName.ToUpper());
                        command.Parameters.AddWithValue("@PC_Name", Environment.MachineName.ToUpper());
                        try
                        {
                            connection.Open();
                            SqlParameter returnValueParam = new SqlParameter();
                            returnValueParam.ParameterName = "@ReturnValue";
                            returnValueParam.SqlDbType = SqlDbType.Int;
                            returnValueParam.Direction = ParameterDirection.ReturnValue;
                            command.Parameters.Add(returnValueParam);

                            int intRecordsAffected = command.ExecuteNonQuery();

                            connection.Close();
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                            connection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        public static UserDataResult GetUserDetails()
        {
            dbConnectionDetails.DatabaseRequestsIn++;
            dbConnectionDetails.DatabaseRequestsOut++;
            UserDataResult userDataResult = new UserDataResult();
            var userNTID = GenericUtility.GetCurrentUserNTID();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            //string query = "Select USR.OpUnitID, OpUnit, UserTypeID, UT.Type as UserType, USR.UserID From tblUser USR, tblOpUnit OU, tblUserType UT Where USR.OpUnitID=OU.OpUnitID   and  USR.UserTypeID = UT.TypeID and UserName= @userNTID";
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {

                using (SqlCommand command = new SqlCommand("SP_GetUserDetails", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@userNTID", userNTID);
                    try
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                bool hasNextRecord = reader.Read();

                                if (hasNextRecord)
                                {
                                    User user = new User()
                                    {
                                        UserID = Convert.ToInt32(reader["UserID"]),
                                        UserTypeID = Convert.ToInt32(reader["UserTypeID"]),
                                        Type = reader["UserType"].ToString(),
                                        OpUnitID = Convert.ToInt32(reader["OpUnitID"]),
                                        OpUnitName = reader["OpUnit"].ToString()
                                    };
                                    CommonService.gblnUserType_ADMIN = user.Type.ToUpper().Contains(CommonService.userType_Admin);
                                    CommonService.gblnUserType_REVIEWER = user.Type.ToUpper().Contains(CommonService.userType_REVIEWER);
                                    CommonService.gintUserID = (int)user.UserID;
                                    CommonService.gstrUserName = userNTID;

                                    InsertUserDetailsInAudit(userNTID);
                                    return new UserDataResult { User = user };

                                }
                                while (hasNextRecord = reader.Read())
                                {
                                    InsertUserDetailsInAudit(userNTID);
                                    return new UserDataResult { ErrMsg = $"ERROR:  MULTIPLE ROWS for USER {userNTID}" };
                                }
                            }
                            else
                            {
                                User user = new User()
                                {
                                    UserID = -1,
                                    Type = "UNKNOWN"
                                };
                                CommonService.gblnUserType_ADMIN = user.Type.ToUpper().Contains(CommonService.userType_Admin);
                                CommonService.gblnUserType_REVIEWER = user.Type.ToUpper().Contains(CommonService.userType_REVIEWER);

                                InsertUserDetailsInAudit(userNTID);
                                return new UserDataResult { User = user };
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            InsertUserDetailsInAudit(userNTID);
            return new UserDataResult();
        }

        
        public static BuildDocumentNameDetail BuildDocumentName(int opUnitID, int docTypeID, long materialID, long workCenterID, string freeText, string docBaseName)
        {
            BuildDocumentNameDetail buildDocumentName = new BuildDocumentNameDetail();

            string stroreProcedure = "SP_BuildDocumentName";
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand(stroreProcedure, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        SqlParameter[] parameters = new SqlParameter[]
                            {

                new SqlParameter("@OpUnitID",SqlDbType.Int){SqlValue =opUnitID },
                new SqlParameter("@DocTypeID",System.Data.SqlDbType.Int){SqlValue = docTypeID},
                new SqlParameter("@MaterialID",System.Data.SqlDbType.Int){SqlValue = materialID},
                new SqlParameter("@VersionMajor",System.Data.SqlDbType.Int){SqlValue =1 },
                new SqlParameter("@VersionMinor",System.Data.SqlDbType.Int){SqlValue = 0},
                new SqlParameter("@WorkCenterID",System.Data.SqlDbType.Int){SqlValue = workCenterID},
                new SqlParameter("@FreeText",System.Data.SqlDbType.VarChar){SqlValue = freeText},
                new SqlParameter("@DocBaseName",System.Data.SqlDbType.VarChar){SqlValue=docBaseName},
                new SqlParameter("@return_value",System.Data.SqlDbType.Int){Direction = ParameterDirection.ReturnValue},
                new SqlParameter("@ErrorReturn",System.Data.SqlDbType.VarChar,300){SqlValue = "",Direction = ParameterDirection.InputOutput},
                new SqlParameter("@DocumentReturn",System.Data.SqlDbType.VarChar,100){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput},
                new SqlParameter("@BaseNameReturn",System.Data.SqlDbType.VarChar,100){SqlValue = DBNull.Value, Direction = ParameterDirection.InputOutput},
                new SqlParameter("@ErrorNumber",System.Data.SqlDbType.Int){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput},
           };

                        command.Parameters.AddRange(parameters);
                        connection.Open();
                        command.ExecuteNonQuery();

                        buildDocumentName.ErrorReturn = (string)command.Parameters["@ErrorReturn"].Value;
                        buildDocumentName.DocumentReturn = (string)command.Parameters["@DocumentReturn"].Value;
                        buildDocumentName.BaseNameReturn = (string)command.Parameters["@BaseNameReturn"].Value;
                        //buildDocumentName.ErrorNumber = (int)command.Parameters["@ErrorNumber"].Value;
                        buildDocumentName.ErrorNumber = command.Parameters["@ErrorNumber"].Value == DBNull.Value ? 0 : Convert.ToInt32(command.Parameters["@ErrorNumber"].Value);

                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            return buildDocumentName;
        }

        public static BuildDocumentNameVersionDetail BuildDocumentNameVersion(int docID, string variantText)
        {
            BuildDocumentNameVersionDetail buildDocumentNameVersion = new BuildDocumentNameVersionDetail();

            string stroreProcedure = "SP_BuildDocumentNameVersion";
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand(stroreProcedure, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        SqlParameter[] parameters = new SqlParameter[]
            {
                new SqlParameter("@DocID",SqlDbType.Int){SqlValue = docID },
                new SqlParameter("@VariantText",SqlDbType.VarChar){SqlValue = variantText},
                new SqlParameter("@ErrorReturn",SqlDbType.VarChar,300){SqlValue = "FAILED!",Direction = ParameterDirection.InputOutput},
                new SqlParameter("@ErrorNumber",SqlDbType.Int){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput},
                new SqlParameter("@DocumentReturn",SqlDbType.VarChar,100){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput},
                new SqlParameter("@BaseNameReturn",SqlDbType.VarChar,100){SqlValue = DBNull.Value, Direction = ParameterDirection.InputOutput},
                new SqlParameter("@VersionMajor",SqlDbType.Int){SqlValue = DBNull.Value, Direction = ParameterDirection.InputOutput},
                new SqlParameter("@return_value",SqlDbType.Int){Direction = ParameterDirection.ReturnValue}
            };
                        command.Parameters.AddRange(parameters);
                        connection.Open();
                        command.ExecuteNonQuery();

                        buildDocumentNameVersion.ErrorReturn = (string)command.Parameters["@ErrorReturn"].Value;
                        buildDocumentNameVersion.ErrorNumber = (int)command.Parameters["@ErrorNumber"].Value;
                        buildDocumentNameVersion.DocumentReturn = (string)command.Parameters["@DocumentReturn"].Value;
                        buildDocumentNameVersion.BaseNameReturn = (string)command.Parameters["@BaseNameReturn"].Value;
                        buildDocumentNameVersion.VersionMajor = (int)command.Parameters["@VersionMajor"].Value;

                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            return buildDocumentNameVersion;
        }

        public static FileCreateDetail File_Create(int opUnitID, long docTypeID, string documentTitle, long workCenterID, long materialID, string freeText, string docBaseName, long superDocID)
        {
            FileCreateDetail fileCreate = new FileCreateDetail();
            string stroreProcedure = "SP_CreateDocument";
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand(stroreProcedure, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        SqlParameter[] parameters = new SqlParameter[]
            {

                new SqlParameter("@OpUnitID",SqlDbType.Int){SqlValue =opUnitID },
                new SqlParameter("@DocTypeID",System.Data.SqlDbType.Int){SqlValue = docTypeID},
                new SqlParameter("@WorkCenterID",System.Data.SqlDbType.Int){ Value = (workCenterID == 0)?DBNull.Value:(object)workCenterID},
                new SqlParameter("@DocumentTitle",System.Data.SqlDbType.VarChar, 100){SqlValue = documentTitle},
                new SqlParameter("@MaterialID",System.Data.SqlDbType.Int){ Value = (materialID == 0)?DBNull.Value:(object)materialID},
                new SqlParameter("@FreeText",System.Data.SqlDbType.VarChar){SqlValue = freeText},
                new SqlParameter("@DocBaseName",System.Data.SqlDbType.VarChar){SqlValue = docBaseName},
                new SqlParameter("@VersionMajor",System.Data.SqlDbType.Int){SqlValue =1 },
                new SqlParameter("@VersionMinor",System.Data.SqlDbType.Int){SqlValue = 0},
                new SqlParameter("@SupercedesDocID",System.Data.SqlDbType.Int){Value=(superDocID ==0)?  DBNull.Value: (object)superDocID},
                new SqlParameter("@FormatVersion",System.Data.SqlDbType.Int){SqlValue = 8},

                new SqlParameter("@ErrorReturn",System.Data.SqlDbType.VarChar,100){SqlValue = "",Direction = ParameterDirection.InputOutput},
                new SqlParameter("@ErrorNumber",System.Data.SqlDbType.Int){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput},
                new SqlParameter("@DocumentNameReturn",System.Data.SqlDbType.VarChar,100){SqlValue = "", Direction = ParameterDirection.InputOutput},
                new SqlParameter("@DocumentBaseReturn",System.Data.SqlDbType.VarChar,100){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput},
                new SqlParameter("@DocumentIDReturn",System.Data.SqlDbType.Int){SqlValue = 0,Direction = ParameterDirection.InputOutput},
                new SqlParameter("@DocPasswordReadReturn",System.Data.SqlDbType.VarChar,32){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput},
                new SqlParameter("@DocPasswordWriteReturn",System.Data.SqlDbType.VarChar,32){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput},
                new SqlParameter("@DocPasswordProtectReturn",System.Data.SqlDbType.VarChar,32){SqlValue = DBNull.Value,Direction = ParameterDirection.InputOutput}
                //new SqlParameter("@VariantText",System.Data.SqlDbType.VarChar){SqlValue = DBNull.Value}
            };
                        command.Parameters.AddRange(parameters);
                        connection.Open();
                        command.ExecuteNonQuery();

                        fileCreate.ErrorReturn = (string)command.Parameters["@ErrorReturn"].Value;
                        fileCreate.ErrorNumber = (int)command.Parameters["@ErrorNumber"].Value;
                        fileCreate.DocumentNameReturn = (string)command.Parameters["@DocumentNameReturn"].Value;
                        fileCreate.DocumentBaseReturn = (string)command.Parameters["@DocumentBaseReturn"].Value;
                        fileCreate.DocumentIDReturn = (int)command.Parameters["@DocumentIDReturn"].Value;
                        fileCreate.DocPasswordReadReturn = (string)command.Parameters["@DocPasswordReadReturn"].Value;
                        fileCreate.DocPasswordWriteReturn = (string)command.Parameters["@DocPasswordWriteReturn"].Value;
                        fileCreate.DocPasswordProtectReturn = (string)command.Parameters["@DocPasswordProtectReturn"].Value;

                    }
                }
            }
            catch (Exception e)
            {
                throw e;  // NewCode1
            }

            return fileCreate;
        }

        public static FileCopyVersDetail File_CopyVers(long oldDOCId, string variantText)
        {
            FileCopyVersDetail fileCopyVers = new FileCopyVersDetail();
            string stroreProcedure = "SP_CopyDocumentVersion";
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand(stroreProcedure, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        SqlParameter[] parameters = new SqlParameter[]
            {

                new SqlParameter("@DocID_OLD",SqlDbType.Int){SqlValue =oldDOCId },
                new SqlParameter("@VersionMajor",System.Data.SqlDbType.Int){SqlValue =1 },
                new SqlParameter("@VariantText",System.Data.SqlDbType.VarChar){SqlValue = variantText},

                new SqlParameter("@DocumentReturn",System.Data.SqlDbType.VarChar,100){SqlValue = "NOTSET", Direction = ParameterDirection.InputOutput},
                new SqlParameter("@ErrorReturn",System.Data.SqlDbType.VarChar,100){SqlValue = "NOVALUE",Direction = ParameterDirection.InputOutput},
            };
                        command.Parameters.AddRange(parameters);
                        connection.Open();
                        command.ExecuteNonQuery();

                        fileCopyVers.DocumentReturn = (string)command.Parameters["@DocumentReturn"].Value;
                        fileCopyVers.ErrorReturn = (string)command.Parameters["@ErrorReturn"].Value;

                    }
                }
            }
            catch (Exception e)
            {
                throw e;  // NewCode1
            }
            return fileCopyVers;
        }

        public static FileCopyDiff_Detail File_CopyDiff(long oldDOCId, int opUnitID, int docTypeID, int materialID, string docBaseName)
        {
            FileCopyDiff_Detail fileCopyDiff = new FileCopyDiff_Detail();
            string stroreProcedure = "SP_CopyDocument";
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand(stroreProcedure, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        SqlParameter[] parameters = new SqlParameter[]
            {

                new SqlParameter("@DocID_OLD",SqlDbType.Int){SqlValue =oldDOCId },
                new SqlParameter("@VersionMajor",System.Data.SqlDbType.Int){SqlValue =1 },
                new SqlParameter("@VariantText",System.Data.SqlDbType.VarChar){SqlValue = ""},
                new SqlParameter("@OpUnitID",SqlDbType.Int){SqlValue =opUnitID },
                new SqlParameter("@DocTypeID",System.Data.SqlDbType.Int){SqlValue = docTypeID},
                new SqlParameter("@MaterialID",System.Data.SqlDbType.Int){ Value = (materialID == 0)?DBNull.Value:(object)materialID},
                new SqlParameter("@DocBaseName",System.Data.SqlDbType.VarChar){SqlValue=docBaseName },

                new SqlParameter("@DocumentReturn",System.Data.SqlDbType.VarChar,100){SqlValue = "NOTSET", Direction = ParameterDirection.InputOutput},
                new SqlParameter("@ErrorReturn",System.Data.SqlDbType.VarChar,100){SqlValue = "NOVALUE",Direction = ParameterDirection.InputOutput},
            };

                        command.Parameters.AddRange(parameters);
                        connection.Open();
                        command.ExecuteNonQuery();

                        fileCopyDiff.DocumentReturn = (string)command.Parameters["@DocumentReturn"].Value;
                        fileCopyDiff.ErrorReturn = (string)command.Parameters["@ErrorReturn"].Value;

                    }
                }
            }
            catch (Exception e)
            {
                throw e;  // NewCode1
            }
            return fileCopyDiff;
        }
        #endregion

        #region FileOpen

        public static int Document_UpdateTitle(int docID, string docTitle)
        {
            const string MyProcName = "Document_UpdateTitle";
            int rowAffected;
            dbConnectionDetails.DatabaseRequestsIn++;
            Connection_Assure();
            try
            {
                //string query = "UPDATE tblDocument SET DocumentTitle = @docTitle WHERE DocID = @docId";
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    SqlCommand command = new SqlCommand("SP_UpdateDocumentTitle", connection);
                    connection.Open();
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@docTitle", docTitle);
                    command.Parameters.AddWithValue("@docId", docID);

                    rowAffected = (int)command.ExecuteScalar();
                    connection.Close();

                }

            }
            catch (Exception e)
            {
                throw e;
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return rowAffected;
        }

        public static List<DocumentDetails> Document_SelectByOpUnit(int opUnitID, int docTypeID, bool includeLocked, bool includeAuthoring, bool includeApproved, bool includeBlocked, bool includeArchived)
        {
            List<DocumentDetails> documentList = new List<DocumentDetails>();
            //string query = "SELECT   DocID,DocumentName,DocLocation, DocPasswordRead, DocPasswordWrite, DocumentPath,  DocumentTitle FROM tblDocument D, tblOpUnitDocType OUDT WHERE D.OpUnitID = @opUnitID  AND D.DocTypeID = @docTypeID AND OUDT.OpUnitID = @opUnitID AND OUDT.DocTypeID = @docTypeID AND DocLocation in (987";// and DocLocation in (-1, 0) order by DocumentName";
            //if (includeLocked) { query += ",-1"; }
            //if (includeAuthoring) { query += ",0"; }
            //if (includeApproved) { query += ",1"; }
            //if (includeArchived) { query += ",2"; }
            //if (includeBlocked) { query += ",3"; }
            //query += ") ORDER BY DocumentName";

            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {

                using (SqlCommand command = new SqlCommand("sp_DocumentByOpUnit", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@OpUnitID", opUnitID);
                    command.Parameters.AddWithValue("@DocTypeID", docTypeID);
                    command.Parameters.AddWithValue("@includeLocked", includeLocked);
                    command.Parameters.AddWithValue("@includeAuthoring", includeAuthoring);
                    command.Parameters.AddWithValue("@includeApproved", includeApproved);
                    command.Parameters.AddWithValue("@includeArchived", includeArchived);
                    command.Parameters.AddWithValue("@includeBlocked", includeBlocked);
                    try
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {

                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    DocumentDetails document = new DocumentDetails
                                    {

                                        DocID = Convert.ToInt32(reader["DocID"]),
                                        DocumentName = reader["DocumentName"].ToString(),
                                        DocLocation = Convert.ToInt32(reader["DocLocation"]),
                                        DocPasswordRead = reader["DocPasswordRead"].ToString(),
                                        DocPasswordWrite = reader["DocPasswordWrite"].ToString(),
                                        DocumentPath = reader["DocumentPath"].ToString(),
                                        DocumentTitle = reader["DocumentTitle"].ToString()
                                    };

                                    documentList.Add(document);

                                }
                            }
                            connection.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }

            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return documentList;
        }

        public static string AppStatus_Select(string keyName)
        {
            const string MyProcName = "AppStatus_Select";
            dbConnectionDetails.DatabaseRequestsIn++;
            Connection_Assure();
            string appStatus = string.Empty;
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("SP_AppStatus", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@AppName", keyName);
                    try
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    appStatus = reader["AppStatus"].ToString();
                                }
                            }
                            else
                            {
                                appStatus = "UNKNOWN";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }


            }
            dbConnectionDetails.DatabaseRequestsOut++;
            dbConnectionDetails.ConnectEndTime = DateTime.Now;
            return appStatus;
        }

        public static DocImageLog DocImageLog_SelectLast(int docID)
        {
            DocImageLog docImageLog = null;
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection con = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand cmd = new SqlCommand("SP_DocImageLogLast", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@DocID", docID);
                    try
                    {
                        con.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {

                                    //var props = typeof(DocImageLog).GetProperties();
                                    //foreach (var prop in props)
                                    //{
                                    //    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                    //    {
                                    //        prop.SetValue(imageLog, reader[prop.Name]);
                                    //    }
                                    //}

                                    docImageLog = new DocImageLog
                                    {
                                        LogDate = reader.GetDateTime(reader.GetOrdinal("LogDate")),
                                        LogNTUser = reader["LogNTUser"].ToString(),
                                        LogHost = reader["LogHost"].ToString(),
                                        LogProgram = reader["LogProgram"].ToString(),
                                        DocBookmarkCount = Convert.ToInt32(reader["DocBookmarkCount"]),
                                        DocCharactersCount = Convert.ToInt32(reader["DocCharactersCount"]),
                                        DocFieldsCount = Convert.ToInt32(reader["DocFieldsCount"]),
                                        DocParagraphsCount = Convert.ToInt32(reader["DocParagraphsCount"]),
                                        DocCustomPropertiesCount = Convert.ToInt32(reader["DocCustomPropertiesCount"]),
                                        DocEndPageNumber = Convert.ToInt32(reader["DocEndPageNumber"]),
                                        DocSectionsCount = Convert.ToInt32(reader["DocSectionsCount"]),
                                        DocSentencesCount = Convert.ToInt32(reader["DocSentencesCount"]),
                                        DocTablesCount = Convert.ToInt32(reader["DocTablesCount"]),
                                        DocWordsCount = Convert.ToInt32(reader["DocWordsCount"])

                                    };
                                }
                            }
                            con.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return docImageLog;
        }

        public static Document_Details Document_Details(int docID)
        {
            Document_Details docDetails = null;
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection con = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand cmd = new SqlCommand("SP_DocumentDetails", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@docId", docID);
                    try
                    {
                        con.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    //Document_Details docDetails = new Document_Details();

                                    //var props = typeof(Document_Details).GetProperties();
                                    //foreach (var prop in props)
                                    //{
                                    //    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                    //    {
                                    //        prop.SetValue(docDetails, reader[prop.Name]);
                                    //    }
                                    //}


                                    docDetails = new Document_Details()
                                    {
                                        DocumentName = reader["DocumentName"].ToString(),
                                        DocumentTitle = reader["DocumentTitle"].ToString(),
                                        ItemCode = reader["ItemCode"].ToString(),
                                        MaterialName = reader["MaterialName"].ToString(),
                                    };
                                }
                            }
                            con.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return docDetails;
        }

        #endregion

        #region SaveHistory
        public static List<DocumentImageLog> DocImageLog_Select(int DocID)
        {
            List<DocumentImageLog> documentLogList = new List<DocumentImageLog>();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection con = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand cmd = new SqlCommand("SP_DocImageLog", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@DocID", DocID);
                    try
                    {
                        con.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    DocumentImageLog imageLog = new DocumentImageLog();

                                    var props = typeof(DocumentImageLog).GetProperties();
                                    foreach (var prop in props)
                                    {
                                        if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                        {
                                            prop.SetValue(imageLog, reader[prop.Name]);
                                        }
                                    }

                                    documentLogList.Add(imageLog);
                                }
                            }
                            con.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return documentLogList;
        }

        #endregion

        #region VerifyWizard
        //Get Document path for Document Verify wizard
        public static string OpUnitDocType_DocumentPath_Select()
        {
            string docPath = "Unknown";
            dbConnectionDetails.DatabaseRequestsIn++;
            Connection_Assure();
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand("SP_OpUnitDocTypeDocumentPath", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                docPath = reader["DocumentPath"].ToString();
                            }
                        }

                    }
                }


            }
            catch (Exception ex)
            {
                throw ex;
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return docPath;
        }

        public static List<DocumentValidateDetails> DocumentValidate_Select(int docID)
        {
            //DataTable dt = new DataTable();
            List<DocumentValidateDetails> documentValidateList = new List<DocumentValidateDetails>();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            string stroreProcedure = "usp_ModuleValidate";
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand(stroreProcedure, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@DocID", docID);

                        connection.Open();
                        //using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        //{
                        //    adapter.Fill(dt);
                        //}
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    DocumentValidateDetails documentValidateDetails = new DocumentValidateDetails();

                                    var props = typeof(DocumentValidateDetails).GetProperties();
                                    foreach (var prop in props)
                                    {
                                        if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                        {
                                            prop.SetValue(documentValidateDetails, reader[prop.Name]);
                                        }
                                    }

                                    documentValidateList.Add(documentValidateDetails);
                                }
                            }
                            connection.Close();
                        }
                    }

                }
                //if (dt.Columns.Count > 0)
                //{
                //    dt.Columns.RemoveAt(0);
                //}

            }
            catch (Exception ex)
            {
                throw ex;
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return documentValidateList;
        }

        //Document Verify Wizard - Verify property:
        public static string VerifyProperty(int pDocId)
        {
            string resultMessage = string.Empty;
            string storedProcedur = "usp_DocReqPropCheck";

            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand(storedProcedur, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@DocID", pDocId);

                        //Output parameter for return message
                        SqlParameter returnMessageParam = new SqlParameter("@ReturnMessage", SqlDbType.NVarChar, 255);
                        returnMessageParam.Direction = ParameterDirection.Output;
                        command.Parameters.Add(returnMessageParam);

                        connection.Open();
                        command.ExecuteNonQuery();

                        string returnPropertyMessage = returnMessageParam.Value as string;

                        if (!string.IsNullOrEmpty(returnPropertyMessage))
                        {
                            resultMessage = "ERROR in Property-Verification: " + returnPropertyMessage;
                        }
                    }
                }

            }
            catch (Exception e)
            {
                throw new Exception("Error (" + e.HResult + "): " + e.Message, e);
            }

            return resultMessage;

        }

        //Document Verify wizard --> lock the document
        public static int SetDocumentToLockedState(int DocId)
        {
            const string MyProcName = "SetDocumentToLockedState";
            int rowAffected;
            Connection_Assure();
            dbConnectionDetails.DatabaseRequestsIn++;
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    SqlCommand command = new SqlCommand("SP_SetDocumentToLockedState", connection);

                    connection.Open();
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@DocID", DocId);
                    rowAffected = (int)command.ExecuteScalar();

                    connection.Close();

                }

            }
            catch (Exception e)
            {
                throw e;
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return rowAffected;
        }

        #endregion
        public static bool Read(string documentName, DocumentInfo documentInfo)
        {
            const string MyProcName = "Read";
            const string DOCUMENT_PATH = "DocumentPath";
            const string OP_UNIT_ID = "OpUnitID";
            const string DOC_TRACK_CHANGES = "DocTrackChanges";
            const string OP_UNIT_TYPE = "OpUnitType";
            const string OP_UNIT = "OpUnit";
            const string DOC_TYPE_ID = "DocTypeID";
            const string cDocType = "DocType";
            const string cDocTypeABB = "DocTypeAbb";

            documentInfo.GetClsDocument = new ClsDocument();
            List<ClsDocument> documents = new List<ClsDocument>();
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {

                    using (SqlCommand command = new SqlCommand("SP_DocumentInfo", connection))
                    {

                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@DocumentName", documentName.Trim());
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    var props = typeof(ClsDocument).GetProperties();
                                    foreach (var prop in props)
                                    {
                                        if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                        {
                                            prop.SetValue(documentInfo.GetClsDocument, reader[prop.Name]);
                                        }
                                    }

                                    documents.Add(documentInfo.GetClsDocument);
                                    if (documentInfo.GetClsDocument.DocTrackChanges == "T")
                                    {
                                        documentInfo.GetClsDocument.mblnTrackChangesReq = true;

                                    }
                                    else
                                    {
                                        documentInfo.GetClsDocument.mblnTrackChangesReq = false;
                                    }
                                }
                                documentInfo.GetClsDocument.mintOpUnitID = documents[0].OpUnitID;

                                documentInfo.GetTblOpUnit = OpUnit_SelectONE(documentInfo.GetClsDocument.mintOpUnitID);
                                documentInfo.GetClsDocument.mstrOpUnitType = documentInfo.GetTblOpUnit.OpUnitType;
                                documentInfo.GetClsDocument.mstrOpUnit = documentInfo.GetTblOpUnit.OpUnit;

                                documentInfo.GetClsDocument.mintDocTypeID = documents[0].DocTypeID;

                                documentInfo.GetDocumentType = DocType_Select(documentInfo.GetClsDocument.mintDocTypeID);
                                documentInfo.GetClsDocument.mstrDocTypeAbb = documentInfo.GetDocumentType.DocTypeAbb;
                                documentInfo.GetClsDocument.mstrDocType = documentInfo.GetDocumentType.DocType;

                                if (documents[0].MaterialID == null)
                                {
                                    documentInfo.GetClsDocument.mlngMaterialID = -1;
                                    documentInfo.GetClsDocument.mstrMaterialName = "";
                                    documentInfo.GetClsDocument.mstrMaterialEDPNUM = "";
                                    documentInfo.GetClsDocument.mManufacturing_Product_Code = "";
                                    documentInfo.GetClsDocument.mNDC_Code = "";
                                    documentInfo.GetClsDocument.mProduct_Identification_Code = "";
                                    documentInfo.GetClsDocument.mRCS_Sales_Code = "";
                                    documentInfo.GetClsDocument.mGlobal_Trade_ID_Number = "";
                                    //Added for GTIN - PANNEJ02
                                    documentInfo.GetClsDocument.mVialLabelGlobalTradeIDNumber = "";
                                }
                                else
                                {
                                    documentInfo.GetClsDocument.mlngMaterialID = documents[0].MaterialID;

                                    documentInfo.GetProductDetail = Product_Select(documentInfo.GetClsDocument.mlngMaterialID);
                                    if(documentInfo.GetProductDetail == null)  // NewCode1
                                    {
                                        var ex = new Exception();
                                        ex.Data["errNum"] = "3021";
                                        ex.Data["errDesc"] = "The current record has been deleted. Requested operation requires a current record.";
                                        CommonService.Error_Handler(ex.Data["errNum"].ToString(), ex.Data["errDesc"].ToString(), MyCompName, MyProcName);
                                        return false;
                                    }
                                    documentInfo.GetClsDocument.mstrMaterialName = documentInfo.GetProductDetail.MaterialTitle;
                                    documentInfo.GetClsDocument.mstrMaterialEDPNUM = documentInfo.GetProductDetail.ItemCode;
                                    documentInfo.GetClsDocument.mManufacturing_Product_Code = documentInfo.GetProductDetail.Manufacturing_Product_Code;
                                    documentInfo.GetClsDocument.mNDC_Code = documentInfo.GetProductDetail.NDC_Code;
                                    documentInfo.GetClsDocument.mProduct_Identification_Code = documentInfo.GetProductDetail.Product_Identification_Code;
                                    documentInfo.GetClsDocument.mRCS_Sales_Code = documentInfo.GetProductDetail.RCS_Sales_Code;
                                    documentInfo.GetClsDocument.mGlobal_Trade_ID_Number = documentInfo.GetProductDetail.Global_Trade_ID_Number;
                                    //Added for GTIN - PANNEJ02 -- its taking material level details 
                                    documentInfo.GetClsDocument.mVialLabelGlobalTradeIDNumber = documentInfo.GetProductDetail.VialLabelGlobalTradeIDNumber;

                                }
                                documentInfo.GetOpUnitDocType = OpUnitDocType_SelectOne(documentInfo.GetClsDocument.mintOpUnitID, documentInfo.GetClsDocument.mintDocTypeID);

                                documentInfo.GetClsDocument.mDocIsModule = documentInfo.GetOpUnitDocType.DocIsModule;
                                documentInfo.GetClsDocument.mProjectNameUsed = documentInfo.GetOpUnitDocType.ProjectNameUsed;
                                documentInfo.GetClsDocument.mDocPermitsQuantitySufficient = documentInfo.GetOpUnitDocType.DocPermitsQuantitySufficient;
                                documentInfo.GetClsDocument.mDocHasMaterial = documentInfo.GetOpUnitDocType.DocHasMaterial;
                                documentInfo.GetClsDocument.mDocHasModule = documentInfo.GetOpUnitDocType.DocHasModule;
                                documentInfo.GetClsDocument.mDocHasProduct = documentInfo.GetOpUnitDocType.DocHasProduct;
                                documentInfo.GetClsDocument.mDocHasProperty = documentInfo.GetOpUnitDocType.DocHasProperty;


                            }
                            documentInfo.GetClsDocument.mblnUsable = true;
                        }
                    }

                }


                return true;
            }

            catch (Exception ex)
            {
                CommonService.Error_Handler(ex.HResult.ToString(), ex.Message, MyCompName, MyProcName);
                // clsdocument.docThis = null;
                return false;
            }
        }

        public static TblOpUnit OpUnit_SelectONE(int OpUnitID)
        {
            TblOpUnit tblOpUnits = new TblOpUnit();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand command = new SqlCommand("SP_OpUnitSelectONE", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@OpUnitID", OpUnitID);
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {

                            while (reader.Read())
                            {

                                var props = typeof(TblOpUnit).GetProperties();
                                foreach (var prop in props)
                                {
                                    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                    {
                                        prop.SetValue(tblOpUnits, reader[prop.Name]);
                                    }
                                }

                            }

                        }
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return tblOpUnits;
        }

        public static DocumentType DocType_Select(int DocTypeID)
        {
            DocumentType documentType = new DocumentType();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand command = new SqlCommand("SP_DocTypeSelected", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@DocTypeID", DocTypeID);
                    connection.Open();

                    try
                    {

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {

                                var props = typeof(DocumentType).GetProperties();
                                foreach (var prop in props)
                                {
                                    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                    {
                                        prop.SetValue(documentType, reader[prop.Name]);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return documentType;

        }

        public static ProductDetail Product_Select(int materialID)
        {
            int lngMaterialID = 0;
            string strMaterialSource;
            ProductDetail productDetail = new ProductDetail(); // need to check
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            if (materialID == lngMaterialID)
            {
                return productDetail;
            }

            lngMaterialID = materialID;

            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {   //upadte SP name for GTIN changes
                using (SqlCommand command = new SqlCommand("SP_ProductSelected", connection))
                //using (SqlCommand command = new SqlCommand("SP_ProductSelected_NEW", connection))
                {
                    connection.Open();

                    try
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@MaterialID", materialID);
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows) // NewCode1
                            {
                            while (reader.Read())
                            {

                                var props = typeof(ProductDetail).GetProperties();
                                foreach (var prop in props)
                                {
                                    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                    {
                                        prop.SetValue(productDetail, reader[prop.Name]);
                                    }

                                }
                                }

                            }
                            else
                            {
                                productDetail = null;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                }
                dbConnectionDetails.DatabaseRequestsOut++;
                return productDetail;

            }
        }

        public static OpUnitDocType OpUnitDocType_SelectOne(int OpUnitID, int DocTypeID)

        {
            int intOpUnitID = 0;
            int intDocTypeID = 0;
            OpUnitDocType opUnitDoctype = new OpUnitDocType();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            if (OpUnitID == intOpUnitID && DocTypeID == intDocTypeID)
            {
                return opUnitDoctype;
            }

            intOpUnitID = OpUnitID;
            intDocTypeID = DocTypeID;

            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {

                using (SqlCommand command = new SqlCommand("SP_OpUnitDocTypeSelectOne", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@DocTypeID", DocTypeID);
                    command.Parameters.AddWithValue("@OpUnitID", OpUnitID);

                    try
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {

                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    var props = typeof(OpUnitDocType).GetProperties();
                                    foreach (var prop in props)
                                    {
                                        if (Attribute.IsDefined(prop, typeof(NonDBBackedAttribute)))
                                        {
                                            continue;
                                        }

                                        if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                        {
                                            prop.SetValue(opUnitDoctype, reader[prop.Name]);
                                        }
                                    }

                                }
                            }
                            connection.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }

            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return opUnitDoctype;
        }

        #region EditMaterials
        public static List<MaterialCategoryDetails> MaterialCategory_Select(int opUnitID)
        {
            List<MaterialCategoryDetails> materials = new List<MaterialCategoryDetails>();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand command = new SqlCommand("SP_MaterialCategorySelected", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    connection.Open();

                    try
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    MaterialCategoryDetails materialCategoryDetails = new MaterialCategoryDetails();

                                    var props = typeof(MaterialCategoryDetails).GetProperties();
                                    foreach (var prop in props)
                                    {
                                        if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                        {
                                            prop.SetValue(materialCategoryDetails, reader[prop.Name]);
                                        }

                                    }
                                    materials.Add(materialCategoryDetails);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    return materials;
                }
            }
        }

        public static List<OpUnitMaterial> OpUnitMaterial_SelectALL(int opUnitID, string MaterialCategory)
        {
            List<OpUnitMaterial> opUnitmaterials = new List<OpUnitMaterial>();

            //string opUnitMaterialQuery = " Select top 32767 MaterialID,MaterialTitle,ItemCode_MaterialTitle,MaterialTitle_ItemCode,ISNULL(ItemCode,'') AS ItemCode,MaterialCategory,MICK_Material_Item_Code FROM vewMaterial WHERE OpUnitID =@OpUnitID ";
            //if (MaterialCategory != "")
            //{
            //    opUnitMaterialQuery = opUnitMaterialQuery + " AND MaterialCategory=" + "'" + MaterialCategory + "'";
            //}
            //else
            //{
            //    opUnitMaterialQuery = opUnitMaterialQuery + " order by materialcategory";
            //}

            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand command = new SqlCommand("SP_OpUnitMaterialSelectALL", connection))
                {
                    connection.Open();
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@OpUnitID", opUnitID);
                    command.Parameters.AddWithValue("@MaterialCategory", MaterialCategory);
                    try
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    // OpUnitMaterial opUnitMaterial = new OpUnitMaterial();

                                    //var props = typeof(OpUnitMaterial).GetProperties();
                                    //foreach (var prop in props)
                                    //{
                                    //    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                    //    {
                                    //        prop.SetValue(opUnitMaterial, reader[prop.Name]);
                                    //    }

                                    //}

                                    OpUnitMaterial opUnitMaterial = new OpUnitMaterial
                                    {
                                        MaterialID = Convert.ToInt32(reader["MaterialID"]),
                                        MaterialTitle = reader["MaterialTitle"].ToString(),
                                        ItemCode_MaterialTitle = reader["ItemCode_MaterialTitle"].ToString(),
                                        MaterialTitle_ItemCode = reader["MaterialTitle_ItemCode"].ToString(),
                                        ItemCode = reader["ItemCode"].ToString(),
                                        MaterialCategory = reader["MaterialCategory"].ToString(),
                                        MICK_Material_Item_Code = reader["MICK_Material_Item_Code"].ToString()

                                    };
                                    opUnitmaterials.Add(opUnitMaterial);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    dbConnectionDetails.DatabaseRequestsOut++;
                    return opUnitmaterials;
                }
            }


        }

        public static List<MaterialUnit> MaterialUnit_Select()
        {
            List<MaterialUnit> materialUnits = new List<MaterialUnit>();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            try
            {

                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand("SP_MaterialUnit", connection))
                    {
                        connection.Open();
                        command.CommandType = CommandType.StoredProcedure;
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    materialUnits.Add(new MaterialUnit
                                    {

                                        MaterialUOM = reader["MaterialUOM"].ToString()
                                    });
                                }
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return materialUnits;
        }

        public static List<MaterialUsageDetails> MaterialUsage_Select()
        {
            List<MaterialUsageDetails> materialUsageDetails = new List<MaterialUsageDetails>();
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            try
            {

                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand("SP_MaterialUsageSelect", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    MaterialUsageDetails usage = new MaterialUsageDetails();

                                    var props = typeof(MaterialUsageDetails).GetProperties();
                                    foreach (var prop in props)
                                    {
                                        if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                        {
                                            prop.SetValue(usage, reader[prop.Name]);
                                        }

                                    }
                                    materialUsageDetails.Add(usage);

                                }
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return materialUsageDetails;
        }


        public static string DocMaterial_UsedByProperty(int docId, int materailId)
        {
            string docMaterial_UsedByProperty = "";
            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand command = new SqlCommand("SP_DocMaterialUsedByProperty", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    connection.Open();
                    command.Parameters.AddWithValue("@DocID", docId);
                    command.Parameters.AddWithValue("@MaterialId", materailId);

                    try
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)

                            {
                                while (reader.Read())
                                {
                                    docMaterial_UsedByProperty = reader["PropertyInstance"].ToString();
                                }

                            }

                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return docMaterial_UsedByProperty;
        }
        public static void DocMaterial_InsertUpdate(int mDocID, int materialDBID, string materialAmount, string materialUOM, int materialSequence, int docMaterialQuantitySufficient, string materialUsage, string materialImprint, string materialPerUnitData, string materialBOMText)
        {
            string MaterialAmountScale = "";
            if (string.IsNullOrEmpty(materialUOM))
            {
                materialUOM = "'KG'";

            }
            if (string.IsNullOrEmpty(materialAmount))
            {
                materialAmount = "0";
                MaterialAmountScale = "0";
            }
            else
            {
                MaterialAmountScale = GetMaterialAmountScale(materialAmount).ToString();
            }
            Connection_Assure();

            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                using (SqlCommand command = new SqlCommand("usp_InsertUpdateDocMaterial", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@DocId", mDocID);
                    command.Parameters.AddWithValue("@BOMIndex", 1);
                    command.Parameters.AddWithValue("@MaterialID", materialDBID);
                    command.Parameters.AddWithValue("@MaterialSequence", materialSequence);
                    command.Parameters.AddWithValue("@MaterialAmount", materialAmount);
                    command.Parameters.AddWithValue("@MaterialAmountScale", MaterialAmountScale);
                    command.Parameters.AddWithValue("@MaterialUOM", materialUOM);
                    command.Parameters.AddWithValue("@QuantitySufficient", docMaterialQuantitySufficient);
                    command.Parameters.AddWithValue("@MaterialUsage", materialUsage);
                    command.Parameters.AddWithValue("@MaterialImprint", materialImprint);
                    command.Parameters.AddWithValue("@MaterialPerUnitData", materialPerUnitData);
                    command.Parameters.AddWithValue("@MaterialBOMText", materialBOMText);
                    try
                    {
                        connection.Open();
                        command.ExecuteNonQuery();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + Environment.NewLine + Environment.NewLine + command.CommandText);
                    }
                }
            }
        }

        public static void DocMaterial_Delete(int mDocID, int sequence)
        {
            const string MyProcName = "DocMaterial_Delete";
            dbConnectionDetails.DatabaseRequestsIn++;
            Connection_Assure();

            if (string.IsNullOrEmpty(mstrDBTransactionName))
            {
                MessageBox.Show("Active Transaction is Required", MyProcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            DocMaterial_Reset();
            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                try
                {

                    using (SqlCommand command = new SqlCommand("SP_DocMaterialDelete", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        connection.Open();
                        command.Parameters.AddWithValue("@DocID", mDocID);
                        command.Parameters.AddWithValue("@Sequence", sequence);

                        command.ExecuteNonQuery();
                    }

                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
        }

        private static void DocMaterial_Reset()
        {
            DocMaterial_Select(DocID_ResetRequest, true);

        }

        public static List<DocMaterial_SelectDetails> DocMaterial_Select(int docID, bool IncludeZeroAmount)
        {
            int sDocID_IncludeZero = 0;
            int sDocID_ExcludeZero = 0;
            string strDocMaterialSource;
            List<DocMaterial_SelectDetails> DocMaterial_IncludeZero = null;
            List<DocMaterial_SelectDetails> DocMaterial_ExcludeZero = null;

            dbConnectionDetails.DatabaseRequestsIn++;
            if (docID == DocID_ResetRequest)
            {
                sDocID_IncludeZero = docID;
                sDocID_ExcludeZero = docID;
                return null;
            }
            if (IncludeZeroAmount)
            {
                if (docID > 0 && docID == sDocID_IncludeZero)
                {
                    goto ReturnRecordSet_IncludeZero;
                }
                else
                {
                    RecordSet_Assure();
                    DocMaterial_IncludeZero = new List<DocMaterial_SelectDetails>();
                }
            }
            else
            {
                if (docID > 0 && docID == sDocID_ExcludeZero)
                {
                    goto ReturnRecordSet_ExcludeZero;
                }
                else
                {
                    RecordSet_Assure();
                    DocMaterial_ExcludeZero = new List<DocMaterial_SelectDetails>();
                }

            }

            //strDocMaterialSource = "Select M.MaterialTitle, M.MaterialID ,ItemCode_MaterialTitle AS MaterialDisplay ,IsNull (ItemCode,'') as ItemCode, IsNull (MaterialAmountDisplay, '') as MaterialAmountDisplay,IsNull (MaterialUOM, '') as MaterialUOM ,IsNull (MaterialSequence,'0') as MaterialSequence , IsNull (MaterialUsage, '') as MaterialUsage,QuantitySufficient, MaterialCategory, MaterialImprint, MaterialPerUnitData, MaterialBOMText FROM tblDocMaterial DM  , vewMaterial M  WHERE DM.MaterialID = M.MaterialID  and DM.DocID=" + docID;

            //if (!IncludeZeroAmount)
            //{
            //    strDocMaterialSource = strDocMaterialSource + " AND DM.MaterialAmount <> 0 ";
            //}

            //strDocMaterialSource = strDocMaterialSource + " Order By DM.MaterialSequence ";


            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {

                using (SqlCommand command = new SqlCommand("SP_DocMaterialSelect", connection))
                {
                    connection.Open();
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@DocId", docID);
                    command.Parameters.AddWithValue("@IncludeZeroAmount", IncludeZeroAmount);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (IncludeZeroAmount)
                        {

                            var props = typeof(DocMaterial_SelectDetails).GetProperties();

                            while (reader.Read())
                            {
                                DocMaterial_SelectDetails docMaterial = new DocMaterial_SelectDetails()
                                {
                                    MaterialTitle = Convert.ToString(reader["MaterialTitle"]),
                                    MaterialID = Convert.ToInt32(reader["MaterialID"]),
                                    MaterialDisplay = Convert.ToString(reader["MaterialDisplay"]),
                                    ItemCode = Convert.ToString(reader["ItemCode"]),
                                    MaterialAmountDisplay = Convert.ToString(reader["MaterialAmountDisplay"]),
                                    MaterialUOM = Convert.ToString(reader["MaterialUOM"]),
                                    MaterialSequence = Convert.ToInt32(reader["MaterialSequence"]),
                                    MaterialUsage = Convert.ToString(reader["MaterialUsage"]),
                                    QuantitySufficient = Convert.ToBoolean(reader["QuantitySufficient"]),
                                    MaterialCategory = Convert.ToString(reader["MaterialCategory"]),
                                    MaterialImprint = Convert.ToString(reader["MaterialImprint"]),
                                    MaterialPerUnitData = Convert.ToString(reader["MaterialPerUnitData"]),
                                    MaterialBOMText = Convert.ToString(reader["MaterialBOMText"])
                                };

                                //foreach (var prop in props)
                                //{
                                //    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                //    {
                                //        prop.SetValue(docMaterial, reader[prop.Name]);
                                //    }

                                //}
                                DocMaterial_IncludeZero.Add(docMaterial);

                            }
                            sDocID_IncludeZero = docID;
                            goto ReturnRecordSet_IncludeZero;
                        }
                        else
                        {
                            while (reader.Read())
                            {
                                DocMaterial_SelectDetails docMaterial = new DocMaterial_SelectDetails()
                                {

                                    MaterialTitle = Convert.ToString(reader["MaterialTitle"]),
                                    MaterialID = Convert.ToInt32(reader["MaterialID"]),
                                    MaterialDisplay = Convert.ToString(reader["MaterialDisplay"]),
                                    ItemCode = Convert.ToString(reader["ItemCode"]),
                                    MaterialAmountDisplay = Convert.ToString(reader["MaterialAmountDisplay"]),
                                    MaterialUOM = Convert.ToString(reader["MaterialUOM"]),
                                    MaterialSequence = Convert.ToInt32(reader["MaterialSequence"]),
                                    QuantitySufficient = Convert.ToBoolean(reader["QuantitySufficient"]),
                                    MaterialCategory = Convert.ToString(reader["MaterialCategory"]),
                                    MaterialImprint = Convert.ToString(reader["MaterialImprint"]),
                                    MaterialPerUnitData = Convert.ToString(reader["MaterialPerUnitData"]),
                                    MaterialBOMText = Convert.ToString(reader["MaterialBOMText"])
                                };
                                //var props = typeof(DocMaterial_SelectDetails).GetProperties();
                                //foreach (var prop in props)
                                //{
                                //    if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                //    {
                                //        prop.SetValue(docMaterial, reader[prop.Name]);
                                //    }

                                //}
                                DocMaterial_ExcludeZero.Add(docMaterial);

                            }
                            sDocID_ExcludeZero = docID;
                            goto ReturnRecordSet_ExcludeZero;
                        }

                    }
                }
            }
        ReturnRecordSet_IncludeZero:
            dbConnectionDetails.DatabaseRequestsOut++;
            return DocMaterial_IncludeZero;
        ReturnRecordSet_ExcludeZero:
            dbConnectionDetails.DatabaseRequestsOut++;
            return DocMaterial_ExcludeZero;
        }



        private static int GetMaterialAmountScale(string materialAmount)
        {
            string strDecimal;
            int intDecimalPos;
            string SearchChar;
            materialAmount = materialAmount.Trim();
            SearchChar = ".";
            intDecimalPos = materialAmount.IndexOf(SearchChar, StringComparison.Ordinal) + 1;
            if (intDecimalPos > 0)
            {
                return materialAmount.Length - intDecimalPos;
            }
            else
            {
                return 0;
            }
        }

        #endregion

        public static List<SelctedProperty> Property_Select(string docIsModule, string propertyCategory)
        {
            List<SelctedProperty> properties = new List<SelctedProperty>();
            string modulePredicate;
            string systemPredicate;
            string propertySource;

            dbConnectionDetails.DatabaseRequestsIn++;
            RecordSet_Assure();
            //if (docIsModule == "Y")
            //{
            //    modulePredicate = " or PropertyName in ('ModuleTitle','ModuleTableRows')  ";
            //}
            //else
            //{
            //    modulePredicate = " and PropertyName not in ('ModuleTitle','ModuleTableRows')  ";
            //}
            //systemPredicate = " or PropertyName in ('ChangeNum','LotNbr','OrderNum','PrintDate','PrintTime','ProductID','ProductName','Manufacturing_Product_Code','NDC_Code','Product_Identification_Code','Global_Trade_ID_Number','RCS_Sales_Code','PrintPages','RecordType','FirstLotUsage','DocumentTitle') or PropertyName LIKE 'MOD\\_%' ESCAPE '\\' or PropertyName LIKE 'MTL\\_%' ESCAPE '\\' ";
            //propertySource = "Select PropertyName,PropertyDescription,PropertyDataType,PropertyDataTypeName,PropertyIndexMax,ISNULL(PropertyUOMLookupType, '') AS PropertyUOMLookupType, PropertyUsage, ISNULL(PropertyLookupType, '<NoLookup>') AS PropertyLookupType, PropertyCategory, PropertyTextMaxLen from tblProperty WHERE Include =1  and (  PropertyUsage in ('A','B') " + modulePredicate + systemPredicate + " ) ";
            //if (propertyCategory != "")
            //{
            //    propertySource = propertySource + " AND PropertyCategory=" + propertyCategory;
            //}
            //propertySource = propertySource + " Order By PropertyName ";

            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                connection.Open();
                //updated SP name GTIN PANNEJ02
                using (SqlCommand command = new SqlCommand("SP_PropertySelect", connection))
                //using (SqlCommand command = new SqlCommand("SP_PropertySelect_NEW", connection))
                {
                    try
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@docIsModule", docIsModule);
                        command.Parameters.AddWithValue("@propertyCategory", propertyCategory);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    SelctedProperty selctedProperty = new SelctedProperty();
                                    var props = typeof(SelctedProperty).GetProperties();
                                    foreach (var prop in props)
                                    {
                                        if (!reader.IsDBNull(reader.GetOrdinal(prop.Name)))
                                        {
                                            prop.SetValue(selctedProperty, reader[prop.Name]);
                                        }
                                    }


                                    properties.Add(selctedProperty);

                                }
                            }
                            connection.Close();

                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            dbConnectionDetails.DatabaseRequestsOut++;
            return properties;
        }

        public static void DocumentSaved_DB_Record(int docID, Document document)
        {
            string stroreProcedure = "SP_SaveDocument";
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand(stroreProcedure, connection))
                    {
                        connection.Open();
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@DocID", docID);
                        command.Parameters.AddWithValue("@DocBookmarkCount", document.Bookmarks.Count);
                        command.Parameters.AddWithValue("@DocCharactersCount", document.Characters.Count);
                        command.Parameters.AddWithValue("@DocFieldsCount", document.Fields.Count);
                        command.Parameters.AddWithValue("@DocCustomPropertiesCount", document.CustomDocumentProperties.count);
                        command.Parameters.AddWithValue("@DocParagraphsCount", document.Paragraphs.Count);
                        command.Parameters.AddWithValue("@DocEndPageNumber", document.Range().Information[WdInformation.wdNumberOfPagesInDocument]);
                        command.Parameters.AddWithValue("@DocSectionsCount", document.Sections.Count);
                        command.Parameters.AddWithValue("@DocSentencesCount", document.Sentences.Count);
                        command.Parameters.AddWithValue("@DocTablesCount", document.Tables.Count);
                        command.Parameters.AddWithValue("@DocWordsCount", document.Words.Count);
                        command.Parameters.AddWithValue("@LogDate", "");
                        command.Parameters.AddWithValue("@ErrorMessage", "");
                        command.ExecuteNonQuery();

                        if (!string.IsNullOrEmpty(command.Parameters["@ErrorMessage"].Value.ToString()))
                        {
                            MessageBox.Show(command.Parameters["@ErrorMessage"].Value.ToString(),
                                            "DocumentSaved_DB_Record Error",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Exclamation);
                            return;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static void NotifyUser_UnIncludedMaterials(int docID)
        {
            // SqlParameter[] parameters = new SqlParameter[]
            //{
            //     new SqlParameter("@DOcID",SqlDbType.Int){SqlValue =docID }
            //};
            // var resultSet = DatabaseHelper.ExecuteStoredProcedure<object>("usp_GetUnincludedMaterial", parameters, null, out var outputValues, false);

            RecordSet_Assure();
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    using (SqlCommand command = new SqlCommand("usp_GetUnincludedMaterial", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@DocID", docID);

                        connection.Open();

                        command.ExecuteNonQuery();
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
