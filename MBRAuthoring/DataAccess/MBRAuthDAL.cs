using MBRAuthoring.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace MBRAuthoring.DataAccess
{
    class MBRAuthDAL
    {
        private SqlConnection SqlConn = null;
        protected List<KeyValuePair<string, string>> list;
        public static DBDetails dbDetails;
        public static DBConnectionDetails dbConnectionDetails;


        public MBRAuthDAL()
        {
            list = new List<KeyValuePair<string, string>>();
        }
        public DataTable GetDocumentProperty(int DocID)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@DocID", DocID);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "usp_GetDocProperty", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }

        public DataTable GetDocumentDetails(int DocID)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@DocID", DocID);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "usp_GetDocumentDetails", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }

        public DataTable GetTablelist()
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "usp_GetTableList", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }
        public DataTable GetPropertyCategory()
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "usp_GetPropertyCategory", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }

        public DataTable GetPropertyAll(string PropertyCategory, string IsModule) //100125
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                AddSQLParameter("@PropertyCategory", PropertyCategory);
                AddSQLParameter("@IsModule", IsModule);
                // Creating a connections
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "usp_GetPropertyAll", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }

        public string GetProperty_ValueRequired(int DocID, string PropertyName, int PropertyIndex, string PropertyValue)

        {

            SqlConnection con = null;
            SqlCommand cmd;
            DataTable dtResultset = new DataTable();
            SqlTransaction trans = null;
            string strPropertyValueRequired = string.Empty;
            string query = string.Empty;
            try
            {
                query = "SELECT dbo.PROPERTY_VALUEREQUIRED('A'," + "'" + DocID + "','" + PropertyName + "','" + PropertyIndex + "','" + PropertyValue + "')";
                using (con = new SqlConnection(AppConfig.ConnString))
                {
                    OpenConnection(ref con, ref trans);
                    cmd = new SqlCommand(query, con);
                    dtResultset.Load(cmd.ExecuteReader());
                    if (dtResultset != null && dtResultset.Rows.Count > 0)
                    {
                        strPropertyValueRequired = dtResultset.Rows[0][0].ToString().Trim();

                    }
                }

                /*ClearList();
                AddSQLParameter("@DocID", DocID);
                AddSQLParameter("@PropertyOf", "A");
                AddSQLParameter("@PropertyName", PropertyName);
                AddSQLParameter("@PropertyIndex", PropertyIndex);
                AddSQLParameter("@PropertyValue", PropertyValue);
                // Creating a connections
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "SP_GetPropertyValueRequired", list);
                if (dtResultset != null && dtResultset.Rows.Count > 0)
                {
                    strPropertyValueRequired = dtResultset.Rows[0][0].ToString().Trim();

                }*/
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return strPropertyValueRequired;
        }


        public string PropertyValidate(ref string ErrorMessage, int DocID, string PropertyOf, string PropertyName, int PropertyIndex, string PropertyValue, string PropertyUOM, int MaterialID, int WorkCenterID)
        {
            SqlConnection con = null;
            string PropertyValueOUT = string.Empty;
            SqlTransaction trans = null;
            con = new SqlConnection(AppConfig.ConnString);
            /*if (MaterialID < 0)
            {
                MaterialID = 0;
            }
            if (WorkCenterID < 0)
            {
                WorkCenterID = 0;
            }*/
            if (PropertyUOM == null)//added for avoiding exception  
            {
                PropertyUOM = "";
            }
            try
            {
                using (SqlCommand cmd = new SqlCommand("SP_PropertyValidate", con))
                {
                    OpenConnection(ref con, ref trans);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@DocID", DocID);
                    cmd.Parameters.AddWithValue("@PropertyOf", PropertyOf);
                    cmd.Parameters.AddWithValue("@PropertyName", PropertyName);
                    cmd.Parameters.AddWithValue("@PropertyIndex", PropertyIndex);
                    cmd.Parameters.AddWithValue("@PropertyValue", PropertyValue);
                    cmd.Parameters.AddWithValue("@PropertyUOM", PropertyUOM);
                    cmd.Parameters.AddWithValue("@MaterialID", MaterialID <= 0 ? (object)DBNull.Value : MaterialID);
                    cmd.Parameters.AddWithValue("@WorkCenterID", WorkCenterID <= 0 ? (object)DBNull.Value : WorkCenterID);

                    SqlParameter errorMessage = new SqlParameter("@ErrorMessage", SqlDbType.VarChar, 250);
                    errorMessage.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(errorMessage);

                    SqlParameter propertyValueOUT = new SqlParameter("@PropertyValueOUT", SqlDbType.VarChar, 7950);
                    propertyValueOUT.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(propertyValueOUT);


                    SqlDataReader reader = cmd.ExecuteReader();

                    ErrorMessage = (string)errorMessage.Value;
                    PropertyValueOUT = (string)propertyValueOUT.Value;

                }
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return PropertyValueOUT;
        }
        public DataTable DocPropertyRequired_SelectAll(int OpUnitID, int DocTypeID)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                AddSQLParameter("@OpUnitID", OpUnitID);
                AddSQLParameter("@DocTypeID", DocTypeID);
                // Creating a connections
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "usp_GetRequiredProperties", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }

        public DataTable GetPropertySource(int DocID)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@DocID", DocID);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "usp_GetPropertySource", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }

        public DataTable GetPropertyUnused(int DocID)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@DocID", DocID);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "usp_GetUnusedModuleProperties", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }

        public DataTable GetBRPrintDocumentByDocument(int DocID)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@DocID", DocID);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "usp_GetPropertySource", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }
        public int SyncMickProperties(int DocID)
        {

            SqlConnection con = null;
            int returnvalue = -1;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@DocID", DocID);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                //Updated SP name for GTIN
                returnvalue = ExecuteDBQuery(con, CommandType.StoredProcedure, "usp_SyncMickProperties", ref trans, list);
                //returnvalue = ExecuteDBQuery(con, CommandType.StoredProcedure, "usp_SyncMickProperties_NEW", ref trans, list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return returnvalue;
        }
        /*
                public int DeleteSelectedProperty(int DocID, string PropertyName, int PropertyIndex, string PropertyValue, int res)
                {

                    SqlConnection con = null;
                    int returnvalue = -1;
                    SqlTransaction trans = null;
                    try
                    {
                        ClearList();
                        // Creating a connections
                        AddSQLParameter("@DocID", DocID);
                        AddSQLParameter("@PropertyName", PropertyName);
                        AddSQLParameter("@PropertyIndex", PropertyIndex);
                        AddSQLParameter("@PropertyValue", PropertyValue);
                        AddSQLParameter("@DeleteValue", res);
                        con = new SqlConnection(AppConfig.ConnString);
                        OpenConnection(ref con, ref trans);
                        returnvalue = ExecuteDBQuery(con, CommandType.StoredProcedure, "usp_DeleteModuleProperty", ref trans, list);
                    }
                    finally
                    {
                        ClearList();
                        CloseConnection(ref con, ref trans);
                        if (trans != null)
                        {
                            trans = null;
                        }
                    }
                    return returnvalue;
                }*/


        public int DeleteSelectedProperty(int DocID, string PropertyName, int PropertyIndex, string PropertyValue)
        {

            SqlConnection con = null;
            int returnvalue = -1;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@DocID", DocID);
                AddSQLParameter("@PropertyName", PropertyName);
                AddSQLParameter("@PropertyIndex", PropertyIndex);
                AddSQLParameter("@PropertyValue", PropertyValue);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                returnvalue = ExecuteDBQuery(con, CommandType.StoredProcedure, "usp_DeleteModuleProperty", ref trans, list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return returnvalue;
        }
        public int StepText_SelectCOUNT()
        {
            int returnvalue = -1;
            try
            {
                using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
                {
                    connection.Open();

                    string query = "Select COUNT(*) as Count from tblStepText";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                returnvalue = Convert.ToInt32(reader["Count"]);
                            }
                        }

                    }
                }


            }
            catch (Exception ex)
            {
                throw ex;
            }
            return returnvalue;
        }

        public DataTable StepText_Select(string StepTextName)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                AddSQLParameter("@StepTextName", StepTextName);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "SP_StepTextSelect", list);


            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }
        public bool DocProperty_Apply(int DocID, string PropertyName, int PropertyIndex, string PropertyValue, string PropertyUOM, int MaterialID, int WorkCenterID)
        {

            SqlConnection con = null;
            bool returnvalue = false;
            SqlTransaction trans = null;
            string ErrorMessage;
            /*if (PropertyUOM.Length > 0)
            {
                PropertyUOM = PropertyUOM;
            }*/

            if (string.IsNullOrEmpty(PropertyUOM))
                PropertyUOM = "";

            try
            {
                con = new SqlConnection(AppConfig.ConnString);
                using (SqlCommand cmd = new SqlCommand("SP_DocPropertyApply", con))
                {
                    OpenConnection(ref con, ref trans);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@DocID", DocID);
                    cmd.Parameters.AddWithValue("@PropertyName", PropertyName);
                    cmd.Parameters.AddWithValue("@PropertyIndex", PropertyIndex);
                    cmd.Parameters.AddWithValue("@PropertyValue", PropertyValue);
                    //cmd.Parameters.AddWithValue("@PropertyUOM", PropertyUOM);

                    if (!string.IsNullOrEmpty(PropertyUOM))
                    {
                        cmd.Parameters.AddWithValue("@PropertyUOM", PropertyUOM);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@PropertyUOM", DBNull.Value);
                    }

                    cmd.Parameters.AddWithValue("@MaterialID", MaterialID <= 0 ? (object)DBNull.Value : MaterialID);
                    cmd.Parameters.AddWithValue("@WorkCenterID", WorkCenterID <= 0 ? (object)DBNull.Value : WorkCenterID);



                    SqlParameter errorMessage = new SqlParameter("@ErrorMsg", SqlDbType.VarChar, 250);
                    errorMessage.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(errorMessage);

                    SqlDataReader reader = cmd.ExecuteReader();

                    ErrorMessage = (string)errorMessage.Value;
                    if (!string.IsNullOrEmpty(ErrorMessage))
                    {
                        returnvalue = true;
                    }

                }
            }

            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return returnvalue;
        }
        public void DocProperty_DeleteAllExcept(int DocID, string PropertyInstanceList)

        {
            const string MyProcName = "DocProperty_DeleteAllExcept";

            string query = "Delete From tblDocProperty WHERE DocID = " + @DocID + " AND PropertyInstance NOT IN " + @PropertyInstanceList;

            using (SqlConnection connection = new SqlConnection(AppConfig.ConnString))
            {
                try
                {

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        connection.Open();
                        command.Parameters.AddWithValue("@DocID", DocID);
                        command.Parameters.AddWithValue("@PropertyInstanceList", PropertyInstanceList);

                        command.ExecuteNonQuery();
                    }

                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public DataTable DocMaterial_Select(int DocID, bool IncludeZeroAmount)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                AddSQLParameter("@DocID", DocID);
                AddSQLParameter("@IncludeZeroAmount", IncludeZeroAmount);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "SP_DocMaterialSelect", list);

            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }



        #region EditModule
        public DataTable GetDocumentModule(int DocID)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@DocID", DocID);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "usp_DocModule_Auth", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }

        public DataTable GetPacketbyDoc(int DocID)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@DocID", DocID);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "SP_GetPacketByDocument", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }

        public DataTable GetModuleAll(int OpUnitID)
        {

            SqlConnection con = null;
            DataTable dtResultset = null;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@OpUnitID", OpUnitID);
                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                dtResultset = ExecuteDataTable(con, CommandType.StoredProcedure, "SP_GetModuleByOpUnit", list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return dtResultset;
        }
        public void DeleteModule(int DocID, int Sequence)
        {
            SqlConnection con;

            using (con = new SqlConnection(AppConfig.ConnString))
            {
                con.Open();
                string query = "Delete From tblDocModule WHERE DocID= @DocID  AND ModuleSequence > @Sequence";
                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    cmd.Parameters.AddWithValue("@DocID", DocID);
                    cmd.Parameters.AddWithValue("@Sequence", Sequence);
                    cmd.ExecuteNonQuery();
                }


            }


        }
        public int InsertUpdateDocModule(int DocID, string ModuleBaseName, string ModuleComment, string Packet, int ModuleSequence, string ModuleUsage, string ModuleTitle, int ModuleRepetitions, string ModuleTableRows)
        {

            SqlConnection con = null;
            int returnvalue = -1;
            SqlTransaction trans = null;
            try
            {
                ClearList();
                // Creating a connections
                AddSQLParameter("@DocID", DocID);
                AddSQLParameter("@ModuleBaseName", ModuleBaseName);
                AddSQLParameter("@ModuleComment", ModuleComment);
                AddSQLParameter("@Packet", Packet);
                AddSQLParameter("@ModuleSequence", ModuleSequence);
                AddSQLParameter("@ModuleUsage", ModuleUsage);
                AddSQLParameter("@ModuleTitle", ModuleTitle);
                AddSQLParameter("@ModuleRepetitions", ModuleRepetitions);
                AddSQLParameter("@ModuleTableRows", ModuleTableRows);

                con = new SqlConnection(AppConfig.ConnString);
                OpenConnection(ref con, ref trans);
                returnvalue = ExecuteDBQuery(con, CommandType.StoredProcedure, "usp_InsertUpdateDocModule", ref trans, list);
            }
            finally
            {
                ClearList();
                CloseConnection(ref con, ref trans);
                if (trans != null)
                {
                    trans = null;
                }
            }
            return returnvalue;
        }

        #endregion

        #region CommonMethods

        protected void ClearList()
        {
            list.Clear();
            list = new List<KeyValuePair<string, string>>();
        }

        public string GetConnectionString()
        {
            string strConnectionString = string.Empty;
            try
            {
                strConnectionString = ConfigurationManager.AppSettings["DBConnection"].ToString();
            }
            catch (Exception ex)
            {
            }
            return strConnectionString;
        }
        protected void OpenConnection(ref SqlConnection ret_conn, ref SqlTransaction ret_tran, bool isTran = false)
        {
            dbConnectionDetails = new DBConnectionDetails();
            dbConnectionDetails.ConnectStartTime = DateTime.Now;//.ToString("yyyy-MM-dd HH:mm:ss");
            dbConnectionDetails.DatabaseConnections += 1;
            try
            {
                if (ret_conn.State == ConnectionState.Closed)
                {
                    ret_conn.Open();
                    if (isTran)
                        ret_tran = ret_conn.BeginTransaction();
                    dbConnectionDetails.DatabaseRequestsIn += 1;
                    dbConnectionDetails.DatabaseRequestsOut += 1;

                }
                if (list == null)
                    list.Clear();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        protected void CloseConnection(ref SqlConnection ret_conn, ref SqlTransaction ret_tran)
        {
            dbConnectionDetails = new DBConnectionDetails();
            
            try
            {
                if (ret_conn.State == ConnectionState.Open)
                {
                    if (ret_tran != null)
                        ret_tran.Dispose();
                    ret_conn.Close();
                    ret_conn.Dispose();
                }
                if (list != null)
                    list.Clear();
					dbConnectionDetails.ConnectEndTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        protected DataTable ExecuteDataTable(SqlConnection v_sqlConn, CommandType v_cmdType, string SPName, List<KeyValuePair<string, string>> listparam = null)
        {
            SqlCommand sqlCmd = new SqlCommand();
            DataTable dtReturnvalue = new DataTable();
            try
            {
                sqlCmd.CommandText = SPName;
                sqlCmd.CommandType = v_cmdType;
                sqlCmd.Connection = v_sqlConn;
                sqlCmd.CommandTimeout = 500;
                if (listparam != null && listparam.Count > 0)
                {
                    foreach (KeyValuePair<string, string> strkey in listparam)
                        sqlCmd.Parameters.AddWithValue(strkey.Key, strkey.Value);
                }
                SqlDataAdapter adaptor = new SqlDataAdapter(sqlCmd);

                adaptor.Fill(dtReturnvalue);
                return dtReturnvalue;
            }
            finally
            {
                if (sqlCmd != null)
                    sqlCmd.Dispose();
                if (dtReturnvalue != null)
                    dtReturnvalue.Dispose();
            }
        }

        public List<KeyValuePair<string, string>> AddSQLParameter(string v_ParameterName, object v_ParameterValue)
        {
            list.Add(new KeyValuePair<string, string>(v_ParameterName, v_ParameterValue.ToString().Trim()));

            return list;
        }

        protected int ExecuteDBQuery(SqlConnection v_sqlConn, CommandType v_cmdType, string SPName, ref SqlTransaction r_sqlTran, List<KeyValuePair<string, string>> listparam = null)
        {
            SqlCommand sqlCmd = new SqlCommand();
            int returnValue = -1;
            try
            {
                sqlCmd.CommandText = SPName;
                sqlCmd.CommandType = v_cmdType;
                sqlCmd.Connection = v_sqlConn;
                sqlCmd.CommandTimeout = 500;
                if (r_sqlTran != null)
                    sqlCmd.Transaction = r_sqlTran;
                // If objSqlTrans IsNot Nothing Then
                // sqlCmd.Transaction = objSqlTrans
                // End If
                if (listparam != null && listparam.Count > 0)
                {
                    foreach (KeyValuePair<string, string> strkey in listparam)
                        sqlCmd.Parameters.AddWithValue(strkey.Key, strkey.Value);
                }
                SqlParameter parm = new SqlParameter("@ReturnVal", SqlDbType.Int);
                parm.Direction = ParameterDirection.ReturnValue;
                sqlCmd.Parameters.Add(parm);
                sqlCmd.ExecuteNonQuery();
                returnValue = Convert.ToInt32(parm.Value);

                return returnValue;
            }
            finally
            {
                if (sqlCmd != null)
                    sqlCmd.Dispose();
            }
        }

        #endregion


    }
}
