using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace ERPReports.Models
{
    public class DataHelper
    {
        string retVal = "";
        public String GetData(string field, string table, string condition, string asVar = "")
        {
            retVal = "";
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ConnectionString.ToString()))
                {
                    conn.Open();
                    string sql = "SELECT " + field + " FROM " + table + " WHERE " + condition;
                    using (SqlCommand comm = new SqlCommand(sql, conn))
                    {
                        SqlDataReader reader = comm.ExecuteReader();
                        if (reader.Read())
                        {
                            if (asVar != "")
                                retVal = reader[asVar].ToString();
                            else
                                retVal = reader[field].ToString();
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception err)
            {
                string errmsg;
                if (err.InnerException != null)
                    errmsg = "An error occured: " + err.InnerException.ToString();
                else
                    errmsg = "An error occured: " + err.ToString();
            }
            return retVal;
        }
        public DataTable UpdateHeaderID(DataTable table, int HeaderID)
        {
            foreach (DataRow row in table.Rows)
            {
                row["HeaderID"] = HeaderID;
            }
            return table;
        }
        public DataTable UpdateDates(DataTable table)
        {
            foreach (DataRow row in table.Rows)
            {
                row["CreateDate"] = DateTime.Now;
                row["UpdateDate"] = DateTime.Now;
            }
            return table;
        }
        public DataTable UpdateCreateCols(DataTable table, string CreateID, String UpdateID)
        {
            foreach (DataRow row in table.Rows)
            {
                if (CreateID != "")
                {
                    row["CreateID"] = CreateID;
                    row["UpdateID"] = UpdateID;
                }
                else if (UpdateID != "")
                {
                    row["UpdateID"] = UpdateID;
                }
            }
            return table;
        }
    }
}
