using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace ERPReports.Models
{
    public class Select2Data
    {
        public ArrayList GetDataForSelect2(string tblName, string[] value_text, string condition, string specialAction = "", params string[] concatForVal)
        {
            ArrayList List = new ArrayList();

            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["CASDB"].ConnectionString.ToString()))

                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {
                        cmdSql.CommandType = CommandType.Text;
                        cmdSql.CommandText = "SELECT * FROM " + tblName + " WHERE " + condition;
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (specialAction == "concatForValue")
                                {
                                    var val = sdr[concatForVal[0]].ToString() + "." + sdr[concatForVal[1]].ToString();
                                    List.Add(new { value = val, text = (value_text[0].ToString() == value_text[1].ToString()) ? sdr[value_text[0]].ToString() : val + " - " + sdr[value_text[1]].ToString() });
                                }
                                else if (specialAction == "concatForString")
                                {
                                    var text = sdr[value_text[1]].ToString() + ". " + sdr[concatForVal[0]].ToString() + " " + sdr[concatForVal[1]].ToString();
                                    List.Add(new { value = sdr[value_text[0]].ToString(), text = text });
                                }
                                else
                                {
                                    List.Add(new { value = sdr[value_text[0]].ToString(), text = (value_text[0].ToString() == value_text[1].ToString()) ? sdr[value_text[0]].ToString() : sdr[value_text[0]].ToString() + " - " + sdr[value_text[1]].ToString() });
                                }

                            }

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
            return List;
        }
    }
}
