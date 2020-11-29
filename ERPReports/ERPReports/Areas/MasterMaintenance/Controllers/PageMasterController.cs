using System.Data.SqlClient;
using ERPReports.Areas.MasterMaintenance.Models;
using ERPReports.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Web.Mvc;
namespace ERPReports.Areas.MasterMaintenance.Controllers
{
    public class PageMasterController : Controller
    {
        DataHelper dataHelper = new DataHelper();
        List<string> modelErrors = new List<string>();
        bool error = false;
        string errmsg = "";
        public ActionResult Index()
        {
            return View("PageMaster");
        }
        public ActionResult GetPageList()
        {
            List<mPage> data = new List<mPage>();
            DataTableHelper TypeHelper = new DataTableHelper();

            int start = Convert.ToInt32(Request["start"]);
            int length = Convert.ToInt32(Request["length"]);
            string searchValue = Request["search[value]"];
            string sortColumnName = Request["columns[" + Request["order[0][column]"] + "][data]"];
            string sortDirection = Request["order[0][dir]"];

            string GroupLabel = Request["columns[0][search][value]"];
            string PageName = Request["columns[1][search][value]"];
            string PageLabel = Request["columns[2][search][value]"];
            string URL = Request["columns[3][search][value]"];
            string HasSub = Request["columns[4][search][value]"];
            string ParentMenu = Request["columns[5][search][value]"];
            string ParentOrder = Request["columns[6][search][value]"];
            string Order = Request["columns[7][search][value]"];

            int RetTotalRecords = 0;
            int RetFilteredRecords = 0;
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ConnectionString.ToString()))

                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {
                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "spPage_GetPageList";
                        cmdSql.Parameters.AddWithValue("@GroupLabel", GroupLabel);
                        cmdSql.Parameters.AddWithValue("@PageName", PageName);
                        cmdSql.Parameters.AddWithValue("@PageLabel", PageLabel);
                        cmdSql.Parameters.AddWithValue("@URL", URL);
                        cmdSql.Parameters.AddWithValue("@HasSub", HasSub);
                        cmdSql.Parameters.AddWithValue("@ParentMenu", ParentMenu);
                        cmdSql.Parameters.AddWithValue("@ParentOrder", ParentOrder);
                        cmdSql.Parameters.AddWithValue("@Order", Order);

                        cmdSql.Parameters.AddWithValue("@StartPage", start);
                        cmdSql.Parameters.AddWithValue("@RowCount", length);
                        cmdSql.Parameters.AddWithValue("@SearchValue", searchValue);
                        cmdSql.Parameters.AddWithValue("@SortColumnName", sortColumnName);
                        cmdSql.Parameters.AddWithValue("@SortDirection", sortDirection);
                        SqlParameter TotalRecords = cmdSql.Parameters.Add("@TotalRecords", SqlDbType.Int);
                        SqlParameter FilteredRecords = cmdSql.Parameters.Add("@FilteredRecords", SqlDbType.Int);
                        TotalRecords.Direction = ParameterDirection.Output;
                        FilteredRecords.Direction = ParameterDirection.Output;

                        cmdSql.ExecuteNonQuery();

                        RetTotalRecords = Convert.ToInt32(TotalRecords.Value);
                        RetFilteredRecords = Convert.ToInt32(FilteredRecords.Value);
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                data.Add(new mPage
                                {
                                    ID = Convert.ToInt32(sdr["ID"]),
                                    GroupLabel = sdr["GroupLabel"].ToString(),
                                    PageName = sdr["PageName"].ToString(),
                                    PageLabel = sdr["PageLabel"].ToString(),
                                    URL = sdr["URL"].ToString(),
                                    HasSub = Convert.ToInt32(sdr["HasSub"]),
                                    ParentMenu = sdr["ParentMenu"].ToString(),
                                    ParentOrder = Convert.ToInt32(sdr["ParentOrder"]),
                                    Order = Convert.ToInt32(sdr["Order"]),
                                    Icon = sdr["Icon"].ToString(),
                                });
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

                return Json(new { success = false, msg = errmsg }, JsonRequestBehavior.AllowGet);
            }

            return Json(new { data = data, draw = Request["draw"], recordsTotal = RetTotalRecords, recordsFiltered = RetFilteredRecords }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult ValidatePageName(string PageName)
        {
            bool isValid = true;
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {
                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "spPage_ValidatePageName";
                        cmdSql.Parameters.AddWithValue("@PageName", PageName);
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            if (sdr.HasRows)
                                isValid = false;
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception err)
            {
                if (err.InnerException != null)
                    errmsg = "An error occured: " + err.InnerException.ToString();
                else
                    errmsg = "An error occured: " + err.Message.ToString(); ;
                error = true;
            }
            if (error)
                return Json(new { success = false, errors = errmsg }, JsonRequestBehavior.AllowGet);
            else
            {
                return Json(new { success = true, data = new { isValid = isValid } });
            }
        }
        public ActionResult SavePage(mPage Page)
        {
            string endMsg = "";
            ModelState.Remove("ID");
            if (ModelState.IsValid)
            {
                try
                {
                    using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ToString()))
                    {
                        conn.Open();
                        using (SqlCommand cmdSql = conn.CreateCommand())
                        {
                            cmdSql.CommandType = CommandType.StoredProcedure;
                            cmdSql.CommandText = "spPage_InsertUpdate";
                            cmdSql.Parameters.Clear();
                            cmdSql.Parameters.AddWithValue("@ID", Page.ID);
                            cmdSql.Parameters.AddWithValue("@GroupLabel", Page.GroupLabel == null ? "" : Page.GroupLabel);
                            cmdSql.Parameters.AddWithValue("@PageName", Page.PageName);
                            cmdSql.Parameters.AddWithValue("@PageLabel", Page.PageLabel);
                            cmdSql.Parameters.AddWithValue("@URL", Page.URL);
                            cmdSql.Parameters.AddWithValue("@HasSub", Page.HasSub);
                            cmdSql.Parameters.AddWithValue("@ParentMenu", Page.ParentMenu);
                            cmdSql.Parameters.AddWithValue("@ParentOrder", Page.ParentOrder);
                            cmdSql.Parameters.AddWithValue("@Order", Page.Order);
                            cmdSql.Parameters.AddWithValue("@Icon", Page.Icon);
                            cmdSql.Parameters.AddWithValue("@CreateID", Session["UserName"]);
                            SqlParameter EndMsg = cmdSql.Parameters.Add("@EndMsg", SqlDbType.VarChar, 200);
                            SqlParameter ErrorMessage = cmdSql.Parameters.Add("@ErrorMessage", SqlDbType.VarChar, 200);
                            SqlParameter Error = cmdSql.Parameters.Add("@IsError", SqlDbType.Bit);

                            EndMsg.Direction = ParameterDirection.Output;
                            Error.Direction = ParameterDirection.Output;
                            ErrorMessage.Direction = ParameterDirection.Output;

                            cmdSql.ExecuteNonQuery();

                            error = Convert.ToBoolean(Error.Value);
                            if (error)
                                modelErrors.Add(ErrorMessage.Value.ToString());

                            endMsg = EndMsg.Value.ToString();
                        }
                        conn.Close();
                    }
                }
                catch (Exception err)
                {
                    string errmsg;
                    if (err.InnerException != null)
                        errmsg = "Error: " + err.InnerException.ToString();
                    else
                        errmsg = "Error: " + err.Message.ToString();

                    return Json(new { success = false, errors = errmsg }, JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                foreach (var modelStateKey in ViewData.ModelState.Keys)
                {
                    var modelStateVal = ViewData.ModelState[modelStateKey];
                    foreach (var error in modelStateVal.Errors)
                    {
                        var key = modelStateKey;
                        var errMessage = error.ErrorMessage;
                        var exception = error.Exception;
                        modelErrors.Add(errMessage);
                    }
                }
            }
            if (modelErrors.Count != 0 || error)
                return Json(new { success = false, errors = modelErrors });
            else
            {
                return Json(new { success = true, msg = "Page was successfully " + endMsg });
            }
        }
        public ActionResult GetPageDetails(string PageName)
        {
            mPage pageDetails = new mPage();
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {
                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "spPage_GetPageDetails";
                        cmdSql.Parameters.AddWithValue("@PageName", PageName);
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            if (!sdr.Read())
                                throw new InvalidOperationException("No records found.");

                            pageDetails.ID = Convert.ToInt32(sdr["ID"]);
                            pageDetails.GroupLabel = sdr["GroupLabel"].ToString();
                            pageDetails.PageName = sdr["PageName"].ToString();
                            pageDetails.PageLabel = sdr["PageLabel"].ToString();
                            pageDetails.URL = sdr["URL"].ToString();
                            pageDetails.HasSub = Convert.ToInt32(sdr["HasSub"]);
                            pageDetails.ParentMenu = sdr["ParentMenu"].ToString();
                            pageDetails.ParentOrder = Convert.ToInt32(sdr["ParentOrder"]);
                            pageDetails.Order = Convert.ToInt32(sdr["Order"]);
                            pageDetails.Icon = sdr["Icon"].ToString();
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception err)
            {
                string errmsg;
                if (err.InnerException != null)
                    errmsg = "Error: " + err.InnerException.ToString();
                else
                    errmsg = "Error: " + err.Message.ToString();

                return Json(new { success = false, errors = errmsg }, JsonRequestBehavior.AllowGet);
            }
            return Json(new { success = true, data = new { pageDetails = pageDetails } }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult DeletePage(string PageName)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {
                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "spPage_Delete";

                        cmdSql.Parameters.Clear();
                        cmdSql.Parameters.AddWithValue("@PageName", PageName);
                        cmdSql.Parameters.AddWithValue("@UpdateID", Session["UserName"]);

                        SqlParameter Error = cmdSql.Parameters.Add("@IsError", SqlDbType.Bit);
                        SqlParameter ErrorMessage = cmdSql.Parameters.Add("@ErrorMessage", SqlDbType.VarChar, 50);

                        Error.Direction = ParameterDirection.Output;
                        ErrorMessage.Direction = ParameterDirection.Output;

                        cmdSql.ExecuteNonQuery();

                        error = Convert.ToBoolean(Error.Value);
                        if (error)
                            modelErrors.Add(ErrorMessage.Value.ToString());
                    }
                    conn.Close();
                }
            }
            catch (Exception err)
            {
                string errmsg;
                if (err.InnerException != null)
                    errmsg = "Error: " + err.InnerException.ToString();
                else
                    errmsg = "Error: " + err.Message.ToString();

                return Json(new { success = false, errors = errmsg }, JsonRequestBehavior.AllowGet);
            }
            return Json(new { success = true, msg = "Page was successfully deleted." });

        }
    }
}
