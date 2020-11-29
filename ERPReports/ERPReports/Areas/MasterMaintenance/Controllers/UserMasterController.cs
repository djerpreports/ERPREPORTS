using Microsoft.AspNet.Identity;
using System.Data.SqlClient;
using ERPReports.Areas.MasterMaintenance.Models;
using ERPReports.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web.Mvc;
namespace ERPReports.Areas.MasterMaintenance.Controllers
{
    public class UserMasterController : Controller
    {
        PasswordHasher ph = new PasswordHasher();
        DataHelper dataHelper = new DataHelper();
        List<string> modelErrors = new List<string>();
        bool error = false;
        string errmsg = "";

        public ActionResult Index()
        {
            return View("UserMaster");
        }
        public ActionResult GetUserList()
        {
            List<mUser> data = new List<mUser>();
            DataTableHelper TypeHelper = new DataTableHelper();

            int start = Convert.ToInt32(Request["start"]);
            int length = Convert.ToInt32(Request["length"]);
            string searchValue = Request["search[value]"];
            string sortColumnName = Request["columns[" + Request["order[0][column]"] + "][data]"];
            string sortDirection = Request["order[0][dir]"];
            //string UserCategory = Request["UserCategory"];

            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ConnectionString.ToString()))

                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {
                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "spUser_GetUserList";
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                data.Add(new mUser
                                {
                                    ID = Convert.ToInt32(sdr["ID"]),
                                    UserName = sdr["UserName"].ToString(),
                                    FirstName = sdr["FirstName"].ToString(),
                                    MiddleName = sdr["MiddleName"].ToString(),
                                    LastName = sdr["LastName"].ToString(),
                                    EmailAddress = sdr["EmailAddress"].ToString(),
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
                    errmsg = "An error occured: " + err.Message.ToString();

                return Json(new { success = false, msg = errmsg }, JsonRequestBehavior.AllowGet);
            }
            int totalrows = data.Count;
            if (!string.IsNullOrEmpty(searchValue))//filter
                data = data.Where(x =>
                                    x.UserName.ToLower().Contains(searchValue.ToLower()) ||
                                    x.FirstName.ToString().ToLower().Contains(searchValue.ToLower()) ||
                                    x.MiddleName.ToString().ToLower().Contains(searchValue.ToLower()) ||
                                    x.LastName.ToString().ToLower().Contains(searchValue.ToLower()) ||
                                    x.EmailAddress.ToString().ToLower().Contains(searchValue.ToLower())
                                 ).ToList<mUser>();

            int totalrowsafterfiltering = data.Count;
            if (sortDirection == "asc")
                data = data.OrderBy(x => TypeHelper.GetPropertyValue(x, sortColumnName)).ToList();

            if (sortDirection == "desc")
                data = data.OrderByDescending(x => TypeHelper.GetPropertyValue(x, sortColumnName)).ToList();

            data = data.Skip(start).Take(length).ToList<mUser>();


            return Json(new { data = data, draw = Request["draw"], recordsTotal = totalrows, recordsFiltered = totalrowsafterfiltering }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult ValidateEmailAddress(string EmailAddress)
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
                        cmdSql.CommandText = "spUser_ValidateEmailAddress";
                        cmdSql.Parameters.Clear();
                        cmdSql.Parameters.AddWithValue("@EmailAddress", EmailAddress);
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
        public ActionResult SaveUser(mUser User, string OldPassword, bool DidPasswordChange)
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
                            if (User.ID == 0)
                                User.Password = ph.HashPassword(User.Password).ToString();
                            else
                            {
                                if (DidPasswordChange)
                                {
                                    if (ph.VerifyHashedPassword(OldPassword.ToString(), User.Password.ToString()).ToString() != "Success")
                                    {
                                        User.Password = ph.HashPassword(User.Password).ToString();
                                    }
                                }
                            }

                            cmdSql.CommandType = CommandType.StoredProcedure;
                            cmdSql.CommandText = "spUser_InsertUpdate";
                            cmdSql.Parameters.Clear();
                            cmdSql.Parameters.AddWithValue("@ID", User.ID);
                            cmdSql.Parameters.AddWithValue("@UserName", User.UserName);
                            cmdSql.Parameters.AddWithValue("@Password", User.Password);
                            cmdSql.Parameters.AddWithValue("@FirstName", User.FirstName);
                            cmdSql.Parameters.AddWithValue("@MiddleName", User.MiddleName);
                            cmdSql.Parameters.AddWithValue("@LastName", User.LastName);
                            cmdSql.Parameters.AddWithValue("@EmailAddress", User.EmailAddress);
                            cmdSql.Parameters.AddWithValue("@UserID", Session["UserName"]);
                            cmdSql.ExecuteNonQuery();
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
                return Json(new { success = true, msg = "User was successfully " + (User.ID == 0 ? " saved." : " updated.") });
            }
        }
        public ActionResult GetUserDetails(int ID)
        {
            mUser userDetails = new mUser();
            ArrayList CostCenterList = new ArrayList();
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {
                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "spUser_GetDetails";
                        cmdSql.Parameters.Clear();
                        cmdSql.Parameters.AddWithValue("@ID", ID);
                        using (SqlDataReader reader = cmdSql.ExecuteReader())
                        {
                            if (!reader.Read())
                                throw new InvalidOperationException("No records found.");

                            userDetails.ID = Convert.ToInt32(reader["ID"]);
                            userDetails.UserName = reader["UserName"].ToString();
                            userDetails.Password = reader["Password"].ToString();
                            userDetails.FirstName = reader["FirstName"].ToString();
                            userDetails.MiddleName = reader["MiddleName"].ToString();
                            userDetails.LastName = reader["LastName"].ToString();
                            userDetails.EmailAddress = reader["EmailAddress"].ToString();
                            reader.Close();
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
            return Json(new { success = true, data = new { userData = userDetails } }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult DeleteUser(int ID)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {
                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "spUser_Delete";

                        cmdSql.Parameters.Clear();
                        cmdSql.Parameters.AddWithValue("@ID", ID);
                        cmdSql.Parameters.AddWithValue("@UserID", Session["UserName"]);
                        cmdSql.ExecuteNonQuery();
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
            return Json(new { success = true, msg = "User was successfully deleted." });

        }

        public ActionResult GetUserAccess(int ID)
        {
            ArrayList userMenuList = new ArrayList();
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ConnectionString.ToString()))

                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {
                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "spUserPageAccess_GetAccessList";
                        cmdSql.Parameters.Clear();
                        cmdSql.Parameters.AddWithValue("@ID", ID);
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                //var userid = Session["UserID"].ToString();
                                userMenuList.Add(new
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
                                    Status = sdr["Status"].ToString() == "" ? 0 : Convert.ToInt32(sdr["Status"]),
                                    ReadAndWrite = sdr["ReadAndWrite"].ToString() == "" ? 0 : Convert.ToInt32(sdr["ReadAndWrite"]),
                                    Delete = sdr["Delete"].ToString() == "" ? 0 : Convert.ToInt32(sdr["Delete"])
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
                    errmsg = "An error occured: " + err.Message.ToString();

                return Json(new { success = false, errors = errmsg }, JsonRequestBehavior.AllowGet);
            }
            return Json(new { success = true, data = userMenuList }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult SaveUserAccess(List<mUserPageAccess> userAccessList, int UserID)
        {
            string endMsg = " saved.";
            ModelState.Remove("ID");
            if (ModelState.IsValid)
            {
                try
                {
                    using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ToString()))
                    {
                        conn.Open();
                        SqlTransaction transaction;
                        transaction = conn.BeginTransaction();
                        try
                        {
                            foreach (mUserPageAccess userAccess in userAccessList)
                            {
                                using (SqlCommand cmdSql = conn.CreateCommand())
                                {
                                    cmdSql.Connection = conn;
                                    cmdSql.Transaction = transaction;
                                    cmdSql.CommandType = CommandType.StoredProcedure;
                                    cmdSql.CommandText = "spUserPageAccess_INSERT_UPDATE";

                                    cmdSql.Parameters.Clear();
                                    cmdSql.Parameters.AddWithValue("@UserID", Convert.ToInt32(userAccess.UserID));
                                    cmdSql.Parameters.AddWithValue("@PageID", userAccess.PageID);
                                    cmdSql.Parameters.AddWithValue("@Status", userAccess.Status);
                                    cmdSql.Parameters.AddWithValue("@ReadAndWrite", Convert.ToBoolean(userAccess.ReadAndWrite));
                                    cmdSql.Parameters.AddWithValue("@Delete", Convert.ToBoolean(userAccess.Delete));
                                    cmdSql.Parameters.AddWithValue("@CreateID", Session["ID"].ToString());
                                    SqlParameter ErrorMessage = cmdSql.Parameters.Add("@ErrorMessage", SqlDbType.VarChar, 50);
                                    SqlParameter IsError = cmdSql.Parameters.Add("@IsError", SqlDbType.Bit);

                                    IsError.Direction = ParameterDirection.Output;
                                    ErrorMessage.Direction = ParameterDirection.Output;

                                    cmdSql.ExecuteNonQuery();

                                    error = Convert.ToBoolean(IsError.Value);
                                    if (error)
                                    {
                                        modelErrors.Add(ErrorMessage.Value.ToString());
                                        //throw new System.InvalidOperationException(ErrorMessage.Value.ToString());
                                    }
                                }
                            }
                            transaction.Commit();
                        }
                        catch (Exception err)
                        {
                            if (err.InnerException != null)
                                modelErrors.Add("An error occured: " + err.InnerException.ToString());
                            else
                                modelErrors.Add("An error occured: " + err.Message.ToString());
                            error = true;
                        }

                        conn.Close();
                    }
                }
                catch (Exception err)
                {
                    if (err.InnerException != null)
                        modelErrors.Add("An error occured: " + err.InnerException.ToString());
                    else
                        modelErrors.Add("An error occured: " + err.Message.ToString());
                    error = true;
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
                return Json(new { success = true, msg = "User Page Access was successfully " + endMsg });
        }
    }
}
