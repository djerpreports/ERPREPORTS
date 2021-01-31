using Microsoft.AspNet.Identity;
using ERPReports.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Web.Mvc;
using System.Data.SqlClient;
namespace ERPReports.Controllers
{
    [AllowAnonymous]
    public class LoginController : Controller
    {
        List<string> modelErrors = new List<string>();
        string errmsg;

        // GET: Login
        public ActionResult Index()
        {
            if (Session["ID"] != null)
            {
                return RedirectToAction("Index", "Home", new { area = "" });
            }
            else
            {
                return View("Login");
            }
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult LoginEntry(Login data)
        {
            try
            {
                PasswordHasher ph = new PasswordHasher();
                DataHelper dataHelper = new DataHelper();
                int ID = 0;
                string UserName = "";
                string Password = "";
                string FirstName = "";
                string MiddleName = "";
                string LastName = "";
                string EmailAddress = "";
                string username = data.UserName;
                string password = ph.HashPassword(data.Password.ToString());
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ERPReports"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {
                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "spUserLogin";

                        cmdSql.Parameters.Clear();
                        cmdSql.Parameters.AddWithValue("@UserName", username);
                        cmdSql.Parameters.AddWithValue("@Password", password);

                        using (SqlDataReader rdSql = cmdSql.ExecuteReader())
                        {
                            if (rdSql.Read())
                            {
                                Password = rdSql["Password"].ToString();
                                if (ph.VerifyHashedPassword(Password.ToString(), data.Password.ToString()).ToString() != "Success")
                                {
                                    errmsg = "Invalid UserName or Password. Please try again.";
                                }else
                                {

                                    ID = Convert.ToInt32(rdSql["ID"]);
                                    UserName = rdSql["UserName"].ToString();
                                    FirstName = rdSql["FirstName"].ToString();
                                    MiddleName = rdSql["MiddleName"].ToString();
                                    LastName = rdSql["LastName"].ToString();
                                    EmailAddress = rdSql["EmailAddress"].ToString();


                                    Session["ID"] = ID;
                                    Session["UserName"] = UserName;
                                    Session["FirstName"] = FirstName;
                                    Session["MiddleName"] = MiddleName;
                                    Session["LastName"] = LastName;
                                    Session["EmailAddress"] = EmailAddress;
                                }
                            }
                            else
                            {
                                errmsg = "Invalid UserName or Password. Please try again.";
                            }
                        }
                    }
                }
            }
            catch (Exception err)
            {
                if (err.InnerException != null)
                    errmsg = "Error: " + err.InnerException.ToString();
                else
                    errmsg = "Error: " + err.Message.ToString();
            }
            if (errmsg != null)
                return Json(new { success = true, data = new { error = true, errmsg = errmsg } });
            else
            {
                if (data.IsFromOld)
                    return RedirectToAction("Index", "Home", new { area = "" });
                else
                    return Json(new { success = true, data = new { error = false } });
            }
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Logout(int ID)
        {
            Session.Abandon();
            return RedirectToAction("Index", "Login", new { area = "" });
        }
        public ActionResult SessionError()
        {
            return Json(new { success = false, type = "Login", errors = "Session has expired. Please login again. Thank you." }, JsonRequestBehavior.AllowGet);
        }
    }
}
