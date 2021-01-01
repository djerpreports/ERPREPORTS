using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;
using ERPReports.Models;
using System.IO;
using ERPReports.Areas.Reports.Models;
namespace ERPReports.Areas.Reports.Controllers
{
    public class LSPDirectMaterialController : Controller
    {
        // GET: Reports/LSPDirectMaterial
        public ActionResult Index()
        {
            return View("LSPDirectMaterial");
        }

        public ActionResult GetSelect2DataModel()
        {
            ArrayList results = new ArrayList();
            string val = Request.QueryString["q"];
            string id = Request.QueryString["id"];
            string text = Request.QueryString["text"];
            string table = Request.QueryString["table"];
            string db = Request.QueryString["db"];
            string condition = Request.QueryString["condition"] == null ? "" : Request.QueryString["condition"];
            string isDistict = Request.QueryString["isDistict"] == null ? "" : Request.QueryString["isDistict"];
            string display = Request.QueryString["display"];
            string addOptionVal = Request.QueryString["addOptionVal"];
            string addOptionText = Request.QueryString["addOptionText"];
            string sp = Request.QueryString["sp"];
            string StartProdCode = Request.QueryString["StartProdCode"];
            string EndProdCode = Request.QueryString["EndProdCode"];
            string orderBy = Request.QueryString["orderBy"] == null ? "" : Request.QueryString["orderBy"];

            if (addOptionVal != null && display == "id&text")
                results.Add(new { id = addOptionVal, text = addOptionText });

            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings[db].ConnectionString.ToString()))

                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {


                        #region
                        cmdSql.CommandType = CommandType.Text;
                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "LSP_ERPReport_GetFGItemListPerProdCodeWihtNullSp";
                        cmdSql.Parameters.Clear();
                        cmdSql.Parameters.AddWithValue("@StartProdCode", StartProdCode == null ? "" : StartProdCode);
                        cmdSql.Parameters.AddWithValue("@EndProdCode", EndProdCode == null ? "" : EndProdCode);
                        cmdSql.Parameters.AddWithValue("@Search", val == null ? "" : val);
                        cmdSql.ExecuteNonQuery();
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (display == "id&text")
                                    results.Add(new { id = sdr[id].ToString(), text = sdr[text].ToString() });
                                if (display == "id&id-text")
                                    results.Add(new { id = sdr[id].ToString(), text = sdr[id].ToString() + "-" + sdr[text].ToString() });
                            }

                        }
                    }
                    #endregion
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
            return Json(new { results }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult GenerateDMAndLaborPercentageReport()
        {
            List<DMAndLaborPercentageReport> ProductModel = new List<DMAndLaborPercentageReport>();
            var FinishedGoodsAndSalesReport = Request["FinishedGoodsAndSalesReport"];
            var MiscellaneousTransactionReport = Request["MiscellaneousTransactionReport"];
            var DMAndLaborPercentageReport = Request["DMAndLaborPercentageReport"];
            var StartDate = Request["StartDate"];
            var EndDate = Request["EndDate"];
            var ProductCode1 = Request["ProductCode1"];
            var ProductCode2 = Request["ProductCode2"];
            var Model1 = Request["Model1"];
            var Model2 = Request["Model2"];
            var InventoryTurnoverReport = Request["InventoryTurnoverReport"];
            var ShowDetailedTransaction = Request["ShowDetailedTransaction"];

            string MonthYear = DateTime.Parse(StartDate).ToString("MMMyyyy");
            try
            {
                List<ExcelColumns> datalist = new List<ExcelColumns>();
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LSPI803_App"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {

                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "LSP_Rpt_NewDM_DirectMaterialLaborPercentageReportSp";
                        cmdSql.CommandTimeout = 0;
                        cmdSql.Parameters.Clear();

                        cmdSql.Parameters.AddWithValue("@StartDate", StartDate);
                        cmdSql.Parameters.AddWithValue("@EndDate", EndDate);
                        cmdSql.Parameters.AddWithValue("@StartProdCode", ProductCode1 == null ? "" : ProductCode1);
                        cmdSql.Parameters.AddWithValue("@EndProdCode", ProductCode2 == null ? "" : ProductCode2);
                        cmdSql.Parameters.AddWithValue("@StartModel", Model1 == null ? "" : Model1);
                        cmdSql.Parameters.AddWithValue("@EndModel", Model2 == null ? "" : Model2);
                        cmdSql.ExecuteNonQuery();
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                ProductModel.Add(new DMAndLaborPercentageReport
                                {
                                    item = sdr["item"].ToString().Remove(0, 3),
                                    product_code = sdr["product_code"].ToString().Remove(0, 3),
                                    fam_code = sdr["fam_code"].ToString(),
                                    trans_date = DateTime.Parse(sdr["trans_date"].ToString()).ToString("MMM-yyyy"),
                                    qty_completed = Convert.ToDecimal(sdr["qty_completed"]),
                                    produced_amt = Convert.ToDecimal(sdr["produced_amt"]),
                                    actl_rm_cost = Convert.ToDecimal(sdr["actl_rm_cost"]),
                                    std_rm_cost = Convert.ToDecimal(sdr["std_rm_cost"]),
                                    actl_lbr_cost = Convert.ToDecimal(sdr["actl_lbr_cost"]),
                                    std_lbr_cost = Convert.ToDecimal(sdr["std_lbr_cost"]),
                                });
                            }

                        }
                    }
                    conn.Close();
                }


                var groupedProductModel = ProductModel
                    .GroupBy(u => u.item)
                    .Select(grp => grp.ToList())
                    .ToList();
                var groupedProductCode = ProductModel
                    .GroupBy(u => u.product_code)
                    .Select(grp => grp.ToList())
                    .ToList();
                List<ExcelColumns> ProductModelSheetData = new List<ExcelColumns>();
                List<ExcelColumns> ProductCodeSheetData = new List<ExcelColumns>();
                foreach (var ProductModelItem in groupedProductModel)
                {
                    string sum_item = "";
                    string sum_fam_code = "";
                    string sum_trans_date = "";
                    decimal sum_qty_completed = 0;
                    decimal sum_produced_amt = 0;
                    decimal sum_actl_rm_cost = 0;
                    decimal sum_std_rm_cost = 0;
                    decimal sum_actl_lbr_cost = 0;
                    decimal sum_std_lbr_cost = 0;

                    for (int count = 0; count < ProductModelItem.Count; count++)
                    {
                        sum_item = ProductModelItem[count].item;
                        sum_fam_code = ProductModelItem[count].fam_code;
                        sum_trans_date = ProductModelItem[count].trans_date;
                        sum_qty_completed += ProductModelItem[count].qty_completed;
                        sum_produced_amt += ProductModelItem[count].produced_amt;
                        sum_actl_rm_cost += ProductModelItem[count].actl_rm_cost;
                        sum_std_rm_cost += ProductModelItem[count].std_rm_cost;
                        sum_actl_lbr_cost += ProductModelItem[count].actl_lbr_cost;
                        sum_std_lbr_cost += ProductModelItem[count].std_lbr_cost;
                    }
                    ProductModelSheetData.Add(new ExcelColumns
                    {
                        A = sum_item,
                        B = sum_fam_code,
                        C = sum_trans_date,
                        D = sum_qty_completed.ToString(),
                        E = sum_produced_amt.ToString(),
                        F = sum_produced_amt == 0 ? "0" : (sum_actl_rm_cost / sum_produced_amt).ToString(),
                        G = sum_actl_rm_cost.ToString(),
                        H = sum_std_rm_cost.ToString(),
                        I = sum_std_rm_cost == 0 ? "0" : (sum_actl_rm_cost / sum_std_rm_cost).ToString(),
                        J = sum_qty_completed == 0 ? "0" : (sum_actl_rm_cost / sum_qty_completed).ToString(),
                        K = sum_actl_lbr_cost.ToString(),
                        L = sum_std_lbr_cost.ToString(),
                        M = sum_std_lbr_cost == 0 ? "0" : (sum_actl_lbr_cost / sum_std_lbr_cost).ToString(),

                    });
                }
                foreach (var ProductCodeItem in groupedProductCode)
                {
                    string sum_product_code = "";
                    string sum_fam_code = "";
                    string sum_trans_date = "";
                    decimal sum_qty_completed = 0;
                    decimal sum_produced_amt = 0;
                    decimal sum_actl_rm_cost = 0;
                    decimal sum_std_rm_cost = 0;
                    decimal sum_actl_lbr_cost = 0;
                    decimal sum_std_lbr_cost = 0;

                    for (int count = 0; count < ProductCodeItem.Count; count++)
                    {
                        sum_product_code = ProductCodeItem[count].product_code;
                        sum_fam_code = ProductCodeItem[count].fam_code;
                        sum_trans_date = ProductCodeItem[count].trans_date;
                        sum_qty_completed += ProductCodeItem[count].qty_completed;
                        sum_produced_amt += ProductCodeItem[count].produced_amt;
                        sum_actl_rm_cost += ProductCodeItem[count].actl_rm_cost;
                        sum_std_rm_cost += ProductCodeItem[count].std_rm_cost;
                        sum_actl_lbr_cost += ProductCodeItem[count].actl_lbr_cost;
                        sum_std_lbr_cost += ProductCodeItem[count].std_lbr_cost;
                    }
                    ProductCodeSheetData.Add(new ExcelColumns
                    {
                        A = sum_product_code,
                        B = sum_fam_code,
                        C = sum_trans_date,
                        D = sum_qty_completed.ToString(),
                        E = sum_produced_amt.ToString(),
                        F = sum_produced_amt == 0 ? "0" : (sum_actl_rm_cost / sum_produced_amt).ToString(),
                        G = sum_actl_rm_cost.ToString(),
                        H = sum_std_rm_cost.ToString(),
                        I = sum_std_rm_cost == 0 ? "0" : (sum_actl_rm_cost / sum_std_rm_cost).ToString(),
                        J = sum_qty_completed == 0 ? "0" : (sum_actl_rm_cost / sum_qty_completed).ToString(),
                        K = sum_actl_lbr_cost.ToString(),
                        L = sum_std_lbr_cost.ToString(),
                        M = sum_std_lbr_cost == 0 ? "0" : (sum_actl_lbr_cost / sum_std_lbr_cost).ToString(),

                    });
                }
                string filePath = "";
                string Filename = "LSP_Rpt_NewDM_DirectMaterialLaborPercentageReport_" + MonthYear + ".xlsx";
                filePath = Path.Combine(Server.MapPath("~/Areas/Reports/Templates/") + "LSP_Rpt_NewDM_DirectMaterialLaborPercentageReport.xlsx");
                FileInfo file = new FileInfo(filePath);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    #region ProductModelSheet(Sheet1)
                    ExcelWorksheet ProductModelSheet = excelPackage.Workbook.Worksheets["ProductModel"];
                    int sheetrRow = 5;
                    foreach (var SheetData in ProductModelSheetData)
                    {
                        if (sheetrRow < ProductModelSheetData.Count + 4)
                        {
                            ProductModelSheet.InsertRow((sheetrRow + 1), 1);
                            ProductModelSheet.Cells[sheetrRow, 1, sheetrRow, 100].Copy(ProductModelSheet.Cells[(sheetrRow + 1), 1, (sheetrRow + 1), 1]);
                        }
                        ProductModelSheet.Cells[sheetrRow, 1].Value = SheetData.A;
                        ProductModelSheet.Cells[sheetrRow, 1].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 2].Value = SheetData.B;
                        ProductModelSheet.Cells[sheetrRow, 2].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 3].Value = SheetData.C;
                        ProductModelSheet.Cells[sheetrRow, 3].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 4].Value = Convert.ToDecimal(SheetData.D);
                        ProductModelSheet.Cells[sheetrRow, 4].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 5].Value = Convert.ToDecimal(SheetData.E);
                        ProductModelSheet.Cells[sheetrRow, 5].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 6].Value = Convert.ToDecimal(SheetData.F);
                        ProductModelSheet.Cells[sheetrRow, 6].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 7].Value = Convert.ToDecimal(SheetData.G);
                        ProductModelSheet.Cells[sheetrRow, 7].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 8].Value = Convert.ToDecimal(SheetData.H);
                        ProductModelSheet.Cells[sheetrRow, 8].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 9].Value = Convert.ToDecimal(SheetData.I);
                        ProductModelSheet.Cells[sheetrRow, 9].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 10].Value = Convert.ToDecimal(SheetData.J);
                        ProductModelSheet.Cells[sheetrRow, 10].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 11].Value = Convert.ToDecimal(SheetData.K);
                        ProductModelSheet.Cells[sheetrRow, 11].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 12].Value = Convert.ToDecimal(SheetData.L);
                        ProductModelSheet.Cells[sheetrRow, 12].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 13].Value = Convert.ToDecimal(SheetData.M);
                        ProductModelSheet.Cells[sheetrRow, 13].Style.WrapText = false;
                        sheetrRow++;
                    }
                    #endregion
                    #region ProductCodeSheet(Sheet1)
                    ExcelWorksheet ProductCodeSheet = excelPackage.Workbook.Worksheets["ProductCode"];
                    int sheetrRowCode = 5;
                    foreach (var SheetData in ProductCodeSheetData)
                    {
                        if (sheetrRowCode < ProductCodeSheetData.Count + 4)
                        {
                            ProductCodeSheet.InsertRow((sheetrRowCode + 1), 1);
                            ProductCodeSheet.Cells[sheetrRowCode, 1, sheetrRowCode, 100].Copy(ProductCodeSheet.Cells[(sheetrRowCode + 1), 1, (sheetrRowCode + 1), 1]);
                        }
                        ProductCodeSheet.Cells[sheetrRowCode, 1].Value = SheetData.A;
                        ProductCodeSheet.Cells[sheetrRowCode, 1].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 2].Value = SheetData.B;
                        ProductCodeSheet.Cells[sheetrRowCode, 2].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 3].Value = SheetData.C;
                        ProductCodeSheet.Cells[sheetrRowCode, 3].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 4].Value = Convert.ToDecimal(SheetData.D);
                        ProductCodeSheet.Cells[sheetrRowCode, 4].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 5].Value = Convert.ToDecimal(SheetData.E);
                        ProductCodeSheet.Cells[sheetrRowCode, 5].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 6].Value = Convert.ToDecimal(SheetData.F);
                        ProductCodeSheet.Cells[sheetrRowCode, 6].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 7].Value = Convert.ToDecimal(SheetData.G);
                        ProductCodeSheet.Cells[sheetrRowCode, 7].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 8].Value = Convert.ToDecimal(SheetData.H);
                        ProductCodeSheet.Cells[sheetrRowCode, 8].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 9].Value = Convert.ToDecimal(SheetData.I);
                        ProductCodeSheet.Cells[sheetrRowCode, 9].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 10].Value = Convert.ToDecimal(SheetData.J);
                        ProductCodeSheet.Cells[sheetrRowCode, 10].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 11].Value = Convert.ToDecimal(SheetData.K);
                        ProductCodeSheet.Cells[sheetrRowCode, 11].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 12].Value = Convert.ToDecimal(SheetData.L);
                        ProductCodeSheet.Cells[sheetrRowCode, 12].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 13].Value = Convert.ToDecimal(SheetData.M);
                        ProductCodeSheet.Cells[sheetrRowCode, 13].Style.WrapText = false;
                        sheetrRowCode++;
                    }
                    #endregion

                    return File(excelPackage.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Filename);
                }
            }
            catch (Exception err)
            {
                string errmsg;
                if (err.InnerException != null)
                    errmsg = "An error occured: " + err.InnerException.ToString();
                else
                    errmsg = "An error occured: " + err.Message.ToString();
                return null;
            }
        }
        public ActionResult GenerateMiscellaneousTransactionReport()
        {
            List<DMAndLaborPercentageReport> ProductModel = new List<DMAndLaborPercentageReport>();
            var FinishedGoodsAndSalesReport = Request["FinishedGoodsAndSalesReport"];
            var MiscellaneousTransactionReport = Request["MiscellaneousTransactionReport"];
            var DMAndLaborPercentageReport = Request["DMAndLaborPercentageReport"];
            var StartDate = Request["StartDate"];
            var EndDate = Request["EndDate"];
            var ProductCode1 = Request["ProductCode1"];
            var ProductCode2 = Request["ProductCode2"];
            var Model1 = Request["Model1"];
            var Model2 = Request["Model2"];
            var InventoryTurnoverReport = Request["InventoryTurnoverReport"];
            var ShowDetailedTransaction = Request["ShowDetailedTransaction"];

            string MonthYear = DateTime.Parse(StartDate).ToString("MMMyyyy");
            try
            {
                List<ExcelColumns> datalist = new List<ExcelColumns>();
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LSPI803_App"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {

                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "LSP_Rpt_NewDM_DirectMaterialLaborPercentageReportSp";
                        cmdSql.CommandTimeout = 0;
                        cmdSql.Parameters.Clear();

                        cmdSql.Parameters.AddWithValue("@StartDate", StartDate);
                        cmdSql.Parameters.AddWithValue("@EndDate", EndDate);
                        cmdSql.Parameters.AddWithValue("@StartProdCode", ProductCode1 == null ? "" : ProductCode1);
                        cmdSql.Parameters.AddWithValue("@EndProdCode", ProductCode2 == null ? "" : ProductCode2);
                        cmdSql.Parameters.AddWithValue("@StartModel", Model1 == null ? "" : Model1);
                        cmdSql.Parameters.AddWithValue("@EndModel", Model2 == null ? "" : Model2);
                        cmdSql.ExecuteNonQuery();
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                ProductModel.Add(new DMAndLaborPercentageReport
                                {
                                    item = sdr["item"].ToString().Remove(0, 3),
                                    product_code = sdr["product_code"].ToString().Remove(0, 3),
                                    fam_code = sdr["fam_code"].ToString(),
                                    trans_date = DateTime.Parse(sdr["trans_date"].ToString()).ToString("MMM-yyyy"),
                                    qty_completed = Convert.ToDecimal(sdr["qty_completed"]),
                                    produced_amt = Convert.ToDecimal(sdr["produced_amt"]),
                                    actl_rm_cost = Convert.ToDecimal(sdr["actl_rm_cost"]),
                                    std_rm_cost = Convert.ToDecimal(sdr["std_rm_cost"]),
                                    actl_lbr_cost = Convert.ToDecimal(sdr["actl_lbr_cost"]),
                                    std_lbr_cost = Convert.ToDecimal(sdr["std_lbr_cost"]),
                                });
                            }

                        }
                    }
                    conn.Close();
                }


                var groupedProductModel = ProductModel
                    .GroupBy(u => u.item)
                    .Select(grp => grp.ToList())
                    .ToList();
                var groupedProductCode = ProductModel
                    .GroupBy(u => u.product_code)
                    .Select(grp => grp.ToList())
                    .ToList();
                List<ExcelColumns> ProductModelSheetData = new List<ExcelColumns>();
                List<ExcelColumns> ProductCodeSheetData = new List<ExcelColumns>();
                foreach (var ProductModelItem in groupedProductModel)
                {
                    string sum_item = "";
                    string sum_fam_code = "";
                    string sum_trans_date = "";
                    decimal sum_qty_completed = 0;
                    decimal sum_produced_amt = 0;
                    decimal sum_actl_rm_cost = 0;
                    decimal sum_std_rm_cost = 0;
                    decimal sum_actl_lbr_cost = 0;
                    decimal sum_std_lbr_cost = 0;

                    for (int count = 0; count < ProductModelItem.Count; count++)
                    {
                        sum_item = ProductModelItem[count].item;
                        sum_fam_code = ProductModelItem[count].fam_code;
                        sum_trans_date = ProductModelItem[count].trans_date;
                        sum_qty_completed += ProductModelItem[count].qty_completed;
                        sum_produced_amt += ProductModelItem[count].produced_amt;
                        sum_actl_rm_cost += ProductModelItem[count].actl_rm_cost;
                        sum_std_rm_cost += ProductModelItem[count].std_rm_cost;
                        sum_actl_lbr_cost += ProductModelItem[count].actl_lbr_cost;
                        sum_std_lbr_cost += ProductModelItem[count].std_lbr_cost;
                    }
                    ProductModelSheetData.Add(new ExcelColumns
                    {
                        A = sum_item,
                        B = sum_fam_code,
                        C = sum_trans_date,
                        D = sum_qty_completed.ToString(),
                        E = sum_produced_amt.ToString(),
                        F = sum_produced_amt == 0 ? "0" : (sum_actl_rm_cost / sum_produced_amt).ToString(),
                        G = sum_actl_rm_cost.ToString(),
                        H = sum_std_rm_cost.ToString(),
                        I = sum_std_rm_cost == 0 ? "0" : (sum_actl_rm_cost / sum_std_rm_cost).ToString(),
                        J = sum_qty_completed == 0 ? "0" : (sum_actl_rm_cost / sum_qty_completed).ToString(),
                        K = sum_actl_lbr_cost.ToString(),
                        L = sum_std_lbr_cost.ToString(),
                        M = sum_std_lbr_cost == 0 ? "0" : (sum_actl_lbr_cost / sum_std_lbr_cost).ToString(),

                    });
                }
                foreach (var ProductCodeItem in groupedProductCode)
                {
                    string sum_product_code = "";
                    string sum_fam_code = "";
                    string sum_trans_date = "";
                    decimal sum_qty_completed = 0;
                    decimal sum_produced_amt = 0;
                    decimal sum_actl_rm_cost = 0;
                    decimal sum_std_rm_cost = 0;
                    decimal sum_actl_lbr_cost = 0;
                    decimal sum_std_lbr_cost = 0;

                    for (int count = 0; count < ProductCodeItem.Count; count++)
                    {
                        sum_product_code = ProductCodeItem[count].product_code;
                        sum_fam_code = ProductCodeItem[count].fam_code;
                        sum_trans_date = ProductCodeItem[count].trans_date;
                        sum_qty_completed += ProductCodeItem[count].qty_completed;
                        sum_produced_amt += ProductCodeItem[count].produced_amt;
                        sum_actl_rm_cost += ProductCodeItem[count].actl_rm_cost;
                        sum_std_rm_cost += ProductCodeItem[count].std_rm_cost;
                        sum_actl_lbr_cost += ProductCodeItem[count].actl_lbr_cost;
                        sum_std_lbr_cost += ProductCodeItem[count].std_lbr_cost;
                    }
                    ProductCodeSheetData.Add(new ExcelColumns
                    {
                        A = sum_product_code,
                        B = sum_fam_code,
                        C = sum_trans_date,
                        D = sum_qty_completed.ToString(),
                        E = sum_produced_amt.ToString(),
                        F = sum_produced_amt == 0 ? "0" : (sum_actl_rm_cost / sum_produced_amt).ToString(),
                        G = sum_actl_rm_cost.ToString(),
                        H = sum_std_rm_cost.ToString(),
                        I = sum_std_rm_cost == 0 ? "0" : (sum_actl_rm_cost / sum_std_rm_cost).ToString(),
                        J = sum_qty_completed == 0 ? "0" : (sum_actl_rm_cost / sum_qty_completed).ToString(),
                        K = sum_actl_lbr_cost.ToString(),
                        L = sum_std_lbr_cost.ToString(),
                        M = sum_std_lbr_cost == 0 ? "0" : (sum_actl_lbr_cost / sum_std_lbr_cost).ToString(),

                    });
                }
                string filePath = "";
                string Filename = "LSP_Rpt_NewGenerateMiscellaneousTransactionReport_" + MonthYear + ".xlsx";
                filePath = Path.Combine(Server.MapPath("~/Areas/Reports/Templates/") + "LSP_Rpt_NewDM_DirectMaterialLaborPercentageReport.xlsx");
                FileInfo file = new FileInfo(filePath);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    #region ProductModelSheet(Sheet1)
                    ExcelWorksheet ProductModelSheet = excelPackage.Workbook.Worksheets["ProductModel"];
                    int sheetrRow = 5;
                    foreach (var SheetData in ProductModelSheetData)
                    {
                        if (sheetrRow < ProductModelSheetData.Count + 4)
                        {
                            ProductModelSheet.InsertRow((sheetrRow + 1), 1);
                            ProductModelSheet.Cells[sheetrRow, 1, sheetrRow, 100].Copy(ProductModelSheet.Cells[(sheetrRow + 1), 1, (sheetrRow + 1), 1]);
                        }
                        ProductModelSheet.Cells[sheetrRow, 1].Value = SheetData.A;
                        ProductModelSheet.Cells[sheetrRow, 1].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 2].Value = SheetData.B;
                        ProductModelSheet.Cells[sheetrRow, 2].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 3].Value = SheetData.C;
                        ProductModelSheet.Cells[sheetrRow, 3].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 4].Value = Convert.ToDecimal(SheetData.D);
                        ProductModelSheet.Cells[sheetrRow, 4].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 5].Value = Convert.ToDecimal(SheetData.E);
                        ProductModelSheet.Cells[sheetrRow, 5].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 6].Value = Convert.ToDecimal(SheetData.F);
                        ProductModelSheet.Cells[sheetrRow, 6].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 7].Value = Convert.ToDecimal(SheetData.G);
                        ProductModelSheet.Cells[sheetrRow, 7].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 8].Value = Convert.ToDecimal(SheetData.H);
                        ProductModelSheet.Cells[sheetrRow, 8].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 9].Value = Convert.ToDecimal(SheetData.I);
                        ProductModelSheet.Cells[sheetrRow, 9].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 10].Value = Convert.ToDecimal(SheetData.J);
                        ProductModelSheet.Cells[sheetrRow, 10].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 11].Value = Convert.ToDecimal(SheetData.K);
                        ProductModelSheet.Cells[sheetrRow, 11].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 12].Value = Convert.ToDecimal(SheetData.L);
                        ProductModelSheet.Cells[sheetrRow, 12].Style.WrapText = false;
                        ProductModelSheet.Cells[sheetrRow, 13].Value = Convert.ToDecimal(SheetData.M);
                        ProductModelSheet.Cells[sheetrRow, 13].Style.WrapText = false;
                        sheetrRow++;
                    }
                    #endregion
                    #region ProductCodeSheet(Sheet1)
                    ExcelWorksheet ProductCodeSheet = excelPackage.Workbook.Worksheets["ProductCode"];
                    int sheetrRowCode = 5;
                    foreach (var SheetData in ProductCodeSheetData)
                    {
                        if (sheetrRowCode < ProductCodeSheetData.Count + 4)
                        {
                            ProductCodeSheet.InsertRow((sheetrRowCode + 1), 1);
                            ProductCodeSheet.Cells[sheetrRowCode, 1, sheetrRowCode, 100].Copy(ProductCodeSheet.Cells[(sheetrRowCode + 1), 1, (sheetrRowCode + 1), 1]);
                        }
                        ProductCodeSheet.Cells[sheetrRowCode, 1].Value = SheetData.A;
                        ProductCodeSheet.Cells[sheetrRowCode, 1].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 2].Value = SheetData.B;
                        ProductCodeSheet.Cells[sheetrRowCode, 2].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 3].Value = SheetData.C;
                        ProductCodeSheet.Cells[sheetrRowCode, 3].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 4].Value = Convert.ToDecimal(SheetData.D);
                        ProductCodeSheet.Cells[sheetrRowCode, 4].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 5].Value = Convert.ToDecimal(SheetData.E);
                        ProductCodeSheet.Cells[sheetrRowCode, 5].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 6].Value = Convert.ToDecimal(SheetData.F);
                        ProductCodeSheet.Cells[sheetrRowCode, 6].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 7].Value = Convert.ToDecimal(SheetData.G);
                        ProductCodeSheet.Cells[sheetrRowCode, 7].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 8].Value = Convert.ToDecimal(SheetData.H);
                        ProductCodeSheet.Cells[sheetrRowCode, 8].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 9].Value = Convert.ToDecimal(SheetData.I);
                        ProductCodeSheet.Cells[sheetrRowCode, 9].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 10].Value = Convert.ToDecimal(SheetData.J);
                        ProductCodeSheet.Cells[sheetrRowCode, 10].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 11].Value = Convert.ToDecimal(SheetData.K);
                        ProductCodeSheet.Cells[sheetrRowCode, 11].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 12].Value = Convert.ToDecimal(SheetData.L);
                        ProductCodeSheet.Cells[sheetrRowCode, 12].Style.WrapText = false;
                        ProductCodeSheet.Cells[sheetrRowCode, 13].Value = Convert.ToDecimal(SheetData.M);
                        ProductCodeSheet.Cells[sheetrRowCode, 13].Style.WrapText = false;
                        sheetrRowCode++;
                    }
                    #endregion

                    return File(excelPackage.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Filename);
                }
            }
            catch (Exception err)
            {
                string errmsg;
                if (err.InnerException != null)
                    errmsg = "An error occured: " + err.InnerException.ToString();
                else
                    errmsg = "An error occured: " + err.Message.ToString();
                return null;
            }
        }
    }
}