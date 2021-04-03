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
using OfficeOpenXml.Style;

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
        public ActionResult SlowMonitoringAnalysisReport()
        {
            List<ExcelColumns> SlowMonitoringAnalysis = new List<ExcelColumns>();
            var Month = Request["Month"];

            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LSPI803_App"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {

                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "LSP_Rpt_NewDM_SlowMovingAnalysisReportSp";
                        cmdSql.CommandTimeout = 0;
                        cmdSql.Parameters.Clear();
                        cmdSql.Parameters.AddWithValue("@Months", Month);
                        cmdSql.ExecuteNonQuery();
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                SlowMonitoringAnalysis.Add(new ExcelColumns
                                {
                                    A = sdr["item"].ToString(),
                                    B = sdr["description"].ToString(),
                                    C = sdr["product_code"].ToString(),
                                    D = sdr["Uf_location"].ToString(),
                                    E = sdr["matl_stat"].ToString(),
                                    F = Convert.ToDecimal(sdr["QtyOnHand"]).ToString(),
                                    G = Convert.ToDecimal(sdr["TotalMatlCostPHP"]).ToString(),
                                    H = Convert.ToDecimal(sdr["TotalLandedCostPHP"]).ToString(),
                                    I = Convert.ToDecimal(sdr["TotalPIFGProcessCostPHP"]).ToString(),
                                    J = Convert.ToDecimal(sdr["TotalPIResinCostPHP"]).ToString(),
                                    K = Convert.ToDecimal(sdr["TotalPIHiddenPHP"]).ToString(),
                                    L = Convert.ToDecimal(sdr["TotalSFLbrCostPHP"]).ToString(),
                                    M = Convert.ToDecimal(sdr["TotalCostPHP"]).ToString(),
                                    N = sdr["LatestPODate"].ToString() == "" ? "" : DateTime.Parse(sdr["LatestPODate"].ToString()).ToString("MM/dd/yyyy"),
                                    O = sdr["LatestIssueDate"].ToString() == "" ? "" : DateTime.Parse(sdr["LatestIssueDate"].ToString()).ToString("MM/dd/yyyy"),
                                    P = sdr["ItemRemarks"].ToString()
                                });
                            }

                        }
                    }
                    conn.Close();
                }

                string filePath = "";
                string Filename = "LSP_Rpt_DM_SlowMovingAnalysisReport.xlsx";
                filePath = Path.Combine(Server.MapPath("~/Areas/Reports/Templates/") + "LSP_Rpt_DM_SlowMovingAnalysisReport.xlsx");
                FileInfo file = new FileInfo(filePath);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    #region ProductModelSheet(Sheet1)
                    ExcelWorksheet SlowMovingAnalysisReportSheet = excelPackage.Workbook.Worksheets["SlowMovingAnalysisReport"];
                    int sheetrRow = 4;
                    foreach (var SheetData in SlowMonitoringAnalysis)
                    {
                        if (sheetrRow < SlowMonitoringAnalysis.Count + 3)
                        {
                            SlowMovingAnalysisReportSheet.InsertRow((sheetrRow + 1), 1);
                            SlowMovingAnalysisReportSheet.Cells[sheetrRow, 1, sheetrRow, 100].Copy(SlowMovingAnalysisReportSheet.Cells[(sheetrRow + 1), 1, (sheetrRow + 1), 1]);
                        }
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 1].Value = SheetData.A;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 1].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 2].Value = SheetData.B;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 2].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 3].Value = SheetData.C;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 3].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 4].Value = SheetData.D;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 4].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 5].Value = SheetData.E;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 5].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 6].Value = Convert.ToDecimal(SheetData.F);
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 6].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 7].Value = Convert.ToDecimal(SheetData.G);
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 7].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 8].Value = Convert.ToDecimal(SheetData.H);
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 8].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 9].Value = Convert.ToDecimal(SheetData.I);
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 9].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 10].Value = Convert.ToDecimal(SheetData.J);
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 10].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 11].Value = Convert.ToDecimal(SheetData.K);
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 11].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 12].Value = Convert.ToDecimal(SheetData.L);
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 12].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 13].Value = Convert.ToDecimal(SheetData.M);
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 13].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 14].Value = SheetData.N;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 14].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 15].Value = SheetData.O;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 15].Style.WrapText = false;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 16].Value = SheetData.P;
                        SlowMovingAnalysisReportSheet.Cells[sheetrRow, 16].Style.WrapText = false;
                        sheetrRow++;
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
        public ActionResult GenerateWIPShopFloorReport()
        {
            List<ExcelColumns> WIPShopFloorReport = new List<ExcelColumns>();
            decimal Total_WIPQty = 0;
            decimal Total_TotalWIPCost_PHP = 0;
            decimal Total_TotalWIPCost_USD = 0;
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LSPI803_App"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {

                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "LSP_Rpt_NewDM_WIPShopFloorReportSp";
                        cmdSql.CommandTimeout = 0;
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                Total_WIPQty += Convert.ToDecimal(sdr["WIPQty"]);
                                Total_TotalWIPCost_PHP += Convert.ToDecimal(sdr["TotalWIPCost_PHP"]);
                                Total_TotalWIPCost_USD += Convert.ToDecimal(sdr["TotalWIPCost_USD"]);
                                WIPShopFloorReport.Add(new ExcelColumns
                                {
                                    A = sdr["TransDate"].ToString() == "" ? "" : DateTime.Parse(sdr["TransDate"].ToString()).ToString("MM/dd/yyyy"),
                                    B = sdr["Item"].ToString(),
                                    C = sdr["ItemDesc"].ToString(),
                                    D = sdr["JOReference"].ToString(),
                                    E = sdr["WIPQty"].ToString(),
                                    F = sdr["MatlUnit_PHP"].ToString(),
                                    G = sdr["LandedUnit_PHP"].ToString(),
                                    H = sdr["PIFGProcessUnit_PHP"].ToString(),
                                    I = sdr["PIResinUnit_PHP"].ToString(),
                                    J = sdr["PIHiddenUnit_PHP"].ToString(),
                                    K = sdr["SFAddedUnit_PHP"].ToString(),
                                    L = sdr["FGAddedUnit_PHP"].ToString(),
                                    M = sdr["TotalWIPCost_PHP"].ToString(),
                                    N = sdr["MatlUnit_USD"].ToString(),
                                    O = sdr["LandedUnit_USD"].ToString(),
                                    P = sdr["PIFGProcessUnit_USD"].ToString(),
                                    Q = sdr["PIResinUnit_USD"].ToString(),
                                    R = sdr["PIHiddenUnit_USD"].ToString(),
                                    S = sdr["SFAddedUnit_USD"].ToString(),
                                    T = sdr["FGAddedUnit_USD"].ToString(),
                                    U = sdr["TotalWIPCost_USD"].ToString(),
                                });
                            }

                        }
                    }
                    conn.Close();
                }

                string filePath = "";
                string Filename = "LSP_Rpt_DM_WIPShopFloorReport.xlsx";
                filePath = Path.Combine(Server.MapPath("~/Areas/Reports/Templates/") + "LSP_Rpt_DM_WIPShopFloorReport.xlsx");
                FileInfo file = new FileInfo(filePath);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    #region WIPShopFloorReport(Sheet1)
                    ExcelWorksheet WIPShopFloorReportSheet = excelPackage.Workbook.Worksheets["LSP_Rpt_DM_WIPShopFloorReport"];
                    int sheetrRow = 5;
                    foreach (var SheetData in WIPShopFloorReport)
                    {
                        if (sheetrRow < WIPShopFloorReport.Count + 4)
                        {
                            WIPShopFloorReportSheet.InsertRow((sheetrRow + 1), 1);
                            WIPShopFloorReportSheet.Cells[sheetrRow, 1, sheetrRow, 100].Copy(WIPShopFloorReportSheet.Cells[(sheetrRow + 1), 1, (sheetrRow + 1), 1]);
                        }
                        WIPShopFloorReportSheet.Cells[sheetrRow, 1].Value = SheetData.A;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 1].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 2].Value = SheetData.B;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 2].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 3].Value = SheetData.C;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 3].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 4].Value = SheetData.D;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 4].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 5].Value = Convert.ToDecimal(SheetData.E);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 5].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 6].Value = Convert.ToDecimal(SheetData.F);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 6].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 7].Value = Convert.ToDecimal(SheetData.G);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 7].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 8].Value = Convert.ToDecimal(SheetData.H);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 8].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 9].Value = Convert.ToDecimal(SheetData.I);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 9].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 10].Value = Convert.ToDecimal(SheetData.J);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 10].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 11].Value = Convert.ToDecimal(SheetData.K);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 11].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 12].Value = Convert.ToDecimal(SheetData.L);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 12].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 13].Value = Convert.ToDecimal(SheetData.M);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 13].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 14].Value = Convert.ToDecimal(SheetData.N);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 14].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 15].Value = Convert.ToDecimal(SheetData.O);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 15].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 16].Value = Convert.ToDecimal(SheetData.P);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 16].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 17].Value = Convert.ToDecimal(SheetData.Q);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 17].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 18].Value = Convert.ToDecimal(SheetData.R);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 18].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 19].Value = Convert.ToDecimal(SheetData.S);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 19].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 20].Value = Convert.ToDecimal(SheetData.T);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 20].Style.WrapText = false;
                        WIPShopFloorReportSheet.Cells[sheetrRow, 21].Value = Convert.ToDecimal(SheetData.U);
                        WIPShopFloorReportSheet.Cells[sheetrRow, 21].Style.WrapText = false;
                        sheetrRow++;
                    }

                    WIPShopFloorReportSheet.Cells[sheetrRow, 5].Value = Convert.ToDecimal(Total_WIPQty);
                    WIPShopFloorReportSheet.Cells[sheetrRow, 5].Style.WrapText = false;
                    WIPShopFloorReportSheet.Cells[sheetrRow, 13].Value = Convert.ToDecimal(Total_TotalWIPCost_PHP);
                    WIPShopFloorReportSheet.Cells[sheetrRow, 13].Style.WrapText = false;
                    WIPShopFloorReportSheet.Cells[sheetrRow, 21].Value = Convert.ToDecimal(Total_TotalWIPCost_USD);
                    WIPShopFloorReportSheet.Cells[sheetrRow, 21].Style.WrapText = false;
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
            List<MiscellaneousTransactionReport> MiscellaneousTransaction = new List<MiscellaneousTransactionReport>();
            var StartDate = Request["StartDate"];
            var EndDate = Request["EndDate"];

            string MonthYear = DateTime.Parse(StartDate).ToString("MMMdd_yyyy") + "to" + DateTime.Parse(EndDate).ToString("MMMdd_yyyy");
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LSPI803_App"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {

                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "LSP_Rpt_NewDM_MiscellaneousTransactionReportSp";
                        cmdSql.CommandTimeout = 0;
                        cmdSql.Parameters.Clear();

                        cmdSql.Parameters.AddWithValue("@StartDate", StartDate);
                        cmdSql.Parameters.AddWithValue("@EndDate", EndDate);
                        //cmdSql.ExecuteNonQuery();
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                MiscellaneousTransaction.Add(new MiscellaneousTransactionReport
                                {
                                    SummaryGroup = sdr["SummaryGroup"].ToString(),
                                    TransType = sdr["TransType"].ToString(),
                                    TransDesc = sdr["TransDesc"].ToString(),
                                    ReasonDesc = sdr["ReasonDesc"].ToString(),
                                    MiscTransClass = sdr["MiscTransClass"].ToString(),
                                    JobOrLot = sdr["JobOrLot"].ToString(),
                                    TransDate = sdr["TransDate"].ToString(),
                                    Item = sdr["Item"].ToString(),
                                    ItemDesc = sdr["ItemDesc"].ToString(),
                                    QtyCompleted = sdr["QtyCompleted"].ToString() == "" ? 0 : Convert.ToInt32(sdr["QtyCompleted"]),
                                    QtyScrapped = sdr["QtyScrapped"].ToString() == "" ? 0 : Convert.ToInt32(sdr["QtyScrapped"]),
                                    Employee = sdr["Employee"].ToString() == "" ? 0 : Convert.ToInt32(sdr["Employee"]),
                                    Wc = sdr["Wc"].ToString(),
                                    MatlCost_PHP = sdr["MatlCost_PHP"].ToString() == "" ? 0 : Convert.ToDecimal(sdr["MatlCost_PHP"]),
                                    MatlLandedCost_PHP = sdr["MatlLandedCost_PHP"].ToString() == "" ? 0 : Convert.ToDecimal(sdr["MatlLandedCost_PHP"]),
                                    PIResin_PHP = sdr["PIResin_PHP"].ToString() == "" ? 0 : Convert.ToDecimal(sdr["PIResin_PHP"]),
                                    PIFGProcess_PHP = sdr["PIFGProcess_PHP"].ToString() == "" ? 0 : Convert.ToDecimal(sdr["PIFGProcess_PHP"]),
                                    PIHiddenProfit_PHP = sdr["PIHiddenProfit_PHP"].ToString() == "" ? 0 : Convert.ToDecimal(sdr["PIHiddenProfit_PHP"]),
                                    SFAddedCost_PHP = sdr["SFAddedCost_PHP"].ToString() == "" ? 0 : Convert.ToDecimal(sdr["SFAddedCost_PHP"]),
                                    FGAddedCost_PHP = sdr["FGAddedCost_PHP"].ToString() == "" ? 0 : Convert.ToDecimal(sdr["FGAddedCost_PHP"]),
                                    TotalCost_PHP = sdr["TotalCost_PHP"].ToString() == "" ? 0 : Convert.ToDecimal(sdr["TotalCost_PHP"]),
                                    TransQty = sdr["TransQty"].ToString() == "" ? 0 : Convert.ToDecimal(sdr["TransQty"]),
                                });
                            }

                        }
                    }
                    conn.Close();
                }

                var filteredMiscellaneousTransaction = MiscellaneousTransaction.Where(x => x.TransDesc == "SF Scrap Data").ToList();
                

                string filePath = "";
                string Filename = "LSP_Rpt_DM_MiscellaneousTransactionReport_LSPI_" + MonthYear + ".xlsx";
                filePath = Path.Combine(Server.MapPath("~/Areas/Reports/Templates/") + "LSP_Rpt_DM_MiscellaneousTransactionReport_LSPI.xlsx");
                FileInfo file = new FileInfo(filePath);
                
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {

                    #region MiscellaneousTransactions
                    var Summary_GroupByTransDesc = MiscellaneousTransaction
                        .Where(x=>x.TransDesc != "SF Scrap Data")
                        .GroupBy(u => u.TransDesc)
                        .ToList();

                    decimal Total_TransQty =0;
                    decimal Total_TotalCost_PHP =0;
                    int sheetRowMisc = 6;
                    foreach (var TransDescList in Summary_GroupByTransDesc)
                    {
                        ExcelWorksheet MiscTrxSheetCycleCount = excelPackage.Workbook.Worksheets["Cycle Count"];
                        ExcelWorksheet MiscTrxSheetMiscellaneousIssue = excelPackage.Workbook.Worksheets["Miscellaneous Issue"];
                        ExcelWorksheet MiscTrxSheetMiscellaneousReceipt = excelPackage.Workbook.Worksheets["Miscellaneous Receipt"];
                        if (TransDescList.Key.ToString().Trim() == "Cycle Count")
                        {

                            Total_TransQty = 0;
                            Total_TotalCost_PHP = 0;
                            sheetRowMisc = 6;
                            foreach (var SheetData in TransDescList)
                            {
                                if (sheetRowMisc < TransDescList.ToList().Count + 5)
                                {
                                    MiscTrxSheetCycleCount.InsertRow((sheetRowMisc + 1), 1);
                                    MiscTrxSheetCycleCount.Cells[sheetRowMisc, 1, sheetRowMisc, 100].Copy(MiscTrxSheetCycleCount.Cells[(sheetRowMisc + 1), 1, (sheetRowMisc + 1), 1]);
                                }
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 1].Value = DateTime.Parse(SheetData.TransDate).ToString("MM/dd/yyyy");
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 1].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 2].Value = SheetData.Item;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 2].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 3].Value = SheetData.ItemDesc;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 3].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 4].Value = SheetData.TransDesc;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 4].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 5].Value = SheetData.ReasonDesc;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 5].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 6].Value = Convert.ToDecimal(SheetData.TransQty);
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 6].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 7].Value = Convert.ToDecimal(SheetData.MatlCost_PHP);
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 7].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 8].Value = Convert.ToDecimal(SheetData.MatlLandedCost_PHP);
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 8].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 9].Value = Convert.ToDecimal(SheetData.PIResin_PHP);
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 9].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 10].Value = Convert.ToDecimal(SheetData.PIFGProcess_PHP);
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 10].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 11].Value = Convert.ToDecimal(SheetData.PIHiddenProfit_PHP);
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 11].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 12].Value = Convert.ToDecimal(SheetData.SFAddedCost_PHP)+ Convert.ToDecimal(SheetData.FGAddedCost_PHP);
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 12].Style.WrapText = false;
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 13].Value = Convert.ToDecimal(SheetData.TotalCost_PHP);
                                MiscTrxSheetCycleCount.Cells[sheetRowMisc, 13].Style.WrapText = false;
                                Total_TransQty += Convert.ToDecimal(SheetData.TransQty);
                                Total_TotalCost_PHP += Convert.ToDecimal(SheetData.TotalCost_PHP);
                                sheetRowMisc++;
                            }

                            MiscTrxSheetCycleCount.Cells[sheetRowMisc, 6].Value = Convert.ToDecimal(Total_TransQty);
                            MiscTrxSheetCycleCount.Cells[sheetRowMisc, 6].Style.WrapText = false;
                            MiscTrxSheetCycleCount.Cells[sheetRowMisc, 13].Value = Convert.ToDecimal(Total_TotalCost_PHP);
                            MiscTrxSheetCycleCount.Cells[sheetRowMisc, 13].Style.WrapText = false;
                        }
 
                        else if (TransDescList.Key.ToString().Trim() == "Miscellaneous Issue")
                        {
                            Total_TransQty = 0;
                            Total_TotalCost_PHP = 0;
                            sheetRowMisc = 6;
                            var IssueTransDescList = TransDescList.Where(x=>x.Item!="Scrap Item" && x.Item!="Request Item").ToList();
                            foreach (var SheetData in IssueTransDescList)
                            {
                                if (sheetRowMisc < TransDescList.ToList().Count + 5)
                                {
                                    MiscTrxSheetMiscellaneousIssue.InsertRow((sheetRowMisc + 1), 1);
                                    MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 1, sheetRowMisc, 100].Copy(MiscTrxSheetMiscellaneousIssue.Cells[(sheetRowMisc + 1), 1, (sheetRowMisc + 1), 1]);
                                }
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 1].Value = DateTime.Parse(SheetData.TransDate).ToString("MM/dd/yyyy");
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 1].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 2].Value = SheetData.Item;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 2].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 3].Value = SheetData.ItemDesc;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 3].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 4].Value = SheetData.TransDesc;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 4].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 5].Value = SheetData.ReasonDesc;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 5].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 6].Value = Convert.ToDecimal(SheetData.TransQty);
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 6].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 7].Value = Convert.ToDecimal(SheetData.MatlCost_PHP);
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 7].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 8].Value = Convert.ToDecimal(SheetData.MatlLandedCost_PHP);
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 8].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 9].Value = Convert.ToDecimal(SheetData.PIResin_PHP);
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 9].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 10].Value = Convert.ToDecimal(SheetData.PIFGProcess_PHP);
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 10].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 11].Value = Convert.ToDecimal(SheetData.PIHiddenProfit_PHP);
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 11].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 12].Value = Convert.ToDecimal(SheetData.SFAddedCost_PHP) + Convert.ToDecimal(SheetData.FGAddedCost_PHP);
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 12].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 13].Value = Convert.ToDecimal(SheetData.TotalCost_PHP);
                                MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 13].Style.WrapText = false;
                                Total_TransQty += Convert.ToDecimal(SheetData.TransQty);
                                Total_TotalCost_PHP += Convert.ToDecimal(SheetData.TotalCost_PHP);
                                sheetRowMisc++;
                            }

                            MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 6].Value = Convert.ToDecimal(Total_TransQty);
                            MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 6].Style.WrapText = false;
                            MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 13].Value = Convert.ToDecimal(Total_TotalCost_PHP);
                            MiscTrxSheetMiscellaneousIssue.Cells[sheetRowMisc, 13].Style.WrapText = false;
                        }
                        else
                        {
                            Total_TransQty = 0;
                            Total_TotalCost_PHP = 0;
                            sheetRowMisc = 6;
                            foreach (var SheetData in TransDescList)
                            {
                                if (sheetRowMisc < TransDescList.ToList().Count + 5)
                                {
                                    MiscTrxSheetMiscellaneousReceipt.InsertRow((sheetRowMisc + 1), 1);
                                    MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 1, sheetRowMisc, 100].Copy(MiscTrxSheetMiscellaneousReceipt.Cells[(sheetRowMisc + 1), 1, (sheetRowMisc + 1), 1]);
                                }
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 1].Value = DateTime.Parse(SheetData.TransDate).ToString("MM/dd/yyyy");
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 1].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 2].Value = SheetData.Item;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 2].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 3].Value = SheetData.ItemDesc;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 3].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 4].Value = SheetData.TransDesc;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 4].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 5].Value = SheetData.ReasonDesc;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 5].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 6].Value = Convert.ToDecimal(SheetData.TransQty);
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 6].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 7].Value = Convert.ToDecimal(SheetData.MatlCost_PHP);
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 7].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 8].Value = Convert.ToDecimal(SheetData.MatlLandedCost_PHP);
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 8].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 9].Value = Convert.ToDecimal(SheetData.PIResin_PHP);
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 9].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 10].Value = Convert.ToDecimal(SheetData.PIFGProcess_PHP);
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 10].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 11].Value = Convert.ToDecimal(SheetData.PIHiddenProfit_PHP);
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 11].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 12].Value = Convert.ToDecimal(SheetData.SFAddedCost_PHP) + Convert.ToDecimal(SheetData.FGAddedCost_PHP);
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 12].Style.WrapText = false;
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 13].Value = Convert.ToDecimal(SheetData.TotalCost_PHP);
                                MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 13].Style.WrapText = false;
                                Total_TransQty += Convert.ToDecimal(SheetData.TransQty);
                                Total_TotalCost_PHP += Convert.ToDecimal(SheetData.TotalCost_PHP);
                                sheetRowMisc++;
                            }

                            MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 6].Value = Convert.ToDecimal(Total_TransQty);
                            MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 6].Style.WrapText = false;
                            MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 13].Value = Convert.ToDecimal(Total_TotalCost_PHP);
                            MiscTrxSheetMiscellaneousReceipt.Cells[sheetRowMisc, 13].Style.WrapText = false;
                        }

                        
                    }
                    var FoundCycleCount = false;
                    var FoundMiscellaneousIssue = false;
                    var FoundMiscellaneousReceipt = false;
                    foreach (var TransDescList in Summary_GroupByTransDesc)
                    {
                        if (TransDescList.Key.ToString().Trim() == "Cycle Count")
                        {
                            FoundCycleCount = true;
                        }
                        else if (TransDescList.Key.ToString().Trim() == "Miscellaneous Issue")
                        {
                            FoundMiscellaneousIssue = true;
                        }
                        else if (TransDescList.Key.ToString().Trim() == "Miscellaneous Receipt")
                        {
                            FoundMiscellaneousReceipt = true;
                        }
                    }

                    if (FoundCycleCount == false)
                        excelPackage.Workbook.Worksheets.Delete("Cycle Count");
                    if (FoundMiscellaneousIssue == false)
                        excelPackage.Workbook.Worksheets.Delete("Miscellaneous Issue");
                    if (FoundMiscellaneousReceipt == false)
                        excelPackage.Workbook.Worksheets.Delete("Miscellaneous Receipt");

                    #endregion
                    #region Summary(1st Sheet)

                    ExcelWorksheet Summary1stSheet = excelPackage.Workbook.Worksheets["Summary"];
                    var Summary_GroupBySummaryGroup = MiscellaneousTransaction
                        .OrderBy(x=>x.TransType)
                        .ThenBy(x => x.MiscTransClass)
                        .ThenBy(x => x.ReasonDesc)
                        .GroupBy(u => u.SummaryGroup)
                        .ToList();
                    int summaryGroupCtr = 0;
                    int summary1stSheetRow = 5;
                    foreach (var SheetData in Summary_GroupBySummaryGroup)
                    {
                        var SummaryGroupByReasonDesc = SheetData
                                .OrderBy(x => x.TransType).ThenBy(x => x.MiscTransClass).ThenBy(x => x.ReasonDesc)
                                .GroupBy(u => u.ReasonDesc)
                                .ToList();
                        Summary1stSheet.Cells[summary1stSheetRow-1, 1].Value = SheetData.Key.ToString();
                        var groupRow = 0;
                        decimal GRANDTOTAL_MatlCost_PHP_TransQty = 0;
                        decimal GRANDTOTAL_MatlLandedCost_PHP_TransQty = 0;
                        decimal GRANDTOTAL_PIResin_PHP_TransQty = 0;
                        decimal GRANDTOTAL_PIFGProcess_PHP_TransQty = 0;
                        decimal GRANDTOTAL_PIHiddenProfit_PHP_TransQty = 0;
                        decimal GRANDTOTAL_SFAddedCost_PHP_FGAddedCost_PHP_TransQty = 0;
                        decimal GRANDTOTAL_TotalCost_PHP = 0;
                        foreach (var ReasonDescList in SummaryGroupByReasonDesc)
                        {
                            var ReasonDescList_ = ReasonDescList.ToList();
                            var ReasonDesc = "";
                            decimal TOTAL_MatlCost_PHP_TransQty = 0;
                            decimal TOTAL_MatlLandedCost_PHP_TransQty = 0;
                            decimal TOTAL_PIResin_PHP_TransQty = 0;
                            decimal TOTAL_PIFGProcess_PHP_TransQty = 0;
                            decimal TOTAL_PIHiddenProfit_PHP_TransQty = 0;
                            decimal TOTAL_SFAddedCost_PHP_FGAddedCost_PHP_TransQty = 0;
                            decimal TOTAL_TotalCost_PHP = 0;
                            var SFScrapList = ReasonDescList_
                                .Where(x => x.ReasonDesc == "SF Scrap")
                                .OrderBy(x=>x.Wc)
                                .GroupBy(x => x.Wc).ToList();
                            foreach (var SummarySheetData in ReasonDescList_)
                            {
                                ReasonDesc = SummarySheetData.ReasonDesc;
                                TOTAL_MatlCost_PHP_TransQty += Convert.ToDecimal(SummarySheetData.MatlCost_PHP) * Convert.ToDecimal(SummarySheetData.TransQty);
                                TOTAL_MatlLandedCost_PHP_TransQty += Convert.ToDecimal(SummarySheetData.MatlLandedCost_PHP) * Convert.ToDecimal(SummarySheetData.TransQty);
                                TOTAL_PIResin_PHP_TransQty += Convert.ToDecimal(SummarySheetData.PIResin_PHP) * Convert.ToDecimal(SummarySheetData.TransQty);
                                TOTAL_PIFGProcess_PHP_TransQty += Convert.ToDecimal(SummarySheetData.PIFGProcess_PHP) * Convert.ToDecimal(SummarySheetData.TransQty);
                                TOTAL_PIHiddenProfit_PHP_TransQty += Convert.ToDecimal(SummarySheetData.PIHiddenProfit_PHP) * Convert.ToDecimal(SummarySheetData.TransQty);
                                TOTAL_SFAddedCost_PHP_FGAddedCost_PHP_TransQty += (Convert.ToDecimal(SummarySheetData.SFAddedCost_PHP) + Convert.ToDecimal(SummarySheetData.FGAddedCost_PHP)) * Convert.ToDecimal(SummarySheetData.TransQty);
                                TOTAL_TotalCost_PHP += Convert.ToDecimal(SummarySheetData.TotalCost_PHP);

                                GRANDTOTAL_MatlCost_PHP_TransQty += Convert.ToDecimal(SummarySheetData.MatlCost_PHP) * Convert.ToDecimal(SummarySheetData.TransQty);
                                GRANDTOTAL_MatlLandedCost_PHP_TransQty += Convert.ToDecimal(SummarySheetData.MatlLandedCost_PHP) * Convert.ToDecimal(SummarySheetData.TransQty);
                                GRANDTOTAL_PIResin_PHP_TransQty += Convert.ToDecimal(SummarySheetData.PIResin_PHP) * Convert.ToDecimal(SummarySheetData.TransQty);
                                GRANDTOTAL_PIFGProcess_PHP_TransQty += Convert.ToDecimal(SummarySheetData.PIFGProcess_PHP) * Convert.ToDecimal(SummarySheetData.TransQty);
                                GRANDTOTAL_PIHiddenProfit_PHP_TransQty += Convert.ToDecimal(SummarySheetData.PIHiddenProfit_PHP) * Convert.ToDecimal(SummarySheetData.TransQty);
                                GRANDTOTAL_SFAddedCost_PHP_FGAddedCost_PHP_TransQty += (Convert.ToDecimal(SummarySheetData.SFAddedCost_PHP) + Convert.ToDecimal(SummarySheetData.FGAddedCost_PHP)) * Convert.ToDecimal(SummarySheetData.TransQty);
                                GRANDTOTAL_TotalCost_PHP += Convert.ToDecimal(SummarySheetData.TotalCost_PHP);
                            }
                            groupRow++;
                            if (groupRow < SummaryGroupByReasonDesc.Count)
                            {
                                Summary1stSheet.InsertRow((summary1stSheetRow + 1), 1);
                                Summary1stSheet.Cells[summary1stSheetRow, 1, summary1stSheetRow, 100].Copy(Summary1stSheet.Cells[(summary1stSheetRow + 1), 1, (summary1stSheetRow + 1), 1]);
                            }
                            Summary1stSheet.Cells[summary1stSheetRow, 1].Value = ReasonDesc ;
                            Summary1stSheet.Cells[summary1stSheetRow, 1].Style.WrapText = false;
                            Summary1stSheet.Cells[summary1stSheetRow, 3].Value = TOTAL_MatlCost_PHP_TransQty ;
                            Summary1stSheet.Cells[summary1stSheetRow, 3].Style.WrapText = false;
                            Summary1stSheet.Cells[summary1stSheetRow, 4].Value = TOTAL_MatlLandedCost_PHP_TransQty ;
                            Summary1stSheet.Cells[summary1stSheetRow, 4].Style.WrapText = false;
                            Summary1stSheet.Cells[summary1stSheetRow, 5].Value = TOTAL_PIResin_PHP_TransQty ;
                            Summary1stSheet.Cells[summary1stSheetRow, 5].Style.WrapText = false;
                            Summary1stSheet.Cells[summary1stSheetRow, 6].Value = TOTAL_PIFGProcess_PHP_TransQty ;
                            Summary1stSheet.Cells[summary1stSheetRow, 6].Style.WrapText = false;
                            Summary1stSheet.Cells[summary1stSheetRow, 7].Value = TOTAL_PIHiddenProfit_PHP_TransQty ;
                            Summary1stSheet.Cells[summary1stSheetRow, 7].Style.WrapText = false;
                            Summary1stSheet.Cells[summary1stSheetRow, 8].Value = TOTAL_SFAddedCost_PHP_FGAddedCost_PHP_TransQty ;
                            Summary1stSheet.Cells[summary1stSheetRow, 8].Style.WrapText = false;
                            Summary1stSheet.Cells[summary1stSheetRow, 9].Value = TOTAL_TotalCost_PHP ;
                            Summary1stSheet.Cells[summary1stSheetRow, 9].Style.WrapText = false;
                            summary1stSheetRow++;

                            int sfScrapRow = 0;
                            if(ReasonDesc=="SF Scrap"){
                                foreach (var SFScrapDataWC in SFScrapList)
                                {
                                    decimal SCRAP_MatlCost_PHP_TransQty = 0;
                                    decimal SCRAP_MatlLandedCost_PHP_TransQty = 0;
                                    decimal SCRAP_PIResin_PHP_TransQty = 0;
                                    decimal SCRAP_PIFGProcess_PHP_TransQty = 0;
                                    decimal SCRAP_PIHiddenProfit_PHP_TransQty = 0;
                                    decimal SCRAP_SFAddedCost_PHP_FGAddedCost_PHP_TransQty = 0;
                                    decimal SCRAP_TotalCost_PHP = 0;
                                    foreach (var SFScrapData in SFScrapDataWC)
                                    {
                                        SCRAP_MatlCost_PHP_TransQty += Convert.ToDecimal(SFScrapData.MatlCost_PHP) * Convert.ToDecimal(SFScrapData.TransQty);
                                        SCRAP_MatlLandedCost_PHP_TransQty += Convert.ToDecimal(SFScrapData.MatlLandedCost_PHP) * Convert.ToDecimal(SFScrapData.TransQty);
                                        SCRAP_PIResin_PHP_TransQty += Convert.ToDecimal(SFScrapData.PIResin_PHP) * Convert.ToDecimal(SFScrapData.TransQty);
                                        SCRAP_PIFGProcess_PHP_TransQty += Convert.ToDecimal(SFScrapData.PIFGProcess_PHP) * Convert.ToDecimal(SFScrapData.TransQty);
                                        SCRAP_PIHiddenProfit_PHP_TransQty += Convert.ToDecimal(SFScrapData.PIHiddenProfit_PHP) * Convert.ToDecimal(SFScrapData.TransQty);
                                        SCRAP_SFAddedCost_PHP_FGAddedCost_PHP_TransQty += (Convert.ToDecimal(SFScrapData.SFAddedCost_PHP) + Convert.ToDecimal(SFScrapData.FGAddedCost_PHP)) * Convert.ToDecimal(SFScrapData.TransQty);
                                        SCRAP_TotalCost_PHP += Convert.ToDecimal(SFScrapData.TotalCost_PHP);
                                    }
                                    sfScrapRow++;
                                    if (sfScrapRow <= SFScrapList.Count)
                                    {
                                        Summary1stSheet.InsertRow((summary1stSheetRow), 1);
                                        Summary1stSheet.Cells[summary1stSheetRow - 1, 1, summary1stSheetRow - 1, 100].Copy(Summary1stSheet.Cells[(summary1stSheetRow), 1, (summary1stSheetRow), 1]);
                                    }
                                    Summary1stSheet.Cells["A" + summary1stSheetRow + ":B" + summary1stSheetRow].Merge = false;
                                    Summary1stSheet.Cells[summary1stSheetRow, 1].Value = "";
                                    Summary1stSheet.Cells[summary1stSheetRow, 1].Style.WrapText = false;
                                    Summary1stSheet.Cells[summary1stSheetRow, 2].Value = SFScrapDataWC.Key.ToString();
                                    Summary1stSheet.Cells[summary1stSheetRow, 2].Style.WrapText = false;
                                    Summary1stSheet.Cells[summary1stSheetRow, 3].Value = SCRAP_MatlCost_PHP_TransQty;
                                    Summary1stSheet.Cells[summary1stSheetRow, 3].Style.WrapText = false;
                                    Summary1stSheet.Cells[summary1stSheetRow, 4].Value = SCRAP_MatlLandedCost_PHP_TransQty;
                                    Summary1stSheet.Cells[summary1stSheetRow, 4].Style.WrapText = false;
                                    Summary1stSheet.Cells[summary1stSheetRow, 5].Value = SCRAP_PIResin_PHP_TransQty;
                                    Summary1stSheet.Cells[summary1stSheetRow, 5].Style.WrapText = false;
                                    Summary1stSheet.Cells[summary1stSheetRow, 6].Value = SCRAP_PIFGProcess_PHP_TransQty;
                                    Summary1stSheet.Cells[summary1stSheetRow, 6].Style.WrapText = false;
                                    Summary1stSheet.Cells[summary1stSheetRow, 7].Value = SCRAP_PIHiddenProfit_PHP_TransQty;
                                    Summary1stSheet.Cells[summary1stSheetRow, 7].Style.WrapText = false;
                                    Summary1stSheet.Cells[summary1stSheetRow, 8].Value = SCRAP_SFAddedCost_PHP_FGAddedCost_PHP_TransQty;
                                    Summary1stSheet.Cells[summary1stSheetRow, 8].Style.WrapText = false;
                                    Summary1stSheet.Cells[summary1stSheetRow, 9].Value = SCRAP_TotalCost_PHP;
                                    Summary1stSheet.Cells[summary1stSheetRow, 9].Style.WrapText = false;
                                    summary1stSheetRow++;
                                }
                            }
                        }

                        Summary1stSheet.Cells[summary1stSheetRow, 3].Value = GRANDTOTAL_MatlCost_PHP_TransQty;
                        Summary1stSheet.Cells[summary1stSheetRow, 3].Style.WrapText = false;
                        Summary1stSheet.Cells[summary1stSheetRow, 4].Value = GRANDTOTAL_MatlLandedCost_PHP_TransQty;
                        Summary1stSheet.Cells[summary1stSheetRow, 4].Style.WrapText = false;
                        Summary1stSheet.Cells[summary1stSheetRow, 5].Value = GRANDTOTAL_PIResin_PHP_TransQty;
                        Summary1stSheet.Cells[summary1stSheetRow, 5].Style.WrapText = false;
                        Summary1stSheet.Cells[summary1stSheetRow, 6].Value = GRANDTOTAL_PIFGProcess_PHP_TransQty;
                        Summary1stSheet.Cells[summary1stSheetRow, 6].Style.WrapText = false;
                        Summary1stSheet.Cells[summary1stSheetRow, 7].Value = GRANDTOTAL_PIHiddenProfit_PHP_TransQty;
                        Summary1stSheet.Cells[summary1stSheetRow, 7].Style.WrapText = false;
                        Summary1stSheet.Cells[summary1stSheetRow, 8].Value = GRANDTOTAL_SFAddedCost_PHP_FGAddedCost_PHP_TransQty;
                        Summary1stSheet.Cells[summary1stSheetRow, 8].Style.WrapText = false;
                        Summary1stSheet.Cells[summary1stSheetRow, 9].Value = GRANDTOTAL_TotalCost_PHP;
                        Summary1stSheet.Cells[summary1stSheetRow, 9].Style.WrapText = false;

                        summaryGroupCtr++;
                        if (summaryGroupCtr < Summary_GroupBySummaryGroup.Count)
                        {
                            Summary1stSheet.InsertRow((summary1stSheetRow + 1), 1);
                            Summary1stSheet.Cells[4, 1, 4, 100].Copy(Summary1stSheet.Cells[(summary1stSheetRow + 2), 1, (summary1stSheetRow + 2), 1]);
                            Summary1stSheet.Cells[5, 1, 5, 100].Copy(Summary1stSheet.Cells[(summary1stSheetRow + 3), 1, (summary1stSheetRow + 3), 1]);
                            Summary1stSheet.Cells[summary1stSheetRow, 1, summary1stSheetRow, 100].Copy(Summary1stSheet.Cells[(summary1stSheetRow + 4), 1, (summary1stSheetRow + 4), 1]);
                            summary1stSheetRow += 3;
                        }
                    }
                    #endregion
                    #region SF Scrap Summary

                    List<ExcelColumns> SFScrapSummarySheet = new List<ExcelColumns>();
                    ExcelWorksheet SummarySheet = excelPackage.Workbook.Worksheets["SF Scrap Summary"];
                    var SFScrapSummaryDataRaw = MiscellaneousTransaction
                        .Where(x => (x.TransDesc == "SF Scrap Data"))
                        .OrderBy(x => x.Wc)
                        .ToList();
                    var groupedSFScrapSummaryDataWC = SFScrapSummaryDataRaw
                        .GroupBy(u => u.Wc)
                        .Select(grp => grp.ToList())
                        .ToList();
                    int TotalRow = 4;
                    int rowCounter = 1;

                    decimal GrandTotalMatlCost_PHP = 0;
                    decimal GrandTotalMatlLandedCost_PHP = 0;
                    decimal GrandTotalPIResin_PHP = 0;
                    decimal GrandTotalPIFGProcess_PHP = 0;
                    decimal GrandTotalPIHiddenProfit_PHP = 0;
                    decimal GrandTotalFGAddedCost_PHP = 0;
                    decimal GrandTotalTotalCost_PHP = 0;

                    int plusRows = groupedSFScrapSummaryDataWC.Count;
                    int SummarySheetSheetrRow = 11 + plusRows - 1;
                    foreach (var SFScrapSummaryData in groupedSFScrapSummaryDataWC)
                    {
                        var groupedSFScrapSummaryDataItem = SFScrapSummaryData
                        .OrderBy(x => x.Item)
                        .GroupBy(u => u.Item)
                        .Select(grp => grp.ToList())
                        .ToList();
                        int groupedItemCount = SFScrapSummaryData.Count;
                        int startMerge = SummarySheetSheetrRow;

                        string TotalWC = "";
                        decimal TotalMatlCost_PHP = 0;
                        decimal TotalMatlLandedCost_PHP = 0;
                        decimal TotalPIResin_PHP = 0;
                        decimal TotalPIFGProcess_PHP = 0;
                        decimal TotalPIHiddenProfit_PHP = 0;
                        decimal TotalFGAddedCost_PHP = 0;
                        decimal TotalTotalCost_PHP = 0;

                        foreach (var PerItem in groupedSFScrapSummaryDataItem)
                        {
                            foreach (var ItemData in PerItem)
                            {
                                TotalWC = ItemData.Wc;
                                TotalMatlCost_PHP += Convert.ToDecimal(ItemData.MatlCost_PHP * ItemData.QtyScrapped);
                                TotalMatlLandedCost_PHP += Convert.ToDecimal(ItemData.MatlLandedCost_PHP * ItemData.QtyScrapped);
                                TotalPIResin_PHP += Convert.ToDecimal(ItemData.PIResin_PHP * ItemData.QtyScrapped);
                                TotalPIFGProcess_PHP += Convert.ToDecimal(ItemData.PIFGProcess_PHP * ItemData.QtyScrapped);
                                TotalPIHiddenProfit_PHP += Convert.ToDecimal(ItemData.PIHiddenProfit_PHP * ItemData.QtyScrapped);
                                TotalFGAddedCost_PHP += Convert.ToDecimal((ItemData.SFAddedCost_PHP + ItemData.FGAddedCost_PHP) * ItemData.QtyScrapped);
                                TotalTotalCost_PHP += Convert.ToDecimal(ItemData.TotalCost_PHP * -1);


                                GrandTotalMatlCost_PHP += Convert.ToDecimal(ItemData.MatlCost_PHP * ItemData.QtyScrapped); ;
                                GrandTotalMatlLandedCost_PHP += Convert.ToDecimal(ItemData.MatlLandedCost_PHP * ItemData.QtyScrapped); ;
                                GrandTotalPIResin_PHP += Convert.ToDecimal(ItemData.PIResin_PHP * ItemData.QtyScrapped); ;
                                GrandTotalPIFGProcess_PHP += Convert.ToDecimal(ItemData.PIFGProcess_PHP * ItemData.QtyScrapped); ;
                                GrandTotalPIHiddenProfit_PHP += Convert.ToDecimal(ItemData.PIHiddenProfit_PHP * ItemData.QtyScrapped); ;
                                GrandTotalFGAddedCost_PHP += Convert.ToDecimal((ItemData.SFAddedCost_PHP + ItemData.FGAddedCost_PHP) * ItemData.QtyScrapped); ;
                                GrandTotalTotalCost_PHP += Convert.ToDecimal(ItemData.TotalCost_PHP * -1); ;
                                rowCounter++;
                            }
                        }
                        if (rowCounter <= SFScrapSummaryDataRaw.Count)
                        {
                            SummarySheet.InsertRow((TotalRow + 1), 1);
                            SummarySheet.Cells[(TotalRow), 1, (TotalRow), 100].Copy(SummarySheet.Cells[(TotalRow + 1), 1, (TotalRow + 1), 1]);
                        }
                        if (rowCounter > SFScrapSummaryDataRaw.Count)
                        {
                            SummarySheet.Cells[(TotalRow + 1), 4].Value = Convert.ToDecimal(GrandTotalMatlCost_PHP);
                            SummarySheet.Cells[(TotalRow + 1), 4].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow + 1), 5].Value = Convert.ToDecimal(GrandTotalMatlLandedCost_PHP);
                            SummarySheet.Cells[(TotalRow + 1), 5].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow + 1), 6].Value = Convert.ToDecimal(GrandTotalPIResin_PHP);
                            SummarySheet.Cells[(TotalRow + 1), 6].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow + 1), 7].Value = Convert.ToDecimal(GrandTotalPIFGProcess_PHP);
                            SummarySheet.Cells[(TotalRow + 1), 7].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow + 1), 8].Value = Convert.ToDecimal(GrandTotalPIHiddenProfit_PHP);
                            SummarySheet.Cells[(TotalRow + 1), 8].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow + 1), 9].Value = Convert.ToDecimal(GrandTotalFGAddedCost_PHP);
                            SummarySheet.Cells[(TotalRow + 1), 9].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow + 1), 10].Value = Convert.ToDecimal(GrandTotalTotalCost_PHP);
                            SummarySheet.Cells[(TotalRow + 1), 10].Style.WrapText = false;
                        }
                        SummarySheet.Cells[TotalRow, 3].Value = TotalWC;
                        SummarySheet.Cells[TotalRow, 3].Style.WrapText = false;
                        SummarySheet.Cells[TotalRow, 4].Value = Convert.ToDecimal(TotalMatlCost_PHP);
                        SummarySheet.Cells[TotalRow, 4].Style.WrapText = false;
                        SummarySheet.Cells[TotalRow, 5].Value = Convert.ToDecimal(TotalMatlLandedCost_PHP);
                        SummarySheet.Cells[TotalRow, 5].Style.WrapText = false;
                        SummarySheet.Cells[TotalRow, 6].Value = Convert.ToDecimal(TotalPIResin_PHP);
                        SummarySheet.Cells[TotalRow, 6].Style.WrapText = false;
                        SummarySheet.Cells[TotalRow, 7].Value = Convert.ToDecimal(TotalPIFGProcess_PHP);
                        SummarySheet.Cells[TotalRow, 7].Style.WrapText = false;
                        SummarySheet.Cells[TotalRow, 8].Value = Convert.ToDecimal(TotalPIHiddenProfit_PHP);
                        SummarySheet.Cells[TotalRow, 8].Style.WrapText = false;
                        SummarySheet.Cells[TotalRow, 9].Value = Convert.ToDecimal(TotalFGAddedCost_PHP);
                        SummarySheet.Cells[TotalRow, 9].Style.WrapText = false;
                        SummarySheet.Cells[TotalRow, 10].Value = Convert.ToDecimal(TotalTotalCost_PHP);
                        SummarySheet.Cells[TotalRow, 10].Style.WrapText = false;
                        TotalRow++;
                    }

                    rowCounter = 1;
                    foreach (var SFScrapSummaryData in groupedSFScrapSummaryDataWC)
                    {
                        var groupedSFScrapSummaryDataItem = SFScrapSummaryData
                        .OrderBy(x => x.Item)
                        .GroupBy(u => u.Item)
                        .Select(grp => grp.ToList())
                        .ToList();
                        decimal Sum_TotalCost_PHP = 0;
                        string sumWCTitle = "";
                        int itemRowCounter = 0;
                        int groupedItemCount = SFScrapSummaryData.Count;
                        int startMerge = SummarySheetSheetrRow;
                        int endMerge = 0;

                        string TotalWC = "";
                        decimal TotalMatlCost_PHP = 0;
                        decimal TotalMatlLandedCost_PHP = 0;
                        decimal TotalPIResin_PHP = 0;
                        decimal TotalPIFGProcess_PHP = 0;
                        decimal TotalPIHiddenProfit_PHP = 0;
                        decimal TotalFGAddedCost_PHP = 0;
                        decimal TotalTotalCost_PHP = 0;

                        foreach (var PerItem in groupedSFScrapSummaryDataItem)
                        {
                            int startMergeItem = SummarySheetSheetrRow;
                            int endMergeItem = 0;
                            foreach (var ItemData in PerItem)
                            {
                                if (rowCounter < SFScrapSummaryDataRaw.Count && itemRowCounter < (groupedItemCount - 1))
                                {
                                    SummarySheet.InsertRow((SummarySheetSheetrRow + 1), 1);
                                    SummarySheet.Cells[SummarySheetSheetrRow, 1, SummarySheetSheetrRow, 100].Copy(SummarySheet.Cells[(SummarySheetSheetrRow + 1), 1, (SummarySheetSheetrRow + 1), 1]);
                                }
                                SummarySheet.Cells[SummarySheetSheetrRow, 1].Value = ItemData.Wc;
                                SummarySheet.Cells[SummarySheetSheetrRow, 1].Style.WrapText = false;
                                SummarySheet.Cells[SummarySheetSheetrRow, 2].Value = ItemData.Item;
                                SummarySheet.Cells[SummarySheetSheetrRow, 2].Style.WrapText = false;
                                SummarySheet.Cells[SummarySheetSheetrRow, 3].Value = ItemData.ItemDesc;
                                SummarySheet.Cells[SummarySheetSheetrRow, 3].Style.WrapText = true;
                                SummarySheet.Cells[SummarySheetSheetrRow, 4].Value = Convert.ToInt32(ItemData.QtyScrapped);
                                SummarySheet.Cells[SummarySheetSheetrRow, 4].Style.WrapText = false;
                                SummarySheet.Cells[SummarySheetSheetrRow, 5].Value = Convert.ToDecimal(ItemData.MatlCost_PHP);
                                SummarySheet.Cells[SummarySheetSheetrRow, 5].Style.WrapText = false;
                                SummarySheet.Cells[SummarySheetSheetrRow, 6].Value = Convert.ToDecimal(ItemData.MatlLandedCost_PHP);
                                SummarySheet.Cells[SummarySheetSheetrRow, 6].Style.WrapText = false;
                                SummarySheet.Cells[SummarySheetSheetrRow, 7].Value = Convert.ToDecimal(ItemData.PIResin_PHP);
                                SummarySheet.Cells[SummarySheetSheetrRow, 7].Style.WrapText = false;
                                SummarySheet.Cells[SummarySheetSheetrRow, 8].Value = Convert.ToDecimal(ItemData.PIFGProcess_PHP);
                                SummarySheet.Cells[SummarySheetSheetrRow, 8].Style.WrapText = false;
                                SummarySheet.Cells[SummarySheetSheetrRow, 9].Value = Convert.ToDecimal(ItemData.PIHiddenProfit_PHP);
                                SummarySheet.Cells[SummarySheetSheetrRow, 9].Style.WrapText = false;
                                SummarySheet.Cells[SummarySheetSheetrRow, 10].Value = Convert.ToDecimal((ItemData.SFAddedCost_PHP + ItemData.FGAddedCost_PHP));
                                SummarySheet.Cells[SummarySheetSheetrRow, 10].Style.WrapText = false;
                                SummarySheet.Cells[SummarySheetSheetrRow, 11].Value = Convert.ToDecimal(ItemData.TotalCost_PHP * -1);
                                SummarySheet.Cells[SummarySheetSheetrRow, 11].Style.WrapText = false;

                                Sum_TotalCost_PHP += (ItemData.TotalCost_PHP * -1);
                                sumWCTitle = ItemData.Wc;
                                endMerge = SummarySheetSheetrRow;
                                endMergeItem = SummarySheetSheetrRow;

                                TotalWC = ItemData.Wc;
                                TotalMatlCost_PHP = Convert.ToDecimal(ItemData.MatlCost_PHP * ItemData.QtyScrapped);
                                TotalMatlLandedCost_PHP = Convert.ToDecimal(ItemData.MatlLandedCost_PHP * ItemData.QtyScrapped);
                                TotalPIResin_PHP = Convert.ToDecimal(ItemData.PIResin_PHP * ItemData.QtyScrapped);
                                TotalPIFGProcess_PHP = Convert.ToDecimal(ItemData.PIFGProcess_PHP * ItemData.QtyScrapped);
                                TotalPIHiddenProfit_PHP = Convert.ToDecimal(ItemData.PIHiddenProfit_PHP * ItemData.QtyScrapped);
                                TotalFGAddedCost_PHP = Convert.ToDecimal((ItemData.SFAddedCost_PHP + ItemData.FGAddedCost_PHP) * ItemData.QtyScrapped);
                                TotalTotalCost_PHP = Convert.ToDecimal(ItemData.TotalCost_PHP * -1);


                                rowCounter++;
                                itemRowCounter++;
                                SummarySheetSheetrRow++;
                            }
                            SummarySheet.Cells["B" + startMergeItem + ":B" + endMergeItem].Merge = true;
                            SummarySheet.Cells["C" + startMergeItem + ":C" + endMergeItem].Merge = true;
                        }
                        if (rowCounter < SFScrapSummaryDataRaw.Count)
                        {
                            SummarySheet.InsertRow((SummarySheetSheetrRow + 1), 1);
                            SummarySheet.Cells[(SummarySheetSheetrRow - 1), 1, (SummarySheetSheetrRow - 1), 100].Copy(SummarySheet.Cells[(SummarySheetSheetrRow + 1), 1, (SummarySheetSheetrRow + 1), 1]);
                            SummarySheet.InsertRow((SummarySheetSheetrRow + 2), 1);
                            SummarySheet.Cells[SummarySheetSheetrRow, 1, SummarySheetSheetrRow, 100].Copy(SummarySheet.Cells[(SummarySheetSheetrRow + 2), 1, (SummarySheetSheetrRow + 2), 1]);
                            SummarySheet.Cells["A" + startMerge + ":A" + endMerge].Merge = true;

                        }
                        if ((rowCounter - 1) <= SFScrapSummaryDataRaw.Count)
                        {
                            SummarySheet.Cells[SummarySheetSheetrRow, 1].Value = sumWCTitle;
                            SummarySheet.Cells[SummarySheetSheetrRow, 1].Style.WrapText = false;
                            SummarySheet.Cells[SummarySheetSheetrRow, 11].Value = Convert.ToDecimal(Sum_TotalCost_PHP);
                            SummarySheet.Cells[SummarySheetSheetrRow, 11].Style.WrapText = false;
                            SummarySheetSheetrRow++;
                        }
                    }

                    #endregion SF Scrap Summary
                    #region SFScrapDataSheet(SFScrapDataSheet)

                    List<ExcelColumns> SFScrapDataSheet = new List<ExcelColumns>();
                    ExcelWorksheet MiscellaneousTransactionSheet = excelPackage.Workbook.Worksheets["SF Scrap Data"];
                    foreach (var MiscellaneousTransactionItem in filteredMiscellaneousTransaction)
                    {
                        SFScrapDataSheet.Add(new ExcelColumns
                        {
                            A = MiscellaneousTransactionItem.JobOrLot,
                            B = DateTime.Parse(MiscellaneousTransactionItem.TransDate).ToString("MM/dd/yyyy"),
                            C = MiscellaneousTransactionItem.Item,
                            D = MiscellaneousTransactionItem.ItemDesc,
                            E = MiscellaneousTransactionItem.QtyCompleted.ToString(),
                            F = MiscellaneousTransactionItem.QtyScrapped.ToString(),
                            G = MiscellaneousTransactionItem.Employee.ToString(),
                            H = MiscellaneousTransactionItem.Wc.ToString(),
                            I = MiscellaneousTransactionItem.MatlCost_PHP.ToString(),
                            J = MiscellaneousTransactionItem.MatlLandedCost_PHP.ToString(),
                            K = MiscellaneousTransactionItem.PIResin_PHP.ToString(),
                            L = MiscellaneousTransactionItem.PIFGProcess_PHP.ToString(),
                            M = MiscellaneousTransactionItem.PIHiddenProfit_PHP.ToString(),
                            N = (MiscellaneousTransactionItem.SFAddedCost_PHP + MiscellaneousTransactionItem.FGAddedCost_PHP).ToString(),
                            O = (MiscellaneousTransactionItem.TotalCost_PHP * -1).ToString(),

                        });
                    }
                    int sheetrRow = 6;
                    foreach (var SheetData in SFScrapDataSheet)
                    {
                        if (sheetrRow < SFScrapDataSheet.Count + 5)
                        {
                            MiscellaneousTransactionSheet.InsertRow((sheetrRow + 1), 1);
                            MiscellaneousTransactionSheet.Cells[sheetrRow, 1, sheetrRow, 100].Copy(MiscellaneousTransactionSheet.Cells[(sheetrRow + 1), 1, (sheetrRow + 1), 1]);
                        }
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 1].Value = SheetData.A;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 1].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 2].Value = SheetData.B;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 2].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 3].Value = SheetData.C;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 3].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 4].Value = SheetData.D;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 4].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 5].Value = Convert.ToInt32(SheetData.E);
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 5].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 6].Value = Convert.ToInt32(SheetData.F);
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 6].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 7].Value = Convert.ToInt32(SheetData.G);
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 7].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 8].Value = SheetData.H;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 8].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 9].Value = Convert.ToDecimal(SheetData.I);
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 9].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 10].Value = Convert.ToDecimal(SheetData.J);
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 10].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 11].Value = Convert.ToDecimal(SheetData.K);
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 11].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 12].Value = Convert.ToDecimal(SheetData.L);
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 12].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 13].Value = Convert.ToDecimal(SheetData.M);
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 13].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 14].Value = Convert.ToDecimal(SheetData.N);
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 14].Style.WrapText = false;
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 15].Value = Convert.ToDecimal(SheetData.O);
                        MiscellaneousTransactionSheet.Cells[sheetrRow, 15].Style.WrapText = false;
                        sheetrRow++;
                    }
                    #endregion
                    return File(excelPackage.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Filename);
                }
            }
            catch (Exception err)
            {
                string errmsg = "";
                if (err.InnerException != null)
                    errmsg = "An error occured: " + err.InnerException.ToString();
                else
                    errmsg = "An error occured: " + err.Message.ToString();
                return null;
            }
        }
        public ActionResult GenerateFinishedGoodsAndSalesReport()
        {
            List<LSP_Rpt_DM_FinishedGoodsSalesReport> LSP_Rpt_DM_FinishedGoodsSalesReportList = new List<LSP_Rpt_DM_FinishedGoodsSalesReport>();
            var StartDate = Request["StartDate"];
            var EndDate = Request["EndDate"];
            string MonthYear = DateTime.Parse(StartDate).ToString("MMMyyyy");
            try
            {
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LSPI803_App"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {

                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "LSP_Rpt_NewDM_FinishedGoodsReportSp";
                        cmdSql.Parameters.AddWithValue("@StartDate", StartDate);
                        cmdSql.Parameters.AddWithValue("@EndDate", EndDate);
                        cmdSql.CommandTimeout = 0;
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                LSP_Rpt_DM_FinishedGoodsSalesReportList.Add(new LSP_Rpt_DM_FinishedGoodsSalesReport
                                {
                                    FGTransType = sdr["FGTransType"].ToString(),
                                    TransDate = sdr["TransDate"].ToString(),
                                    PONum = sdr["PONum"].ToString() + sdr["FGTransType"].ToString(),
                                    CustomerName = sdr["CustomerName"].ToString(),
                                    JobOrder = sdr["JobOrder"].ToString() + sdr["JobSuffix"].ToString(),
                                    Item = sdr["Item"].ToString(),
                                    ItemDesc = sdr["ItemDesc"].ToString(),
                                    ProductCode = sdr["ProductCode"].ToString(),
                                    Family = sdr["Family"].ToString(),
                                    FamilyDesc = sdr["FamilyDesc"].ToString(),
                                    QtyCompleted = sdr["QtyCompleted"] == null ? 0: Convert.ToDecimal(sdr["QtyCompleted"]),
                                    StdMatlCost_PHP = sdr["StdMatlCost_PHP"] == null ? 0: Convert.ToDecimal(sdr["StdMatlCost_PHP"]),    
                                    StdResinCost_PHP = sdr["StdResinCost_PHP"] == null ? 0: Convert.ToDecimal(sdr["StdResinCost_PHP"]), 
                                    StdPIProcess_PHP = sdr["StdPIProcess_PHP"] == null ? 0: Convert.ToDecimal(sdr["StdPIProcess_PHP"]), 
                                    StdHiddenProfit_PHP = sdr["StdHiddenProfit_PHP"] == null ? 0: Convert.ToDecimal(sdr["StdHiddenProfit_PHP"]),    
                                    StdSFAdded_PHP = sdr["StdSFAdded_PHP"] == null ? 0: Convert.ToDecimal(sdr["StdSFAdded_PHP"]),   
                                    StdFGAdded_PHP = sdr["StdFGAdded_PHP"] == null ? 0: Convert.ToDecimal(sdr["StdFGAdded_PHP"]),   
                                    StdUnitCost_PHP = sdr["StdUnitCost_PHP"] == null ? 0: Convert.ToDecimal(sdr["StdUnitCost_PHP"]),    
                                    ActlMatlUnitCost_PHP = sdr["ActlMatlUnitCost_PHP"] == null ? 0: Convert.ToDecimal(sdr["ActlMatlUnitCost_PHP"]), 
                                    ActlLandedCost_PHP = sdr["ActlLandedCost_PHP"] == null ? 0: Convert.ToDecimal(sdr["ActlLandedCost_PHP"]),   
                                    ActlResinCost_PHP = sdr["ActlResinCost_PHP"] == null ? 0: Convert.ToDecimal(sdr["ActlResinCost_PHP"]),  
                                    ActlPIProcess_PHP = sdr["ActlPIProcess_PHP"] == null ? 0: Convert.ToDecimal(sdr["ActlPIProcess_PHP"]),  
                                    ActlHiddenProfit_PHP = sdr["ActlHiddenProfit_PHP"] == null ? 0: Convert.ToDecimal(sdr["ActlHiddenProfit_PHP"]), 
                                    ActlSFAdded_PHP = sdr["ActlSFAdded_PHP"] == null ? 0: Convert.ToDecimal(sdr["ActlSFAdded_PHP"]),
                                    ActlFGAdded_PHP = sdr["ActlFGAdded_PHP"] == null ? 0: Convert.ToDecimal(sdr["ActlFGAdded_PHP"]),    
                                    ActlUnitCost_PHP = sdr["ActlUnitCost_PHP"] == null ? 0: Convert.ToDecimal(sdr["ActlUnitCost_PHP"]), 
                                });
                            }

                        }
                    }
                    conn.Close();
                }
                var LSP_Rpt_DM_FinishedGoodsSalesReportList_FinishedGood = LSP_Rpt_DM_FinishedGoodsSalesReportList
                    .Where(x=>x.FGTransType== "FINISHED GOODS")
                    .OrderBy(x => x.TransDate)
                    .ToList();
                var LSP_Rpt_DM_FinishedGoodsSalesReportList_NotFinishedGood = LSP_Rpt_DM_FinishedGoodsSalesReportList
                    .Where(x => x.FGTransType != "FINISHED GOODS")
                    .OrderBy(x => x.TransDate)
                    .ToList();
                var LSP_Rpt_DM_FinishedGoodsSalesReportList_GroupByTransTpeNotFinishedGood = LSP_Rpt_DM_FinishedGoodsSalesReportList
                    .Where(x => x.FGTransType != "FINISHED GOODS")
                    .OrderBy(x => x.TransDate)
                    .GroupBy(x => x.FGTransType)
                    .ToList();
                int LSP_Rpt_DM_FinishedGoodsSalesReportList_NotFinishedGoodCount = LSP_Rpt_DM_FinishedGoodsSalesReportList_NotFinishedGood.ToList().Count;
                int LSP_Rpt_DM_FinishedGoodsSalesReportList_FinishedGoodCount = LSP_Rpt_DM_FinishedGoodsSalesReportList_FinishedGood.Count;
                int CurrentDataCount = LSP_Rpt_DM_FinishedGoodsSalesReportList_NotFinishedGoodCount + LSP_Rpt_DM_FinishedGoodsSalesReportList_FinishedGoodCount + 4 + 2;
                string filePath = "";
                string Filename = "LSP_Rpt_DM_FinishedGoodsSalesReport_"+ MonthYear+".xlsx";
                filePath = Path.Combine(Server.MapPath("~/Areas/Reports/Templates/") + "LSP_Rpt_DM_FinishedGoodsSalesReport.xlsx");
                FileInfo file = new FileInfo(filePath);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    #region FINISHEDGOODS 
                    ExcelWorksheet FINISHEDGOODS = excelPackage.Workbook.Worksheets["FINISHED GOODS"];
                    decimal Total_QtyCompleted = 0;
                    decimal Total_StdMatlCost_PHPxQtyCompleted = 0;
                    decimal Total_StdResinCost_PHPxQtyCompleted = 0;
                    decimal Total_StdPIProcess_PHPxQtyCompleted = 0;
                    decimal Total_StdHiddenProfit_PHPxQtyCompleted = 0;
                    decimal Total_StdSFAdded_PHPxQtyCompleted = 0;
                    decimal Total_StdFGAdded_PHPxQtyCompleted = 0;
                    decimal Total_StdUnitCost_PHPxQtyCompleted = 0;
                    decimal Total_ActlMatlUnitCost_PHPxQtyCompleted = 0;
                    decimal Total_ActlLandedCost_PHPxQtyCompleted = 0;
                    decimal Total_ActlResinCost_PHPxQtyCompleted = 0;
                    decimal Total_ActlPIProcess_PHPxQtyCompleted = 0;
                    decimal Total_ActlHiddenProfit_PHPxQtyCompleted = 0;
                    decimal Total_ActlSFAdded_PHPxQtyCompleted = 0;
                    decimal Total_ActlFGAdded_PHP_PHPxQtyCompleted = 0;
                    decimal Total_ActlUnitCost_PHPxQtyCompleted = 0;
                    decimal Total_StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                    decimal Total_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                    decimal Total___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ = 0;

                    int sheetsRowFinishedGoods = 5;
                    #region FINISHEDGOODS FINISHEDGOODS 
                    foreach (var LSP_Rpt_DM_FinishedGoodsSalesReportObjList in LSP_Rpt_DM_FinishedGoodsSalesReportList_FinishedGood)
                    {
                        string FGTransType = "";
                        string TransDate = "";
                        string PONum = "";
                        string CustomerName = "";
                        string JobOrder = "";
                        string Item = "";
                        string ItemDesc = "";
                        string ProductCode = "";
                        string FamilyDesc = "";
                        decimal QtyCompleted = 0;
                        decimal StdMatlCost_PHPxQtyCompleted = 0;
                        decimal StdResinCost_PHPxQtyCompleted = 0;
                        decimal StdPIProcess_PHPxQtyCompleted = 0;
                        decimal StdHiddenProfit_PHPxQtyCompleted = 0;
                        decimal StdSFAdded_PHPxQtyCompleted = 0;
                        decimal StdFGAdded_PHPxQtyCompleted = 0;
                        decimal StdUnitCost_PHPxQtyCompleted = 0;
                        decimal ActlMatlUnitCost_PHPxQtyCompleted = 0;
                        decimal ActlLandedCost_PHPxQtyCompleted = 0;
                        decimal ActlResinCost_PHPxQtyCompleted = 0;
                        decimal ActlPIProcess_PHPxQtyCompleted = 0;
                        decimal ActlHiddenProfit_PHPxQtyCompleted = 0;
                        decimal ActlSFAdded_PHPxQtyCompleted = 0;
                        decimal ActlFGAdded_PHP_PHPxQtyCompleted = 0;
                        decimal ActlUnitCost_PHPxQtyCompleted = 0;
                        decimal StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                        decimal StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                        decimal __StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ = 0;
                        //foreach (var LSP_Rpt_DM_FinishedGoodsSalesReportObj in LSP_Rpt_DM_FinishedGoodsSalesReportObjList)
                        //{
                        FGTransType = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.FGTransType;
                        TransDate = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.TransDate;
                        PONum = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.PONum;
                        CustomerName = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.CustomerName;
                        JobOrder = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.JobOrder;
                        Item = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.Item;
                        ItemDesc = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ItemDesc;
                        ProductCode = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ProductCode;
                        FamilyDesc = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.FamilyDesc;

                        QtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        StdMatlCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdMatlCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        StdResinCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        StdPIProcess_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        StdHiddenProfit_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        StdSFAdded_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        StdFGAdded_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        StdUnitCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        ActlMatlUnitCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlMatlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        ActlLandedCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlLandedCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        ActlResinCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        ActlPIProcess_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        ActlHiddenProfit_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        ActlSFAdded_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        ActlFGAdded_PHP_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        ActlUnitCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        StdUnitCost_PHPminusActlUnitCost_PHP = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlUnitCost_PHP;
                        StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdUnitCost_PHP - (LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlLandedCost_PHP);
                        __StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ = (LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted) - (LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted);

                        Total_QtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_StdMatlCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdMatlCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_StdResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_StdPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_StdHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_StdSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_StdFGAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_StdUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_ActlMatlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlMatlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_ActlLandedCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlLandedCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_ActlResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_ActlPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_ActlHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_ActlSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_ActlFGAdded_PHP_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_ActlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted;
                        Total_StdUnitCost_PHPminusActlUnitCost_PHP += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlUnitCost_PHP;
                        Total_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdUnitCost_PHP - (LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlLandedCost_PHP);
                        Total___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ += (LSP_Rpt_DM_FinishedGoodsSalesReportObjList.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted) - (LSP_Rpt_DM_FinishedGoodsSalesReportObjList.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObjList.QtyCompleted);

                        if (sheetsRowFinishedGoods < LSP_Rpt_DM_FinishedGoodsSalesReportList_FinishedGood.ToList().Count + 4)
                        {
                            FINISHEDGOODS.InsertRow((sheetsRowFinishedGoods + 1), 1);
                            FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 1, sheetsRowFinishedGoods, 100].Copy(FINISHEDGOODS.Cells[(sheetsRowFinishedGoods + 1), 1, (sheetsRowFinishedGoods + 1), 1]);
                        }
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 1].Value = DateTime.Parse(TransDate).ToString("MM/dd/yyyy");
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 1].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 2].Value = PONum;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 2].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 3].Value = CustomerName;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 3].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 4].Value = JobOrder;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 4].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 5].Value = Item;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 5].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 6].Value = ItemDesc;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 6].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 7].Value = ProductCode;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 7].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 8].Value = FamilyDesc;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 8].Style.WrapText = false;

                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 9].Value = Convert.ToDecimal(QtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 9].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 10].Value = Convert.ToDecimal(StdMatlCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 10].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 11].Value = Convert.ToDecimal(StdResinCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 11].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 12].Value = Convert.ToDecimal(StdPIProcess_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 12].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 13].Value = Convert.ToDecimal(StdHiddenProfit_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 13].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 14].Value = Convert.ToDecimal(StdSFAdded_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 14].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 15].Value = Convert.ToDecimal(StdFGAdded_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 15].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 16].Value = Convert.ToDecimal(StdUnitCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 16].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 17].Value = Convert.ToDecimal(ActlMatlUnitCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 17].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 18].Value = Convert.ToDecimal(ActlLandedCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 18].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 19].Value = Convert.ToDecimal(ActlResinCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 19].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 20].Value = Convert.ToDecimal(ActlPIProcess_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 20].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 21].Value = Convert.ToDecimal(ActlHiddenProfit_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 21].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 22].Value = Convert.ToDecimal(ActlSFAdded_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 22].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 23].Value = Convert.ToDecimal(ActlFGAdded_PHP_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 23].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 24].Value = Convert.ToDecimal(ActlUnitCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 24].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 25].Value = Convert.ToDecimal(StdUnitCost_PHPminusActlUnitCost_PHP);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 25].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 26].Value = Convert.ToDecimal(StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 26].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 27].Value = Convert.ToDecimal(__StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__);
                        FINISHEDGOODS.Cells[sheetsRowFinishedGoods, 27].Style.WrapText = false;
                        sheetsRowFinishedGoods++;
                        // }

                    }
                    #endregion
                    #region FINISHEDGOODS NOT FINISHED GOODS 
                    int sheetsRowNotFinishedGoods = sheetsRowFinishedGoods + 2;
                    foreach (var LSP_Rpt_DM_FinishedGoodsSalesReportObjList in LSP_Rpt_DM_FinishedGoodsSalesReportList_GroupByTransTpeNotFinishedGood)
                    {
                        FINISHEDGOODS.Cells["A" + (sheetsRowNotFinishedGoods - 1)].Value = LSP_Rpt_DM_FinishedGoodsSalesReportObjList.Key.ToString();
                        FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods - 1, 1].Style.WrapText = false;

                        string FGTransType = "";
                        string TransDate = "";
                        string PONum = "";
                        string CustomerName = "";
                        string JobOrder = "";
                        string Item = "";
                        string ItemDesc = "";
                        string ProductCode = "";
                        string FamilyDesc = "";
                        decimal QtyCompleted = 0;
                        decimal StdMatlCost_PHPxQtyCompleted = 0;
                        decimal StdResinCost_PHPxQtyCompleted = 0;
                        decimal StdPIProcess_PHPxQtyCompleted = 0;
                        decimal StdHiddenProfit_PHPxQtyCompleted = 0;
                        decimal StdSFAdded_PHPxQtyCompleted = 0;
                        decimal StdFGAdded_PHPxQtyCompleted = 0;
                        decimal StdUnitCost_PHPxQtyCompleted = 0;
                        decimal ActlMatlUnitCost_PHPxQtyCompleted = 0;
                        decimal ActlLandedCost_PHPxQtyCompleted = 0;
                        decimal ActlResinCost_PHPxQtyCompleted = 0;
                        decimal ActlPIProcess_PHPxQtyCompleted = 0;
                        decimal ActlHiddenProfit_PHPxQtyCompleted = 0;
                        decimal ActlSFAdded_PHPxQtyCompleted = 0;
                        decimal ActlFGAdded_PHP_PHPxQtyCompleted = 0;
                        decimal ActlUnitCost_PHPxQtyCompleted = 0;
                        decimal StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                        decimal StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                        decimal __StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ = 0;
                        foreach (var LSP_Rpt_DM_FinishedGoodsSalesReportObj in LSP_Rpt_DM_FinishedGoodsSalesReportObjList)
                        {
                            FGTransType = LSP_Rpt_DM_FinishedGoodsSalesReportObj.FGTransType;
                            TransDate = LSP_Rpt_DM_FinishedGoodsSalesReportObj.TransDate;
                            PONum = LSP_Rpt_DM_FinishedGoodsSalesReportObj.PONum;
                            CustomerName = LSP_Rpt_DM_FinishedGoodsSalesReportObj.CustomerName;
                            JobOrder = LSP_Rpt_DM_FinishedGoodsSalesReportObj.JobOrder;
                            Item = LSP_Rpt_DM_FinishedGoodsSalesReportObj.Item;
                            ItemDesc = LSP_Rpt_DM_FinishedGoodsSalesReportObj.ItemDesc;
                            ProductCode = LSP_Rpt_DM_FinishedGoodsSalesReportObj.ProductCode;
                            FamilyDesc = LSP_Rpt_DM_FinishedGoodsSalesReportObj.FamilyDesc;

                            QtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            StdMatlCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdMatlCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            StdResinCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            StdPIProcess_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            StdHiddenProfit_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            StdSFAdded_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            StdFGAdded_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            StdUnitCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            ActlMatlUnitCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlMatlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            ActlLandedCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlLandedCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            ActlResinCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            ActlPIProcess_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            ActlHiddenProfit_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            ActlSFAdded_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            ActlFGAdded_PHP_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            ActlUnitCost_PHPxQtyCompleted = LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            StdUnitCost_PHPminusActlUnitCost_PHP = LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlUnitCost_PHP;
                            StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdUnitCost_PHP - (LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlLandedCost_PHP);
                            __StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ = (LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted) - (LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted);

                            Total_QtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_StdMatlCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdMatlCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_StdResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_StdPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_StdHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_StdSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_StdFGAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_StdUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_ActlMatlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlMatlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_ActlLandedCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlLandedCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_ActlResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_ActlPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_ActlHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_ActlSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_ActlFGAdded_PHP_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_ActlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted;
                            Total_StdUnitCost_PHPminusActlUnitCost_PHP += LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlUnitCost_PHP;
                            Total_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdUnitCost_PHP - (LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlLandedCost_PHP);
                            Total___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ += (LSP_Rpt_DM_FinishedGoodsSalesReportObj.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted) - (LSP_Rpt_DM_FinishedGoodsSalesReportObj.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj.QtyCompleted);

                            if (sheetsRowNotFinishedGoods < CurrentDataCount)
                            {
                                FINISHEDGOODS.InsertRow((sheetsRowNotFinishedGoods + 1), 1);
                                FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 1, sheetsRowNotFinishedGoods, 100].Copy(FINISHEDGOODS.Cells[(sheetsRowNotFinishedGoods + 1), 1, (sheetsRowNotFinishedGoods + 1), 1]);
                            }
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 1].Value = DateTime.Parse(TransDate).ToString("MM/dd/yyyy");
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 1].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 2].Value = PONum;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 2].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 3].Value = CustomerName;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 3].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 4].Value = JobOrder;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 4].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 5].Value = Item;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 5].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 6].Value = ItemDesc;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 6].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 7].Value = ProductCode;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 7].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 8].Value = FamilyDesc;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 8].Style.WrapText = false;

                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 9].Value = Convert.ToDecimal(QtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 9].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 10].Value = Convert.ToDecimal(StdMatlCost_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 10].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 11].Value = Convert.ToDecimal(StdResinCost_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 11].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 12].Value = Convert.ToDecimal(StdPIProcess_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 12].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 13].Value = Convert.ToDecimal(StdHiddenProfit_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 13].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 14].Value = Convert.ToDecimal(StdSFAdded_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 14].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 15].Value = Convert.ToDecimal(StdFGAdded_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 15].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 16].Value = Convert.ToDecimal(StdUnitCost_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 16].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 17].Value = Convert.ToDecimal(ActlMatlUnitCost_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 17].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 18].Value = Convert.ToDecimal(ActlLandedCost_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 18].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 19].Value = Convert.ToDecimal(ActlResinCost_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 19].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 20].Value = Convert.ToDecimal(ActlPIProcess_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 20].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 21].Value = Convert.ToDecimal(ActlHiddenProfit_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 21].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 22].Value = Convert.ToDecimal(ActlSFAdded_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 22].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 23].Value = Convert.ToDecimal(ActlFGAdded_PHP_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 23].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 24].Value = Convert.ToDecimal(ActlUnitCost_PHPxQtyCompleted);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 24].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 25].Value = Convert.ToDecimal(StdUnitCost_PHPminusActlUnitCost_PHP);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 25].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 26].Value = Convert.ToDecimal(StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 26].Style.WrapText = false;
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 27].Value = Convert.ToDecimal(__StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__);
                            FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 27].Style.WrapText = false;
                            sheetsRowNotFinishedGoods++;
                        }

                    }
                    #endregion
                    #region FINISHED GOODS TOTAL

                    sheetsRowNotFinishedGoods++;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 9].Value = Convert.ToDecimal(Total_QtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 9].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 10].Value = Convert.ToDecimal(Total_StdMatlCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 10].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 11].Value = Convert.ToDecimal(Total_StdResinCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 11].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 12].Value = Convert.ToDecimal(Total_StdPIProcess_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 12].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 13].Value = Convert.ToDecimal(Total_StdHiddenProfit_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 13].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 14].Value = Convert.ToDecimal(Total_StdSFAdded_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 14].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 15].Value = Convert.ToDecimal(Total_StdFGAdded_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 15].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 16].Value = Convert.ToDecimal(Total_StdUnitCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 16].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 17].Value = Convert.ToDecimal(Total_ActlMatlUnitCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 17].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 18].Value = Convert.ToDecimal(Total_ActlLandedCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 18].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 19].Value = Convert.ToDecimal(Total_ActlResinCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 19].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 20].Value = Convert.ToDecimal(Total_ActlPIProcess_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 20].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 21].Value = Convert.ToDecimal(Total_ActlHiddenProfit_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 21].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 22].Value = Convert.ToDecimal(Total_ActlSFAdded_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 22].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 23].Value = Convert.ToDecimal(Total_ActlFGAdded_PHP_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 23].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 24].Value = Convert.ToDecimal(Total_ActlUnitCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 24].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 25].Value = Convert.ToDecimal(Total_StdUnitCost_PHPminusActlUnitCost_PHP);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 25].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 26].Value = Convert.ToDecimal(Total_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 26].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 27].Value = Convert.ToDecimal(Total___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__);
                    FINISHEDGOODS.Cells[sheetsRowNotFinishedGoods, 27].Style.WrapText = false;
                    CurrentDataCount += 2;
                    #endregion

                    #region FAMILY TOTAL
                    var LSP_Rpt_DM_FinishedGoodsSalesReportList_GroubyFamily = LSP_Rpt_DM_FinishedGoodsSalesReportList
                    .OrderBy(x => x.Family)
                    .GroupBy(x => x.FamilyDesc)
                    .ToList();
                    int sheetsRowFamily = sheetsRowNotFinishedGoods + 4;
                    decimal GrandTotalFamily_QtyCompleted = 0;
                    decimal GrandTotalFamily_StdMatlCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_StdResinCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_StdPIProcess_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_StdHiddenProfit_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_StdSFAdded_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_StdFGAdded_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_StdUnitCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_ActlMatlUnitCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_ActlLandedCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_ActlResinCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_ActlPIProcess_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_ActlHiddenProfit_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_ActlSFAdded_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_ActlFGAdded_PHP_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_ActlUnitCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                    decimal GrandTotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                    decimal GrandTotalFamily___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ = 0;
                    CurrentDataCount = CurrentDataCount + LSP_Rpt_DM_FinishedGoodsSalesReportList_GroubyFamily.Count + 3;
                    foreach (var LSP_Rpt_DM_FinishedGoodsSalesReportListObj_GroubyFamily in LSP_Rpt_DM_FinishedGoodsSalesReportList_GroubyFamily)
                    {
                        string Family = "";
                        string FamilyDesc = "";

                        decimal TotalFamily_QtyCompleted = 0;
                        decimal TotalFamily_StdMatlCost_PHPxQtyCompleted = 0;
                        decimal TotalFamily_StdResinCost_PHPxQtyCompleted = 0;
                        decimal TotalFamily_StdPIProcess_PHPxQtyCompleted = 0;
                        decimal TotalFamily_StdHiddenProfit_PHPxQtyCompleted = 0;
                        decimal TotalFamily_StdSFAdded_PHPxQtyCompleted = 0;
                        decimal TotalFamily_StdFGAdded_PHPxQtyCompleted = 0;
                        decimal TotalFamily_StdUnitCost_PHPxQtyCompleted = 0;
                        decimal TotalFamily_ActlMatlUnitCost_PHPxQtyCompleted = 0;
                        decimal TotalFamily_ActlLandedCost_PHPxQtyCompleted = 0;
                        decimal TotalFamily_ActlResinCost_PHPxQtyCompleted = 0;
                        decimal TotalFamily_ActlPIProcess_PHPxQtyCompleted = 0;
                        decimal TotalFamily_ActlHiddenProfit_PHPxQtyCompleted = 0;
                        decimal TotalFamily_ActlSFAdded_PHPxQtyCompleted = 0;
                        decimal TotalFamily_ActlFGAdded_PHP_PHPxQtyCompleted = 0;
                        decimal TotalFamily_ActlUnitCost_PHPxQtyCompleted = 0;
                        decimal TotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                        decimal TotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                        decimal TotalFamily___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ = 0;
                        foreach (var LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily in LSP_Rpt_DM_FinishedGoodsSalesReportListObj_GroubyFamily)
                        {

                            Family = LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.Family;
                            FamilyDesc = LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.FamilyDesc;

                            TotalFamily_QtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_StdMatlCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdMatlCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_StdResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_StdPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_StdHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_StdSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_StdFGAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_StdUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_ActlMatlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlMatlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_ActlLandedCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlLandedCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_ActlResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_ActlPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_ActlHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_ActlSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_ActlFGAdded_PHP_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_ActlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            TotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlUnitCost_PHP;
                            TotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdUnitCost_PHP - (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlLandedCost_PHP);
                            TotalFamily___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ += (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted) - (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted);

                            GrandTotalFamily_QtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_StdMatlCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdMatlCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_StdResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_StdPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_StdHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_StdSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_StdFGAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_StdUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_ActlMatlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlMatlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_ActlLandedCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlLandedCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_ActlResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_ActlPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_ActlHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_ActlSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_ActlFGAdded_PHP_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_ActlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted;
                            GrandTotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlUnitCost_PHP;
                            GrandTotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdUnitCost_PHP - (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlLandedCost_PHP);
                            GrandTotalFamily___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ += (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted) - (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyFamily.QtyCompleted);
                        }

                        if (sheetsRowFamily < CurrentDataCount)
                        {
                            FINISHEDGOODS.InsertRow((sheetsRowFamily + 1), 1);
                            FINISHEDGOODS.Cells[sheetsRowFamily, 1, sheetsRowFamily, 100].Copy(FINISHEDGOODS.Cells[(sheetsRowFamily + 1), 1, (sheetsRowFamily + 1), 1]);
                        }

                        FINISHEDGOODS.Cells[sheetsRowFamily, 7].Value = Family;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 7].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 8].Value = FamilyDesc;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 8].Style.WrapText = false;

                        FINISHEDGOODS.Cells[sheetsRowFamily, 9].Value = Convert.ToDecimal(TotalFamily_QtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 9].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 10].Value = Convert.ToDecimal(TotalFamily_StdMatlCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 10].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 11].Value = Convert.ToDecimal(TotalFamily_StdResinCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 11].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 12].Value = Convert.ToDecimal(TotalFamily_StdPIProcess_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 12].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 13].Value = Convert.ToDecimal(TotalFamily_StdHiddenProfit_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 13].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 14].Value = Convert.ToDecimal(TotalFamily_StdSFAdded_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 14].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 15].Value = Convert.ToDecimal(TotalFamily_StdFGAdded_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 15].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 16].Value = Convert.ToDecimal(TotalFamily_StdUnitCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 16].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 17].Value = Convert.ToDecimal(TotalFamily_ActlMatlUnitCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 17].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 18].Value = Convert.ToDecimal(TotalFamily_ActlLandedCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 18].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 19].Value = Convert.ToDecimal(TotalFamily_ActlResinCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 19].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 20].Value = Convert.ToDecimal(TotalFamily_ActlPIProcess_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 20].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 21].Value = Convert.ToDecimal(TotalFamily_ActlHiddenProfit_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 21].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 22].Value = Convert.ToDecimal(TotalFamily_ActlSFAdded_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 22].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 23].Value = Convert.ToDecimal(TotalFamily_ActlFGAdded_PHP_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 23].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 24].Value = Convert.ToDecimal(TotalFamily_ActlUnitCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 24].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 25].Value = Convert.ToDecimal(TotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 25].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 26].Value = Convert.ToDecimal(TotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 26].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowFamily, 27].Value = Convert.ToDecimal(TotalFamily___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__);
                        FINISHEDGOODS.Cells[sheetsRowFamily, 27].Style.WrapText = false;
                        sheetsRowFamily++;
                    }
                    FINISHEDGOODS.Cells[sheetsRowFamily, 9].Value = Convert.ToDecimal(GrandTotalFamily_QtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 9].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 10].Value = Convert.ToDecimal(GrandTotalFamily_StdMatlCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 10].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 11].Value = Convert.ToDecimal(GrandTotalFamily_StdResinCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 11].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 12].Value = Convert.ToDecimal(GrandTotalFamily_StdPIProcess_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 12].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 13].Value = Convert.ToDecimal(GrandTotalFamily_StdHiddenProfit_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 13].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 14].Value = Convert.ToDecimal(GrandTotalFamily_StdSFAdded_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 14].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 15].Value = Convert.ToDecimal(GrandTotalFamily_StdFGAdded_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 15].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 16].Value = Convert.ToDecimal(GrandTotalFamily_StdUnitCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 16].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 17].Value = Convert.ToDecimal(GrandTotalFamily_ActlMatlUnitCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 17].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 18].Value = Convert.ToDecimal(GrandTotalFamily_ActlLandedCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 18].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 19].Value = Convert.ToDecimal(GrandTotalFamily_ActlResinCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 19].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 20].Value = Convert.ToDecimal(GrandTotalFamily_ActlPIProcess_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 20].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 21].Value = Convert.ToDecimal(GrandTotalFamily_ActlHiddenProfit_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 21].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 22].Value = Convert.ToDecimal(GrandTotalFamily_ActlSFAdded_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 22].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 23].Value = Convert.ToDecimal(GrandTotalFamily_ActlFGAdded_PHP_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 23].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 24].Value = Convert.ToDecimal(GrandTotalFamily_ActlUnitCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 24].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 25].Value = Convert.ToDecimal(GrandTotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 25].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 26].Value = Convert.ToDecimal(GrandTotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 26].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowFamily, 27].Value = Convert.ToDecimal(GrandTotalFamily___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__);
                    FINISHEDGOODS.Cells[sheetsRowFamily, 27].Style.WrapText = false;
                    #endregion

                    #region PRODUCT CODE TOTAL
                    var LSP_Rpt_DM_FinishedGoodsSalesReportList_GroubyProductCode = LSP_Rpt_DM_FinishedGoodsSalesReportList
                    .OrderBy(x => x.ProductCode)
                    .GroupBy(x => x.ProductCode)
                    .ToList();
                    int sheetsRowProductCode = sheetsRowFamily + 4;
                    decimal GrandTotalProductCode_QtyCompleted = 0;
                    decimal GrandTotalProductCode_StdMatlCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_StdResinCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_StdPIProcess_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_StdHiddenProfit_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_StdSFAdded_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_StdFGAdded_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_StdUnitCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_ActlMatlUnitCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_ActlLandedCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_ActlResinCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_ActlPIProcess_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_ActlHiddenProfit_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_ActlSFAdded_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_ActlFGAdded_PHP_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_ActlUnitCost_PHPxQtyCompleted = 0;
                    decimal GrandTotalProductCode_StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                    decimal GrandTotalProductCode_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                    decimal GrandTotalProductCode___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ = 0;
                    CurrentDataCount = CurrentDataCount + LSP_Rpt_DM_FinishedGoodsSalesReportList_GroubyProductCode.Count + 4;
                    foreach (var LSP_Rpt_DM_FinishedGoodsSalesReportListObj_GroubyProductCode in LSP_Rpt_DM_FinishedGoodsSalesReportList_GroubyProductCode)
                    {
                        string ProductCode = "";

                        decimal TotalProductCode_QtyCompleted = 0;
                        decimal TotalProductCode_StdMatlCost_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_StdResinCost_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_StdPIProcess_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_StdHiddenProfit_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_StdSFAdded_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_StdFGAdded_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_StdUnitCost_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_ActlMatlUnitCost_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_ActlLandedCost_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_ActlResinCost_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_ActlPIProcess_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_ActlHiddenProfit_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_ActlSFAdded_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_ActlFGAdded_PHP_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_ActlUnitCost_PHPxQtyCompleted = 0;
                        decimal TotalProductCode_StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                        decimal TotalProductCode_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                        decimal TotalProductCode___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ = 0;
                        foreach (var LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode in LSP_Rpt_DM_FinishedGoodsSalesReportListObj_GroubyProductCode)
                        {

                            ProductCode = LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ProductCode;

                            TotalProductCode_QtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_StdMatlCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdMatlCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_StdResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_StdPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_StdHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_StdSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_StdFGAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_StdUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_ActlMatlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlMatlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_ActlLandedCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlLandedCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_ActlResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_ActlPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_ActlHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_ActlSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_ActlFGAdded_PHP_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_ActlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            TotalProductCode_StdUnitCost_PHPminusActlUnitCost_PHP += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlUnitCost_PHP;
                            TotalProductCode_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdUnitCost_PHP - (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlLandedCost_PHP);
                            TotalProductCode___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ += (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted) - (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted);

                            GrandTotalProductCode_QtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_StdMatlCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdMatlCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_StdResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_StdPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_StdHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_StdSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_StdFGAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_StdUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_ActlMatlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlMatlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_ActlLandedCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlLandedCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_ActlResinCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlResinCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_ActlPIProcess_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlPIProcess_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_ActlHiddenProfit_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlHiddenProfit_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_ActlSFAdded_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlSFAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_ActlFGAdded_PHP_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlFGAdded_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_ActlUnitCost_PHPxQtyCompleted += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted;
                            GrandTotalProductCode_StdUnitCost_PHPminusActlUnitCost_PHP += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlUnitCost_PHP;
                            GrandTotalProductCode_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdUnitCost_PHP - (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlUnitCost_PHP - LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlLandedCost_PHP);
                            GrandTotalProductCode___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__ += (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.StdUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted) - (LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.ActlUnitCost_PHP * LSP_Rpt_DM_FinishedGoodsSalesReportObj_GroubyProductCode.QtyCompleted);
                        }

                        if (sheetsRowProductCode < CurrentDataCount)
                        {
                            FINISHEDGOODS.InsertRow((sheetsRowProductCode + 1), 1);
                            FINISHEDGOODS.Cells[sheetsRowProductCode, 1, sheetsRowProductCode, 100].Copy(FINISHEDGOODS.Cells[(sheetsRowProductCode + 1), 1, (sheetsRowProductCode + 1), 1]);
                        }

                        FINISHEDGOODS.Cells[sheetsRowProductCode, 8].Value = ProductCode.Replace("FG-", "");
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 8].Style.WrapText = false;

                        FINISHEDGOODS.Cells[sheetsRowProductCode, 9].Value = Convert.ToDecimal(TotalProductCode_QtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 9].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 10].Value = Convert.ToDecimal(TotalProductCode_StdMatlCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 10].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 11].Value = Convert.ToDecimal(TotalProductCode_StdResinCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 11].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 12].Value = Convert.ToDecimal(TotalProductCode_StdPIProcess_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 12].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 13].Value = Convert.ToDecimal(TotalProductCode_StdHiddenProfit_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 13].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 14].Value = Convert.ToDecimal(TotalProductCode_StdSFAdded_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 14].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 15].Value = Convert.ToDecimal(TotalProductCode_StdFGAdded_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 15].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 16].Value = Convert.ToDecimal(TotalProductCode_StdUnitCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 16].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 17].Value = Convert.ToDecimal(TotalProductCode_ActlMatlUnitCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 17].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 18].Value = Convert.ToDecimal(TotalProductCode_ActlLandedCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 18].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 19].Value = Convert.ToDecimal(TotalProductCode_ActlResinCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 19].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 20].Value = Convert.ToDecimal(TotalProductCode_ActlPIProcess_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 20].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 21].Value = Convert.ToDecimal(TotalProductCode_ActlHiddenProfit_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 21].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 22].Value = Convert.ToDecimal(TotalProductCode_ActlSFAdded_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 22].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 23].Value = Convert.ToDecimal(TotalProductCode_ActlFGAdded_PHP_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 23].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 24].Value = Convert.ToDecimal(TotalProductCode_ActlUnitCost_PHPxQtyCompleted);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 24].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 25].Value = Convert.ToDecimal(TotalProductCode_StdUnitCost_PHPminusActlUnitCost_PHP);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 25].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 26].Value = Convert.ToDecimal(TotalProductCode_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 26].Style.WrapText = false;
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 27].Value = Convert.ToDecimal(TotalProductCode___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__);
                        FINISHEDGOODS.Cells[sheetsRowProductCode, 27].Style.WrapText = false;
                        sheetsRowProductCode++;
                    }
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 9].Value = Convert.ToDecimal(GrandTotalProductCode_QtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 9].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 10].Value = Convert.ToDecimal(GrandTotalProductCode_StdMatlCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 10].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 11].Value = Convert.ToDecimal(GrandTotalProductCode_StdResinCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 11].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 12].Value = Convert.ToDecimal(GrandTotalProductCode_StdPIProcess_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 12].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 13].Value = Convert.ToDecimal(GrandTotalProductCode_StdHiddenProfit_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 13].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 14].Value = Convert.ToDecimal(GrandTotalProductCode_StdSFAdded_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 14].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 15].Value = Convert.ToDecimal(GrandTotalProductCode_StdFGAdded_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 15].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 16].Value = Convert.ToDecimal(GrandTotalProductCode_StdUnitCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 16].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 17].Value = Convert.ToDecimal(GrandTotalProductCode_ActlMatlUnitCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 17].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 18].Value = Convert.ToDecimal(GrandTotalProductCode_ActlLandedCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 18].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 19].Value = Convert.ToDecimal(GrandTotalProductCode_ActlResinCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 19].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 20].Value = Convert.ToDecimal(GrandTotalProductCode_ActlPIProcess_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 20].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 21].Value = Convert.ToDecimal(GrandTotalProductCode_ActlHiddenProfit_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 21].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 22].Value = Convert.ToDecimal(GrandTotalProductCode_ActlSFAdded_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 22].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 23].Value = Convert.ToDecimal(GrandTotalProductCode_ActlFGAdded_PHP_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 23].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 24].Value = Convert.ToDecimal(GrandTotalProductCode_ActlUnitCost_PHPxQtyCompleted);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 24].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 25].Value = Convert.ToDecimal(GrandTotalProductCode_StdUnitCost_PHPminusActlUnitCost_PHP);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 25].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 26].Value = Convert.ToDecimal(GrandTotalProductCode_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 26].Style.WrapText = false;
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 27].Value = Convert.ToDecimal(GrandTotalProductCode___StdUnitCost_PHPxQtyCompleted__minus__ActlUnitCost_PHPxQtyCompleted__);
                    FINISHEDGOODS.Cells[sheetsRowProductCode, 27].Style.WrapText = false;
                    #endregion
                    #endregion
                    #region Sales and Sample JO

                    List<FinishedGoods_Sales_SampleJO> FinishedGoods_Sales_SampleJOList = new List<FinishedGoods_Sales_SampleJO>();
                    using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LSPI803_App"].ConnectionString.ToString()))
                    {
                        conn.Open();
                        using (SqlCommand cmdSql = conn.CreateCommand())
                        {

                            cmdSql.CommandType = CommandType.StoredProcedure;
                            cmdSql.CommandText = "LSP_Rpt_NewDM_SalesAndSampleJOReportSp";
                            cmdSql.Parameters.AddWithValue("@StartDate", StartDate);
                            cmdSql.Parameters.AddWithValue("@EndDate", EndDate);
                            cmdSql.CommandTimeout = 0;
                            using (SqlDataReader sdr = cmdSql.ExecuteReader())
                            {
                                while (sdr.Read())
                                {
                                    FinishedGoods_Sales_SampleJOList.Add(new FinishedGoods_Sales_SampleJO
                                    {
                                        TransDate = sdr["TransDate"] == null ? "" : sdr["TransDate"].ToString(),
                                        Item = sdr["Item"] == null ? "" : sdr["Item"].ToString(),
                                        ItemDesc = sdr["ItemDesc"] == null ? "" : sdr["ItemDesc"].ToString(),
                                        ProductCode = sdr["ProductCode"] == null ? "" : sdr["ProductCode"].ToString(),
                                        Family = sdr["Family"] == null ? "" : sdr["Family"].ToString(),
                                        FamilyDesc = sdr["FamilyDesc"] == null ? "" : sdr["FamilyDesc"].ToString(),
                                        PONum = sdr["PONum"] == null ? "" : sdr["PONum"].ToString(),
                                        LotNo = sdr["LotNo"] == null ? "" : sdr["LotNo"].ToString(),
                                        JobOrder = sdr["JobOrder"] == null ? "" : sdr["JobOrder"].ToString(),
                                        JobSuffix = sdr["JobSuffix"] == null ? "" : sdr["JobSuffix"].ToString(),
                                        CONum = sdr["CONum"] == null ? "" : sdr["CONum"].ToString(),
                                        COLine = sdr["COLine"] == null ? "" : sdr["COLine"].ToString(),
                                        CustNum = sdr["CustNum"] == null ? "" : sdr["CustNum"].ToString(),
                                        ShipToCust = sdr["ShipToCust"] == null ? "" : sdr["ShipToCust"].ToString(),
                                        CustomerName = sdr["CustomerName"] == null ? "" : sdr["CustomerName"].ToString(),
                                        QtyShipped = sdr["QtyShipped"] == null ? 0 : Convert.ToDecimal(sdr["QtyShipped"]),
                                        SalesPrice = sdr["SalesPrice"] == null ? 0 : Convert.ToDecimal(sdr["SalesPrice"]),
                                        SalesPriceConv = sdr["SalesPriceConv"] == null ? 0 : Convert.ToDecimal(sdr["SalesPriceConv"]),
                                        StdMatlCost_PHP = sdr["StdMatlCost_PHP"] == null ? 0 : Convert.ToDecimal(sdr["StdMatlCost_PHP"]),
                                        StdLandedCost_PHP = sdr["StdLandedCost_PHP"] == null ? 0 : Convert.ToDecimal(sdr["StdLandedCost_PHP"]),
                                        StdResinCost_PHP = sdr["StdResinCost_PHP"] == null ? 0 : Convert.ToDecimal(sdr["StdResinCost_PHP"]),
                                        StdPIProcess_PHP = sdr["StdPIProcess_PHP"] == null ? 0 : Convert.ToDecimal(sdr["StdPIProcess_PHP"]),
                                        StdHiddenProfit_PHP = sdr["StdHiddenProfit_PHP"] == null ? 0 : Convert.ToDecimal(sdr["StdHiddenProfit_PHP"]),
                                        StdSFAdded_PHP = sdr["StdSFAdded_PHP"] == null ? 0 : Convert.ToDecimal(sdr["StdSFAdded_PHP"]),
                                        StdFGAdded_PHP = sdr["StdFGAdded_PHP"] == null ? 0 : Convert.ToDecimal(sdr["StdFGAdded_PHP"]),
                                        StdUnitCost_PHP = sdr["StdUnitCost_PHP"] == null ? 0 : Convert.ToDecimal(sdr["StdUnitCost_PHP"]),
                                        ActlMatlUnitCost_PHP = sdr["ActlMatlUnitCost_PHP"] == null ? 0 : Convert.ToDecimal(sdr["ActlMatlUnitCost_PHP"]),
                                        ActlLandedCost_PHP = sdr["ActlLandedCost_PHP"] == null ? 0 : Convert.ToDecimal(sdr["ActlLandedCost_PHP"]),
                                        ActlResinCost_PHP = sdr["ActlResinCost_PHP"] == null ? 0 : Convert.ToDecimal(sdr["ActlResinCost_PHP"]),
                                        ActlPIProcess_PHP = sdr["ActlPIProcess_PHP"] == null ? 0 : Convert.ToDecimal(sdr["ActlPIProcess_PHP"]),
                                        ActlHiddenProfit_PHP = sdr["ActlHiddenProfit_PHP"] == null ? 0 : Convert.ToDecimal(sdr["ActlHiddenProfit_PHP"]),
                                        ActlSFAdded_PHP = sdr["ActlSFAdded_PHP"] == null ? 0 : Convert.ToDecimal(sdr["ActlSFAdded_PHP"]),
                                        ActlFGAdded_PHP = sdr["ActlFGAdded_PHP"] == null ? 0 : Convert.ToDecimal(sdr["ActlFGAdded_PHP"]),
                                        ActlUnitCost_PHP = sdr["ActlUnitCost_PHP"] == null ? 0 : Convert.ToDecimal(sdr["ActlUnitCost_PHP"]),
                                        ShipCategory = sdr["ShipCategory"] == null ? "" : sdr["ShipCategory"].ToString(),
                                        Recoverable = sdr["Recoverable"] == null ? "" : sdr["Recoverable"].ToString(),
                                        JobRemarks = sdr["JobRemarks"] == null ? "" : sdr["JobRemarks"].ToString(),
                                    });
                                }

                            }
                        }
                        conn.Close();

                        var LSP_Rpt_DM_FinishedGoodsSalesReportList_GroupByShipCategory = FinishedGoods_Sales_SampleJOList
                            .OrderBy(x => x.TransDate)
                            .ThenBy(x => x.ShipCategory)
                            .GroupBy(x => x.ShipCategory)
                            .ToList();
                        int startPos = 4;
                        foreach (var FinishedGoodsGroupSales_SampleJOList in LSP_Rpt_DM_FinishedGoodsSalesReportList_GroupByShipCategory)
                        {
                            if (FinishedGoodsGroupSales_SampleJOList.Key.ToString() != "Sales")
                            {
                                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Copy("Sales", FinishedGoodsGroupSales_SampleJOList.Key.ToString());
                                excelPackage.Workbook.Worksheets.MoveBefore(startPos, 3);
                                startPos++;
                                worksheet.Cells["A1"].Value = FinishedGoodsGroupSales_SampleJOList.Key.ToString() + " Report";
                                worksheet.Cells["A1"].Style.WrapText = false;
                            }
                        }


                        foreach (var FinishedGoodsGroupSales_SampleJOList in LSP_Rpt_DM_FinishedGoodsSalesReportList_GroupByShipCategory)
                        {
                            ExcelWorksheet SalesSheets = excelPackage.Workbook.Worksheets[FinishedGoodsGroupSales_SampleJOList.Key.ToString()];
                            decimal Sales_Total_QtyShipped = 0;
                            decimal Sales_Total_SalesPricexQtyShipped = 0;
                            decimal Sales_Total_SalesPriceConvxQtyShipped = 0;
                            decimal Sales_Total_StdMatlCost_PHPxQtyShipped = 0;
                            decimal Sales_Total_StdResinCost_PHPxQtyShipped = 0;
                            decimal Sales_Total_StdPIProcess_PHPxQtyShipped = 0;
                            decimal Sales_Total_StdHiddenProfit_PHPxQtyShipped = 0;
                            decimal Sales_Total_StdSFAdded_PHPxQtyShipped = 0;
                            decimal Sales_Total_StdFGAdded_PHPxQtyShipped = 0;
                            decimal Sales_Total_StdUnitCost_PHPxQtyShipped = 0;
                            decimal Sales_Total_ActlMatlUnitCost_PHPxQtyShipped = 0;
                            decimal Sales_Total_ActlLandedCost_PHPxQtyShipped = 0;
                            decimal Sales_Total_ActlResinCost_PHPxQtyShipped = 0;
                            decimal Sales_Total_ActlPIProcess_PHPxQtyShipped = 0;
                            decimal Sales_Total_ActdHiddenProfit_PHPxQtyShipped = 0;
                            decimal Sales_Total_ActlSFAdded_PHPxQtyShipped = 0;
                            decimal Sales_Total_ActlFGAdded_PHPxQtyShipped = 0;
                            decimal Sales_Total_ActlUnitCost_PHPxQtyShipped = 0;
                            decimal Sales_Total_StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                            decimal Sales_Total_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                            decimal Sales_Total___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ = 0;
                            int salesSheetsRow = 5;
                            #region Details
                            foreach (var FinishedGoodsGroupSales_SampleJOObj in FinishedGoodsGroupSales_SampleJOList)
                            {

                                if (salesSheetsRow < FinishedGoodsGroupSales_SampleJOList.ToList().Count + 4)
                                {
                                    SalesSheets.InsertRow((salesSheetsRow + 1), 1);
                                    SalesSheets.Cells[salesSheetsRow, 1, salesSheetsRow, 100].Copy(SalesSheets.Cells[(salesSheetsRow + 1), 1, (salesSheetsRow + 1), 1]);
                                }
                                SalesSheets.Cells[salesSheetsRow, 1].Value = DateTime.Parse(FinishedGoodsGroupSales_SampleJOObj.TransDate).ToString("MM/dd/yyyy");
                                SalesSheets.Cells[salesSheetsRow, 1].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 2].Value = FinishedGoodsGroupSales_SampleJOObj.PONum;
                                SalesSheets.Cells[salesSheetsRow, 2].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 3].Value = FinishedGoodsGroupSales_SampleJOObj.CustomerName;
                                SalesSheets.Cells[salesSheetsRow, 3].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 4].Value = FinishedGoodsGroupSales_SampleJOObj.JobOrder + FinishedGoodsGroupSales_SampleJOObj.JobSuffix;
                                SalesSheets.Cells[salesSheetsRow, 4].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 5].Value = FinishedGoodsGroupSales_SampleJOObj.Item;
                                SalesSheets.Cells[salesSheetsRow, 5].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 6].Value = FinishedGoodsGroupSales_SampleJOObj.ItemDesc;
                                SalesSheets.Cells[salesSheetsRow, 6].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 7].Value = FinishedGoodsGroupSales_SampleJOObj.ProductCode;
                                SalesSheets.Cells[salesSheetsRow, 7].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 8].Value = FinishedGoodsGroupSales_SampleJOObj.FamilyDesc;
                                SalesSheets.Cells[salesSheetsRow, 8].Style.WrapText = false;

                                decimal QtyShipped = FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal SalesPricexQtyShipped = FinishedGoodsGroupSales_SampleJOObj.SalesPrice * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal SalesPriceConvxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.SalesPriceConv * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal StdMatlCost_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.StdMatlCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal StdResinCost_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.StdResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal StdPIProcess_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.StdPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal StdHiddenProfit_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.StdHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal StdSFAdded_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.StdSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal StdFGAdded_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.StdFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal StdUnitCost_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal ActlMatlUnitCost_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.ActlMatlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal ActlLandedCost_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.ActlLandedCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal ActlResinCost_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.ActlResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal ActlPIProcess_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.ActlPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal ActdHiddenProfit_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.ActlHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal ActlSFAdded_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.ActlSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal ActlFGAdded_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.ActlFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal ActlUnitCost_PHPxQtyShipped = FinishedGoodsGroupSales_SampleJOObj.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                decimal StdUnitCost_PHPminusActlUnitCost_PHP = FinishedGoodsGroupSales_SampleJOObj.StdUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj.ActlUnitCost_PHP;
                                decimal StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = FinishedGoodsGroupSales_SampleJOObj.StdUnitCost_PHP - (FinishedGoodsGroupSales_SampleJOObj.ActlUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj.ActlLandedCost_PHP);
                                decimal __StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ = (FinishedGoodsGroupSales_SampleJOObj.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped) - (FinishedGoodsGroupSales_SampleJOObj.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped);

                                Sales_Total_QtyShipped += FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_SalesPricexQtyShipped += FinishedGoodsGroupSales_SampleJOObj.SalesPrice * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_SalesPriceConvxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.SalesPriceConv * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_StdMatlCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.StdMatlCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_StdResinCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.StdResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_StdPIProcess_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.StdPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_StdHiddenProfit_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.StdHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_StdSFAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.StdSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_StdFGAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.StdFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_StdUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_ActlMatlUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.ActlMatlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_ActlLandedCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.ActlLandedCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_ActlResinCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.ActlResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_ActlPIProcess_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.ActlPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_ActdHiddenProfit_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.ActlHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_ActlSFAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.ActlSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_ActlFGAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.ActlFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_ActlUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped;
                                Sales_Total_StdUnitCost_PHPminusActlUnitCost_PHP += FinishedGoodsGroupSales_SampleJOObj.StdUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj.ActlUnitCost_PHP;
                                Sales_Total_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += FinishedGoodsGroupSales_SampleJOObj.StdUnitCost_PHP - (FinishedGoodsGroupSales_SampleJOObj.ActlUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj.ActlLandedCost_PHP);
                                Sales_Total___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ += (FinishedGoodsGroupSales_SampleJOObj.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped) - (FinishedGoodsGroupSales_SampleJOObj.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj.QtyShipped);

                                string Recoverable = "";
                                string JobRemarks = "";
                                if (FinishedGoodsGroupSales_SampleJOObj.ShipCategory != "Sales")
                                {
                                    JobRemarks = FinishedGoodsGroupSales_SampleJOObj.JobRemarks;
                                    if (FinishedGoodsGroupSales_SampleJOObj.Recoverable == "0")
                                    {
                                        Recoverable = "YES";
                                    }
                                    else
                                    {
                                        Recoverable = "NO";
                                    }
                                }
                                SalesSheets.Cells[salesSheetsRow, 9].Value = Convert.ToDecimal(QtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 9].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 10].Value = Convert.ToDecimal(SalesPricexQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 10].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 11].Value = Convert.ToDecimal(SalesPriceConvxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 11].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 12].Value = Convert.ToDecimal(StdMatlCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 12].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 13].Value = Convert.ToDecimal(StdResinCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 13].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 14].Value = Convert.ToDecimal(StdPIProcess_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 14].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 15].Value = Convert.ToDecimal(StdHiddenProfit_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 15].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 16].Value = Convert.ToDecimal(StdSFAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 16].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 17].Value = Convert.ToDecimal(StdFGAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 17].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 18].Value = Convert.ToDecimal(StdUnitCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 18].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 19].Value = Convert.ToDecimal(ActlMatlUnitCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 19].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 20].Value = Convert.ToDecimal(ActlLandedCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 20].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 21].Value = Convert.ToDecimal(ActlResinCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 21].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 22].Value = Convert.ToDecimal(ActlPIProcess_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 22].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 23].Value = Convert.ToDecimal(ActdHiddenProfit_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 23].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 24].Value = Convert.ToDecimal(ActlSFAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 24].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 25].Value = Convert.ToDecimal(ActlFGAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 25].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 26].Value = Convert.ToDecimal(ActlUnitCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 26].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 27].Value = Convert.ToDecimal(StdUnitCost_PHPminusActlUnitCost_PHP);
                                SalesSheets.Cells[salesSheetsRow, 27].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 28].Value = Convert.ToDecimal(StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                                SalesSheets.Cells[salesSheetsRow, 28].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 29].Value = Convert.ToDecimal(__StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__);
                                SalesSheets.Cells[salesSheetsRow, 29].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 30].Value = Recoverable;
                                SalesSheets.Cells[salesSheetsRow, 30].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 31].Value = JobRemarks;
                                SalesSheets.Cells[salesSheetsRow, 31].Style.WrapText = false;
                                salesSheetsRow++;
                            }
                            SalesSheets.Cells[salesSheetsRow, 9].Value = Convert.ToDecimal(Sales_Total_QtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 9].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 10].Value = Convert.ToDecimal(Sales_Total_SalesPricexQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 10].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 11].Value = Convert.ToDecimal(Sales_Total_SalesPriceConvxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 11].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 12].Value = Convert.ToDecimal(Sales_Total_StdMatlCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 12].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 13].Value = Convert.ToDecimal(Sales_Total_StdResinCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 13].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 14].Value = Convert.ToDecimal(Sales_Total_StdPIProcess_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 14].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 15].Value = Convert.ToDecimal(Sales_Total_StdHiddenProfit_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 15].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 16].Value = Convert.ToDecimal(Sales_Total_StdSFAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 16].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 17].Value = Convert.ToDecimal(Sales_Total_StdFGAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 17].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 18].Value = Convert.ToDecimal(Sales_Total_StdUnitCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 18].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 19].Value = Convert.ToDecimal(Sales_Total_ActlMatlUnitCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 19].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 20].Value = Convert.ToDecimal(Sales_Total_ActlLandedCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 20].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 21].Value = Convert.ToDecimal(Sales_Total_ActlResinCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 21].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 22].Value = Convert.ToDecimal(Sales_Total_ActlPIProcess_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 22].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 23].Value = Convert.ToDecimal(Sales_Total_ActdHiddenProfit_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 23].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 24].Value = Convert.ToDecimal(Sales_Total_ActlSFAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 24].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 25].Value = Convert.ToDecimal(Sales_Total_ActlFGAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 25].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 26].Value = Convert.ToDecimal(Sales_Total_ActlUnitCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 26].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 27].Value = Convert.ToDecimal(Sales_Total_StdUnitCost_PHPminusActlUnitCost_PHP);
                            SalesSheets.Cells[salesSheetsRow, 27].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 28].Value = Convert.ToDecimal(Sales_Total_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                            SalesSheets.Cells[salesSheetsRow, 28].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 29].Value = Convert.ToDecimal(Sales_Total___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__);
                            SalesSheets.Cells[salesSheetsRow, 29].Style.WrapText = false;
                            salesSheetsRow++;
                            #endregion

                            #region FAMILY TOTAL
                            var FinishedGoodsGroupSales_SampleJOList_GroubyFamily = FinishedGoodsGroupSales_SampleJOList
                            .OrderBy(x => x.Family)
                            .GroupBy(x => x.FamilyDesc)
                            .ToList();
                            salesSheetsRow = salesSheetsRow + 2;
                            decimal Sales_GrandTotalFamily_QtyShipped = 0;
                            decimal Sales_GrandTotalFamily_SalesPricexQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_SalesPriceConvxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_StdMatlCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_StdResinCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_StdPIProcess_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_StdHiddenProfit_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_StdSFAdded_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_StdFGAdded_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_StdUnitCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_ActlMatlUnitCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_ActlLandedCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_ActlResinCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_ActlPIProcess_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_ActdHiddenProfit_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_ActlSFAdded_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_ActlFGAdded_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_ActlUnitCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                            decimal Sales_GrandTotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                            decimal Sales_GrandTotalFamily___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ = 0;

                            int CurrentSalesDataCount = salesSheetsRow + FinishedGoodsGroupSales_SampleJOList_GroubyFamily.ToList().Count - 1;
                            foreach (var FinishedGoodsGroupSales_SampleJOListObj_GroubyFamily in FinishedGoodsGroupSales_SampleJOList_GroubyFamily)
                            {
                                string Family = "";
                                string FamilyDesc = "";

                                decimal Sales_TotalFamily_QtyShipped = 0;
                                decimal Sales_TotalFamily_SalesPricexQtyShipped = 0;
                                decimal Sales_TotalFamily_SalesPriceConvxQtyShipped = 0;
                                decimal Sales_TotalFamily_StdMatlCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_StdResinCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_StdPIProcess_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_StdHiddenProfit_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_StdSFAdded_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_StdFGAdded_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_StdUnitCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_ActlMatlUnitCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_ActlLandedCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_ActlResinCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_ActlPIProcess_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_ActdHiddenProfit_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_ActlSFAdded_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_ActlFGAdded_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_ActlUnitCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                                decimal Sales_TotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                                decimal Sales_TotalFamily___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ = 0;
                                foreach (var FinishedGoodsGroupSales_SampleJOObj_GroubyFamily in FinishedGoodsGroupSales_SampleJOListObj_GroubyFamily)
                                {

                                    Family = FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.Family;
                                    FamilyDesc = FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.FamilyDesc;

                                    Sales_TotalFamily_QtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_SalesPricexQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.SalesPrice * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_SalesPriceConvxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.SalesPriceConv * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_StdMatlCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdMatlCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_StdResinCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_StdPIProcess_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_StdHiddenProfit_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_StdSFAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_StdFGAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_StdUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_ActlMatlUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlMatlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_ActlLandedCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlLandedCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_ActlResinCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_ActlPIProcess_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_ActdHiddenProfit_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_ActlSFAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_ActlFGAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_ActlUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_TotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlUnitCost_PHP;
                                    Sales_TotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdUnitCost_PHP - (FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlLandedCost_PHP);
                                    Sales_TotalFamily___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ += (FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped) - (FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped);

                                    Sales_GrandTotalFamily_QtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_SalesPricexQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.SalesPrice * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_SalesPriceConvxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.SalesPriceConv * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_StdMatlCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdMatlCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_StdResinCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_StdPIProcess_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_StdHiddenProfit_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_StdSFAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_StdFGAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_StdUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_ActlMatlUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlMatlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_ActlLandedCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlLandedCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_ActlResinCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_ActlPIProcess_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_ActdHiddenProfit_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_ActlSFAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_ActlFGAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_ActlUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped;
                                    Sales_GrandTotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlUnitCost_PHP;
                                    Sales_GrandTotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdUnitCost_PHP - (FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlLandedCost_PHP);
                                    Sales_GrandTotalFamily___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ += (FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped) - (FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyFamily.QtyShipped);

                                }

                                if (salesSheetsRow < CurrentSalesDataCount)
                                {
                                    SalesSheets.InsertRow((salesSheetsRow + 1), 1);
                                    SalesSheets.Cells[salesSheetsRow, 1, salesSheetsRow, 100].Copy(SalesSheets.Cells[(salesSheetsRow + 1), 1, (salesSheetsRow + 1), 1]);
                                }

                                SalesSheets.Cells[salesSheetsRow, 7].Value = Family;
                                SalesSheets.Cells[salesSheetsRow, 7].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 8].Value = FamilyDesc;
                                SalesSheets.Cells[salesSheetsRow, 8].Style.WrapText = false;

                                SalesSheets.Cells[salesSheetsRow, 9].Value = Convert.ToDecimal(Sales_TotalFamily_QtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 9].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 10].Value = Convert.ToDecimal(Sales_TotalFamily_SalesPricexQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 10].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 11].Value = Convert.ToDecimal(Sales_TotalFamily_SalesPriceConvxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 11].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 12].Value = Convert.ToDecimal(Sales_TotalFamily_StdMatlCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 12].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 13].Value = Convert.ToDecimal(Sales_TotalFamily_StdResinCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 13].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 14].Value = Convert.ToDecimal(Sales_TotalFamily_StdPIProcess_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 14].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 15].Value = Convert.ToDecimal(Sales_TotalFamily_StdHiddenProfit_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 15].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 16].Value = Convert.ToDecimal(Sales_TotalFamily_StdSFAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 16].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 17].Value = Convert.ToDecimal(Sales_TotalFamily_StdFGAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 17].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 18].Value = Convert.ToDecimal(Sales_TotalFamily_StdUnitCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 18].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 19].Value = Convert.ToDecimal(Sales_TotalFamily_ActlMatlUnitCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 19].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 20].Value = Convert.ToDecimal(Sales_TotalFamily_ActlLandedCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 20].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 21].Value = Convert.ToDecimal(Sales_TotalFamily_ActlResinCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 21].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 22].Value = Convert.ToDecimal(Sales_TotalFamily_ActlPIProcess_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 22].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 23].Value = Convert.ToDecimal(Sales_TotalFamily_ActdHiddenProfit_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 23].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 24].Value = Convert.ToDecimal(Sales_TotalFamily_ActlSFAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 24].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 25].Value = Convert.ToDecimal(Sales_TotalFamily_ActlFGAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 25].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 26].Value = Convert.ToDecimal(Sales_TotalFamily_ActlUnitCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 26].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 27].Value = Convert.ToDecimal(Sales_TotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP);
                                SalesSheets.Cells[salesSheetsRow, 27].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 28].Value = Convert.ToDecimal(Sales_TotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                                SalesSheets.Cells[salesSheetsRow, 28].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 29].Value = Convert.ToDecimal(Sales_TotalFamily___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__);
                                SalesSheets.Cells[salesSheetsRow, 29].Style.WrapText = false;
                                salesSheetsRow++;
                            }
                            SalesSheets.Cells[salesSheetsRow, 9].Value = Convert.ToDecimal(Sales_GrandTotalFamily_QtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 9].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 10].Value = Convert.ToDecimal(Sales_GrandTotalFamily_SalesPricexQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 10].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 11].Value = Convert.ToDecimal(Sales_GrandTotalFamily_SalesPriceConvxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 11].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 12].Value = Convert.ToDecimal(Sales_GrandTotalFamily_StdMatlCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 12].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 13].Value = Convert.ToDecimal(Sales_GrandTotalFamily_StdResinCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 13].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 14].Value = Convert.ToDecimal(Sales_GrandTotalFamily_StdPIProcess_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 14].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 15].Value = Convert.ToDecimal(Sales_GrandTotalFamily_StdHiddenProfit_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 15].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 16].Value = Convert.ToDecimal(Sales_GrandTotalFamily_StdSFAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 16].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 17].Value = Convert.ToDecimal(Sales_GrandTotalFamily_StdFGAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 17].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 18].Value = Convert.ToDecimal(Sales_GrandTotalFamily_StdUnitCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 18].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 19].Value = Convert.ToDecimal(Sales_GrandTotalFamily_ActlMatlUnitCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 19].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 20].Value = Convert.ToDecimal(Sales_GrandTotalFamily_ActlLandedCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 20].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 21].Value = Convert.ToDecimal(Sales_GrandTotalFamily_ActlResinCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 21].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 22].Value = Convert.ToDecimal(Sales_GrandTotalFamily_ActlPIProcess_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 22].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 23].Value = Convert.ToDecimal(Sales_GrandTotalFamily_ActdHiddenProfit_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 23].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 24].Value = Convert.ToDecimal(Sales_GrandTotalFamily_ActlSFAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 24].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 25].Value = Convert.ToDecimal(Sales_GrandTotalFamily_ActlFGAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 25].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 26].Value = Convert.ToDecimal(Sales_GrandTotalFamily_ActlUnitCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 26].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 27].Value = Convert.ToDecimal(Sales_GrandTotalFamily_StdUnitCost_PHPminusActlUnitCost_PHP);
                            SalesSheets.Cells[salesSheetsRow, 27].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 28].Value = Convert.ToDecimal(Sales_GrandTotalFamily_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                            SalesSheets.Cells[salesSheetsRow, 28].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 29].Value = Convert.ToDecimal(Sales_GrandTotalFamily___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__);
                            SalesSheets.Cells[salesSheetsRow, 29].Style.WrapText = false;
                            #endregion

                            #region Product TOTAL
                            var FinishedGoodsGroupSales_SampleJOList_GroubyProduct = FinishedGoodsGroupSales_SampleJOList
                            .OrderBy(x => x.ProductCode.Replace("FG-", "").Replace("SA-", "").Replace("RM-", ""))
                            .GroupBy(x => x.ProductCode.Replace("FG-","").Replace("SA-", "").Replace("RM-", ""))
                            .ToList();
                            salesSheetsRow = salesSheetsRow + 3;
                            decimal Sales_GrandTotalProduct_QtyShipped = 0;
                            decimal Sales_GrandTotalProduct_SalesPricexQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_SalesPriceConvxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_StdMatlCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_StdResinCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_StdPIProcess_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_StdHiddenProfit_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_StdSFAdded_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_StdFGAdded_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_StdUnitCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_ActlMatlUnitCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_ActlLandedCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_ActlResinCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_ActlPIProcess_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_ActdHiddenProfit_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_ActlSFAdded_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_ActlFGAdded_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_ActlUnitCost_PHPxQtyShipped = 0;
                            decimal Sales_GrandTotalProduct_StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                            decimal Sales_GrandTotalProduct_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                            decimal Sales_GrandTotalProduct___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ = 0;

                            CurrentSalesDataCount = salesSheetsRow + FinishedGoodsGroupSales_SampleJOList_GroubyProduct.ToList().Count - 1;
                            foreach (var FinishedGoodsGroupSales_SampleJOListObj_GroubyProduct in FinishedGoodsGroupSales_SampleJOList_GroubyProduct)
                            {
                                string ProductCode = "";

                                decimal Sales_TotalProduct_QtyShipped = 0;
                                decimal Sales_TotalProduct_SalesPricexQtyShipped = 0;
                                decimal Sales_TotalProduct_SalesPriceConvxQtyShipped = 0;
                                decimal Sales_TotalProduct_StdMatlCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_StdResinCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_StdPIProcess_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_StdHiddenProfit_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_StdSFAdded_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_StdFGAdded_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_StdUnitCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_ActlMatlUnitCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_ActlLandedCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_ActlResinCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_ActlPIProcess_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_ActdHiddenProfit_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_ActlSFAdded_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_ActlFGAdded_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_ActlUnitCost_PHPxQtyShipped = 0;
                                decimal Sales_TotalProduct_StdUnitCost_PHPminusActlUnitCost_PHP = 0;
                                decimal Sales_TotalProduct_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ = 0;
                                decimal Sales_TotalProduct___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ = 0;
                                foreach (var FinishedGoodsGroupSales_SampleJOObj_GroubyProduct in FinishedGoodsGroupSales_SampleJOListObj_GroubyProduct)
                                {

                                    ProductCode = FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ProductCode.Replace("FG-", "").Replace("SA-", "").Replace("RM-", "");

                                    Sales_TotalProduct_QtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_SalesPricexQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.SalesPrice * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_SalesPriceConvxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.SalesPriceConv * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_StdMatlCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdMatlCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_StdResinCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_StdPIProcess_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_StdHiddenProfit_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_StdSFAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_StdFGAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_StdUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_ActlMatlUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlMatlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_ActlLandedCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlLandedCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_ActlResinCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_ActlPIProcess_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_ActdHiddenProfit_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_ActlSFAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_ActlFGAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_ActlUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_TotalProduct_StdUnitCost_PHPminusActlUnitCost_PHP += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlUnitCost_PHP;
                                    Sales_TotalProduct_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdUnitCost_PHP - (FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlLandedCost_PHP);
                                    Sales_TotalProduct___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ += (FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped) - (FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped);

                                    Sales_GrandTotalProduct_QtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_SalesPricexQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.SalesPrice * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_SalesPriceConvxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.SalesPriceConv * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_StdMatlCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdMatlCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_StdResinCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_StdPIProcess_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_StdHiddenProfit_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_StdSFAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_StdFGAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_StdUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_ActlMatlUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlMatlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_ActlLandedCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlLandedCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_ActlResinCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlResinCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_ActlPIProcess_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlPIProcess_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_ActdHiddenProfit_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlHiddenProfit_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_ActlSFAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlSFAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_ActlFGAdded_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlFGAdded_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_ActlUnitCost_PHPxQtyShipped += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped;
                                    Sales_GrandTotalProduct_StdUnitCost_PHPminusActlUnitCost_PHP += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlUnitCost_PHP;
                                    Sales_GrandTotalProduct_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__ += FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdUnitCost_PHP - (FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlUnitCost_PHP - FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlLandedCost_PHP);
                                    Sales_GrandTotalProduct___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__ += (FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.StdUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped) - (FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.ActlUnitCost_PHP * FinishedGoodsGroupSales_SampleJOObj_GroubyProduct.QtyShipped);

                                }

                                if (salesSheetsRow < CurrentSalesDataCount)
                                {
                                    SalesSheets.InsertRow((salesSheetsRow + 1), 1);
                                    SalesSheets.Cells[salesSheetsRow, 1, salesSheetsRow, 100].Copy(SalesSheets.Cells[(salesSheetsRow + 1), 1, (salesSheetsRow + 1), 1]);
                                }

                                SalesSheets.Cells[salesSheetsRow, 8].Value = ProductCode;
                                SalesSheets.Cells[salesSheetsRow, 8].Style.WrapText = false;

                                SalesSheets.Cells[salesSheetsRow, 9].Value = Convert.ToDecimal(Sales_TotalProduct_QtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 9].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 10].Value = Convert.ToDecimal(Sales_TotalProduct_SalesPricexQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 10].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 11].Value = Convert.ToDecimal(Sales_TotalProduct_SalesPriceConvxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 11].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 12].Value = Convert.ToDecimal(Sales_TotalProduct_StdMatlCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 12].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 13].Value = Convert.ToDecimal(Sales_TotalProduct_StdResinCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 13].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 14].Value = Convert.ToDecimal(Sales_TotalProduct_StdPIProcess_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 14].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 15].Value = Convert.ToDecimal(Sales_TotalProduct_StdHiddenProfit_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 15].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 16].Value = Convert.ToDecimal(Sales_TotalProduct_StdSFAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 16].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 17].Value = Convert.ToDecimal(Sales_TotalProduct_StdFGAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 17].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 18].Value = Convert.ToDecimal(Sales_TotalProduct_StdUnitCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 18].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 19].Value = Convert.ToDecimal(Sales_TotalProduct_ActlMatlUnitCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 19].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 20].Value = Convert.ToDecimal(Sales_TotalProduct_ActlLandedCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 20].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 21].Value = Convert.ToDecimal(Sales_TotalProduct_ActlResinCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 21].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 22].Value = Convert.ToDecimal(Sales_TotalProduct_ActlPIProcess_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 22].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 23].Value = Convert.ToDecimal(Sales_TotalProduct_ActdHiddenProfit_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 23].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 24].Value = Convert.ToDecimal(Sales_TotalProduct_ActlSFAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 24].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 25].Value = Convert.ToDecimal(Sales_TotalProduct_ActlFGAdded_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 25].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 26].Value = Convert.ToDecimal(Sales_TotalProduct_ActlUnitCost_PHPxQtyShipped);
                                SalesSheets.Cells[salesSheetsRow, 26].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 27].Value = Convert.ToDecimal(Sales_TotalProduct_StdUnitCost_PHPminusActlUnitCost_PHP);
                                SalesSheets.Cells[salesSheetsRow, 27].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 28].Value = Convert.ToDecimal(Sales_TotalProduct_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                                SalesSheets.Cells[salesSheetsRow, 28].Style.WrapText = false;
                                SalesSheets.Cells[salesSheetsRow, 29].Value = Convert.ToDecimal(Sales_TotalProduct___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__);
                                SalesSheets.Cells[salesSheetsRow, 29].Style.WrapText = false;
                                salesSheetsRow++;
                            }
                            SalesSheets.Cells[salesSheetsRow, 9].Value = Convert.ToDecimal(Sales_GrandTotalProduct_QtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 9].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 10].Value = Convert.ToDecimal(Sales_GrandTotalProduct_SalesPricexQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 10].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 11].Value = Convert.ToDecimal(Sales_GrandTotalProduct_SalesPriceConvxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 11].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 12].Value = Convert.ToDecimal(Sales_GrandTotalProduct_StdMatlCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 12].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 13].Value = Convert.ToDecimal(Sales_GrandTotalProduct_StdResinCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 13].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 14].Value = Convert.ToDecimal(Sales_GrandTotalProduct_StdPIProcess_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 14].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 15].Value = Convert.ToDecimal(Sales_GrandTotalProduct_StdHiddenProfit_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 15].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 16].Value = Convert.ToDecimal(Sales_GrandTotalProduct_StdSFAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 16].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 17].Value = Convert.ToDecimal(Sales_GrandTotalProduct_StdFGAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 17].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 18].Value = Convert.ToDecimal(Sales_GrandTotalProduct_StdUnitCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 18].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 19].Value = Convert.ToDecimal(Sales_GrandTotalProduct_ActlMatlUnitCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 19].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 20].Value = Convert.ToDecimal(Sales_GrandTotalProduct_ActlLandedCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 20].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 21].Value = Convert.ToDecimal(Sales_GrandTotalProduct_ActlResinCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 21].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 22].Value = Convert.ToDecimal(Sales_GrandTotalProduct_ActlPIProcess_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 22].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 23].Value = Convert.ToDecimal(Sales_GrandTotalProduct_ActdHiddenProfit_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 23].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 24].Value = Convert.ToDecimal(Sales_GrandTotalProduct_ActlSFAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 24].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 25].Value = Convert.ToDecimal(Sales_GrandTotalProduct_ActlFGAdded_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 25].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 26].Value = Convert.ToDecimal(Sales_GrandTotalProduct_ActlUnitCost_PHPxQtyShipped);
                            SalesSheets.Cells[salesSheetsRow, 26].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 27].Value = Convert.ToDecimal(Sales_GrandTotalProduct_StdUnitCost_PHPminusActlUnitCost_PHP);
                            SalesSheets.Cells[salesSheetsRow, 27].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 28].Value = Convert.ToDecimal(Sales_GrandTotalProduct_StdUnitCost_PHPminus__ActlUnitCost_PHPminusActlLandedCost_PHP__);
                            SalesSheets.Cells[salesSheetsRow, 28].Style.WrapText = false;
                            SalesSheets.Cells[salesSheetsRow, 29].Value = Convert.ToDecimal(Sales_GrandTotalProduct___StdUnitCost_PHPxQtyShipped__minus__ActlUnitCost_PHPxQtyShipped__);
                            SalesSheets.Cells[salesSheetsRow, 29].Style.WrapText = false;
                            #endregion
                        }
                    }
                    #endregion
                    #region Sales Summary
                    List<SalesSummary> SalesSummaryList = new List<SalesSummary>();
                    using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LSPI803_App"].ConnectionString.ToString()))
                    {
                        conn.Open();
                        using (SqlCommand cmdSql = conn.CreateCommand())
                        {

                            cmdSql.CommandType = CommandType.StoredProcedure;
                            cmdSql.CommandText = "LSP_Rpt_NewDM_SalesSummaryReportSp";
                            cmdSql.Parameters.AddWithValue("@StartDate", StartDate);
                            cmdSql.Parameters.AddWithValue("@EndDate", EndDate);
                            cmdSql.CommandTimeout = 0;
                            using (SqlDataReader sdr = cmdSql.ExecuteReader())
                            {
                                while (sdr.Read())
                                {
                                    SalesSummaryList.Add(new SalesSummary
                                    {
                                        inv_date = sdr["inv_date"].ToString(),
                                        inv_num = sdr["inv_num"].ToString(),
                                        ship_to_cust = sdr["ship_to_cust"].ToString(),
                                        inv_desc = sdr["inv_desc"].ToString(),
                                        amount = sdr["amount"] == null ? 0 : Convert.ToDecimal(sdr["amount"]),
                                        amount_php = sdr["amount_php"] == null ? 0 : Convert.ToDecimal(sdr["amount_php"]),
                                        exch_rate = sdr["exch_rate"] == null ? 0 : Convert.ToDecimal(sdr["exch_rate"]),
                                        eng_design = sdr["eng_design"] == null ? 0 : Convert.ToDecimal(sdr["eng_design"]),
                                        price = sdr["price"] == null ? 0 : Convert.ToDecimal(sdr["price"]),
                                    });
                                }
                            }
                        }
                        conn.Close();
                    }
                    var SalesSummaryListOrderByinv_date = SalesSummaryList.OrderBy(x => x.inv_date).ToList();
                    ExcelWorksheet SalesSummarySheets = excelPackage.Workbook.Worksheets["SalesSummary"];
                    int salesSummarySheetsRow = 6;

                    SalesSummarySheets.Cells["A2"].Value = "From "+DateTime.Parse(StartDate).ToString("MMM dd, yyyy") +" to "+ DateTime.Parse(EndDate).ToString("MMM dd, yyyy"); 
                    SalesSummarySheets.Cells["A4"].Value = SalesSummaryList[0].exch_rate;
                    //Multiple Fonts in the same cell
                    ExcelRange rg = SalesSummarySheets.Cells["A4"];
                    rg.IsRichText = true;
                    //ExcelRichText uses "using OfficeOpenXml.Style;"
                    ExcelRichText text1 = rg.RichText.Add("Exchange Rate: ");
                    text1.Bold = true;
                    ExcelRichText text2 = rg.RichText.Add(SalesSummaryList[0].exch_rate.ToString("0.00"));
                    text2.UnderLine = true;                    

                    decimal Total_price = 0;
                    decimal Total_amount_php = 0;
                    decimal Product_price = 0;
                    decimal Product_amount_php = 0;
                    decimal GrandTotal_price = 0;
                    decimal GrandTotal_amount_php = 0;
                    decimal Sales_price = 0;
                    decimal Sales_amount_php = 0;
                    foreach (var SalesSummaryListOrderByinv_dateObj in SalesSummaryListOrderByinv_date)
                    {
                        if (salesSummarySheetsRow < SalesSummaryListOrderByinv_date.ToList().Count + 5)
                        {
                            SalesSummarySheets.InsertRow((salesSummarySheetsRow + 1), 1);
                            SalesSummarySheets.Cells[salesSummarySheetsRow, 1, salesSummarySheetsRow, 100].Copy(SalesSummarySheets.Cells[(salesSummarySheetsRow + 1), 1, (salesSummarySheetsRow + 1), 1]);
                        }

                        SalesSummarySheets.Cells[salesSummarySheetsRow, 1].Value = DateTime.Parse(SalesSummaryListOrderByinv_dateObj.inv_date).ToString("mm/dd/yyyy");
                        SalesSummarySheets.Cells[salesSummarySheetsRow, 1].Style.WrapText = false;
                        SalesSummarySheets.Cells[salesSummarySheetsRow, 2].Value = SalesSummaryListOrderByinv_dateObj.inv_num;
                        SalesSummarySheets.Cells[salesSummarySheetsRow, 2].Style.WrapText = false;
                        SalesSummarySheets.Cells[salesSummarySheetsRow, 3].Value = SalesSummaryListOrderByinv_dateObj.ship_to_cust;
                        SalesSummarySheets.Cells[salesSummarySheetsRow, 3].Style.WrapText = false;
                        SalesSummarySheets.Cells[salesSummarySheetsRow, 4].Value = SalesSummaryListOrderByinv_dateObj.inv_desc;
                        SalesSummarySheets.Cells[salesSummarySheetsRow, 4].Style.WrapText = false;

                        SalesSummarySheets.Cells[salesSummarySheetsRow, 5].Value = Convert.ToDecimal(SalesSummaryListOrderByinv_dateObj.price);
                        SalesSummarySheets.Cells[salesSummarySheetsRow, 5].Style.WrapText = false;
                        SalesSummarySheets.Cells[salesSummarySheetsRow, 6].Value = Convert.ToDecimal(SalesSummaryListOrderByinv_dateObj.amount_php);
                        SalesSummarySheets.Cells[salesSummarySheetsRow, 6].Style.WrapText = false;

                        Total_price += SalesSummaryListOrderByinv_dateObj.price - SalesSummaryListOrderByinv_dateObj.eng_design;
                        Total_amount_php += SalesSummaryListOrderByinv_dateObj.amount_php;

                        Product_price += SalesSummaryListOrderByinv_dateObj.eng_design;
                        Product_amount_php += SalesSummaryListOrderByinv_dateObj.eng_design* SalesSummaryListOrderByinv_dateObj.exch_rate;

                        GrandTotal_price += SalesSummaryListOrderByinv_dateObj.price;
                        GrandTotal_amount_php += SalesSummaryListOrderByinv_dateObj.amount_php;
                        
                        salesSummarySheetsRow++;
                    }


                    Sales_price += Total_price;
                    Sales_amount_php += GrandTotal_amount_php - (Product_price * SalesSummaryList[0].exch_rate);

                    SalesSummarySheets.Cells[salesSummarySheetsRow, 5].Value = Convert.ToDecimal(Total_price);
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 5].Style.WrapText = false;
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 6].Value = Convert.ToDecimal(Total_amount_php);
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 6].Style.WrapText = false;
                    salesSummarySheetsRow++;

                    SalesSummarySheets.Cells[salesSummarySheetsRow, 5].Value = Convert.ToDecimal(Product_price);
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 5].Style.WrapText = false;
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 6].Value = Convert.ToDecimal(Product_amount_php);
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 6].Style.WrapText = false;
                    salesSummarySheetsRow++;

                    SalesSummarySheets.Cells[salesSummarySheetsRow, 5].Value = Convert.ToDecimal(GrandTotal_price);
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 5].Style.WrapText = false;
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 6].Value = Convert.ToDecimal(GrandTotal_amount_php);
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 6].Style.WrapText = false;
                    salesSummarySheetsRow++;

                    SalesSummarySheets.Cells[salesSummarySheetsRow, 5].Value = Convert.ToDecimal(Sales_price);
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 5].Style.WrapText = false;
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 6].Value = Convert.ToDecimal(Sales_amount_php);
                    SalesSummarySheets.Cells[salesSummarySheetsRow, 6].Style.WrapText = false;
                    salesSummarySheetsRow++;

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