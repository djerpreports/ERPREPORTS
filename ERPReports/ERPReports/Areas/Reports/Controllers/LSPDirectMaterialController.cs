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
                                    TransDesc = sdr["TransDesc"].ToString(),
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
                                });
                            }

                        }
                    }
                    conn.Close();
                }

                List<ExcelColumns> SFScrapDataSheet = new List<ExcelColumns>();
                List<ExcelColumns> SFScrapSummarySheet = new List<ExcelColumns>();

                foreach (var MiscellaneousTransactionItem in MiscellaneousTransaction)
                {
                    SFScrapDataSheet.Add(new ExcelColumns
                    {
                        A = MiscellaneousTransactionItem.JobOrLot,
                        B = MiscellaneousTransactionItem.TransDate,
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
                        O = MiscellaneousTransactionItem.TotalCost_PHP.ToString(),

                    });
                }

                string filePath = "";
                string Filename = "LSP_Rpt_DM_MiscellaneousTransactionReport_LSPI_" + MonthYear + ".xlsx";
                filePath = Path.Combine(Server.MapPath("~/Areas/Reports/Templates/") + "LSP_Rpt_DM_MiscellaneousTransactionReport_LSPI.xlsx");
                FileInfo file = new FileInfo(filePath);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {

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
                                TotalMatlCost_PHP += Convert.ToInt32(ItemData.MatlCost_PHP);
                                TotalMatlLandedCost_PHP += Convert.ToDecimal(ItemData.MatlLandedCost_PHP);
                                TotalPIResin_PHP += Convert.ToDecimal(ItemData.PIResin_PHP);
                                TotalPIFGProcess_PHP += Convert.ToDecimal(ItemData.PIFGProcess_PHP);
                                TotalPIHiddenProfit_PHP += Convert.ToDecimal(ItemData.PIHiddenProfit_PHP);
                                TotalFGAddedCost_PHP += Convert.ToDecimal((ItemData.SFAddedCost_PHP + ItemData.FGAddedCost_PHP));
                                TotalTotalCost_PHP += Convert.ToDecimal(ItemData.TotalCost_PHP);


                                GrandTotalMatlCost_PHP += Convert.ToInt32(ItemData.MatlCost_PHP);;
                                GrandTotalMatlLandedCost_PHP += Convert.ToDecimal(ItemData.MatlLandedCost_PHP);;
                                GrandTotalPIResin_PHP += Convert.ToDecimal(ItemData.PIResin_PHP);;
                                GrandTotalPIFGProcess_PHP += Convert.ToDecimal(ItemData.PIFGProcess_PHP);;
                                GrandTotalPIHiddenProfit_PHP += Convert.ToDecimal(ItemData.PIHiddenProfit_PHP);;
                                GrandTotalFGAddedCost_PHP += Convert.ToDecimal((ItemData.SFAddedCost_PHP + ItemData.FGAddedCost_PHP));;
                                GrandTotalTotalCost_PHP += Convert.ToDecimal(ItemData.TotalCost_PHP);;
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
                            SummarySheet.Cells[(TotalRow+1), 4].Value = Convert.ToDecimal(GrandTotalMatlCost_PHP);
                            SummarySheet.Cells[(TotalRow+1), 4].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow+1), 5].Value = Convert.ToDecimal(GrandTotalMatlLandedCost_PHP);
                            SummarySheet.Cells[(TotalRow+1), 5].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow+1), 6].Value = Convert.ToDecimal(GrandTotalPIResin_PHP);
                            SummarySheet.Cells[(TotalRow+1), 6].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow+1), 7].Value = Convert.ToDecimal(GrandTotalPIFGProcess_PHP);
                            SummarySheet.Cells[(TotalRow+1), 7].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow+1), 8].Value = Convert.ToDecimal(GrandTotalPIHiddenProfit_PHP);
                            SummarySheet.Cells[(TotalRow+1), 8].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow+1), 9].Value = Convert.ToDecimal(GrandTotalFGAddedCost_PHP);
                            SummarySheet.Cells[(TotalRow+1), 9].Style.WrapText = false;
                            SummarySheet.Cells[(TotalRow+1), 10].Value = Convert.ToDecimal(GrandTotalTotalCost_PHP);
                            SummarySheet.Cells[(TotalRow+1), 10].Style.WrapText = false;
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
                                SummarySheet.Cells[SummarySheetSheetrRow, 11].Value = Convert.ToDecimal(ItemData.TotalCost_PHP);
                                SummarySheet.Cells[SummarySheetSheetrRow, 11].Style.WrapText = false;

                                Sum_TotalCost_PHP += ItemData.TotalCost_PHP;
                                sumWCTitle = ItemData.Wc;
                                endMerge = SummarySheetSheetrRow;
                                endMergeItem = SummarySheetSheetrRow;

                                TotalWC = ItemData.Wc;
                                TotalMatlCost_PHP = Convert.ToDecimal(ItemData.MatlCost_PHP);
                                TotalMatlLandedCost_PHP = Convert.ToDecimal(ItemData.MatlLandedCost_PHP);
                                TotalPIResin_PHP = Convert.ToDecimal(ItemData.PIResin_PHP);
                                TotalPIFGProcess_PHP = Convert.ToDecimal(ItemData.PIFGProcess_PHP);
                                TotalPIHiddenProfit_PHP = Convert.ToDecimal(ItemData.PIHiddenProfit_PHP);
                                TotalFGAddedCost_PHP = Convert.ToDecimal((ItemData.SFAddedCost_PHP + ItemData.FGAddedCost_PHP));
                                TotalTotalCost_PHP = Convert.ToDecimal(ItemData.TotalCost_PHP);


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
                    #region SFScrapDataSheet(SFScrapDataSheet)
                    ExcelWorksheet MiscellaneousTransactionSheet = excelPackage.Workbook.Worksheets["SF Scrap Data"];
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
    }
}