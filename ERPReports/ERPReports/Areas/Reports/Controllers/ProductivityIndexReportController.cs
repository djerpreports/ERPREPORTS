using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using ERPReports.Models;
using Microsoft.AspNet.Identity;
using System.Data.SqlClient;
using ERPReports.Areas.MasterMaintenance.Models;
using System.Collections;
using System.Configuration;
using System.Data;
using System.IO;
namespace ERPReports.Areas.Reports.Controllers
{
    public class ProductivityIndexReportController : Controller
    {
        // GET: Reports/ProductivityIndexReport
        public ActionResult Index()
        {
            return View("ProductivityIndexReport");
        }
        public ActionResult DownloadProductivityIndexReport(string MonthYear)
        {
            List<ExcelColumns> SummarySheetData = new List<ExcelColumns>();
            List<ExcelColumns> DetailedSheetData = new List<ExcelColumns>();
            try
            {
                List<ExcelColumns> datalist = new List<ExcelColumns>();
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LSPI803_App"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {

                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "LSP_Rpt_ProductivityIndex_SummarySp";
                        cmdSql.Parameters.Clear();
                        cmdSql.Parameters.AddWithValue("@MonthYear", MonthYear);
                        cmdSql.Parameters.AddWithValue("@IsSummary", 1);
                        cmdSql.ExecuteNonQuery();
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                SummarySheetData.Add(new ExcelColumns
                                {
                                    A = sdr["WorkCenterClassification"].ToString(),
                                    B = sdr["std_labor_min_1st"].ToString(),
                                    C = sdr["actl_labor_min_1st"].ToString(),
                                    E = sdr["WorkCenterClassification"].ToString(),
                                    F = sdr["std_labor_min_2nd"].ToString(),
                                    G = sdr["actl_labor_min_2nd"].ToString(),
                                    I = sdr["WorkCenterClassification"].ToString(),
                                    J = sdr["std_labor_min_all"].ToString(),
                                    K = sdr["actl_labor_min_all"].ToString(),
                                });
                            }

                        }
                    }
                    conn.Close();
                }

                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LSPI803_App"].ConnectionString.ToString()))
                {
                    conn.Open();
                    using (SqlCommand cmdSql = conn.CreateCommand())
                    {

                        cmdSql.CommandType = CommandType.StoredProcedure;
                        cmdSql.CommandText = "LSP_Rpt_ProductivityIndex_DetailedSp";                        
                        cmdSql.Parameters.Clear();
                        cmdSql.Parameters.AddWithValue("@MonthYear", MonthYear);
                        cmdSql.Parameters.AddWithValue("@IsSummary", 0);
                        cmdSql.ExecuteNonQuery();
                        using (SqlDataReader sdr = cmdSql.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                DetailedSheetData.Add(new ExcelColumns
                                {
                                    A = sdr["wc"].ToString(),
                                    B = sdr["oper_num"].ToString(),
                                    C = sdr["item"].ToString(),
                                    E = sdr["wc_desc"].ToString(),
                                    G = sdr["qty_complete"].ToString(),
                                    I = sdr["labor_min_per_pc"].ToString(),
                                    J = sdr["std_labor_min"].ToString(),
                                    K = sdr["WorkCenterClassification"].ToString(),
                                });
                            }

                        }
                    }
                    conn.Close();
                }

                var groupedWorkCenterClassification = DetailedSheetData
                    .GroupBy(u => u.K)
                    .Select(grp => grp.ToList())
                    .ToList();

                string filePath = "";
                string Filename = "LSP_Rpt_ProductivityIndexReport_" + MonthYear + ".xlsx";
                filePath = Path.Combine(Server.MapPath("~/Areas/Reports/Templates/") + "ProductivityIndexReport.xlsx");
                FileInfo file = new FileInfo(filePath);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    #region Summary(Sheet1)
                    ExcelWorksheet SummaryWorkSheet = excelPackage.Workbook.Worksheets["Summary"];
                    int rowCountTotal = SummarySheetData.Count + 8;
                    int SummaryCounter = 0;
                    decimal std_labor_min_1stTOTAL = 0;
                    decimal actl_labor_min_1stTOTAL = 0;
                    decimal std_labor_min_2ndTOTAL = 0;
                    decimal actl_labor_min_2ndTOTAL = 0;
                    decimal std_labor_min_allTOTAL = 0;
                    decimal actl_labor_min_allTOTAL = 0;

                    DateTime MonthYearDate = DateTime.Parse(MonthYear + "-01");
                    string Month = MonthYearDate.ToString("MMMM");
                    DateTime firstDayOfMonth = new DateTime(MonthYearDate.Year, MonthYearDate.Month, 1);
                    string lastDayOfMonth = (firstDayOfMonth.AddMonths(1).AddDays(-1)).ToString("dd");
                    string GenerationDate = DateTime.Now.ToString("MM/dd/yyyy");
                    string R1 = Month.ToString() + " 1 to 15";
                    string R2 = Month.ToString() + " 15 to " + lastDayOfMonth.ToString();
                    string R3 = Month.ToString() + " 1 to " + lastDayOfMonth.ToString();
                    SummaryWorkSheet.Cells[3, 2].Value = "Run Date: " + GenerationDate;
                    SummaryWorkSheet.Cells[6, 1].Value = R1;
                    SummaryWorkSheet.Cells[6, 5].Value = R2;
                    SummaryWorkSheet.Cells[6, 9].Value = R3;
                    for (int row = 8; row < rowCountTotal; row++)
                    {
                        SummaryWorkSheet.Cells[row, 1].Value = SummarySheetData[SummaryCounter].A;
                        SummaryWorkSheet.Cells[row, 1].Style.WrapText = false;
                        SummaryWorkSheet.Cells[row, 2].Value = Convert.ToDecimal(SummarySheetData[SummaryCounter].B);
                        SummaryWorkSheet.Cells[row, 2].Style.WrapText = false;
                        SummaryWorkSheet.Cells[row, 3].Value = Convert.ToDecimal(SummarySheetData[SummaryCounter].C);
                        SummaryWorkSheet.Cells[row, 3].Style.WrapText = false;
                        SummaryWorkSheet.Cells[row, 5].Value = SummarySheetData[SummaryCounter].E;
                        SummaryWorkSheet.Cells[row, 5].Style.WrapText = false;
                        SummaryWorkSheet.Cells[row, 6].Value = Convert.ToDecimal(SummarySheetData[SummaryCounter].F);
                        SummaryWorkSheet.Cells[row, 6].Style.WrapText = false;
                        SummaryWorkSheet.Cells[row, 7].Value = Convert.ToDecimal(SummarySheetData[SummaryCounter].G);
                        SummaryWorkSheet.Cells[row, 7].Style.WrapText = false;
                        SummaryWorkSheet.Cells[row, 9].Value = SummarySheetData[SummaryCounter].I;
                        SummaryWorkSheet.Cells[row, 9].Style.WrapText = false;
                        SummaryWorkSheet.Cells[row, 10].Value = Convert.ToDecimal(SummarySheetData[SummaryCounter].J);
                        SummaryWorkSheet.Cells[row, 10].Style.WrapText = false;
                        SummaryWorkSheet.Cells[row, 11].Value = Convert.ToDecimal(SummarySheetData[SummaryCounter].K);
                        SummaryWorkSheet.Cells[row, 11].Style.WrapText = false;
                        SummaryWorkSheet.InsertRow((row + 1), 1);
                        SummaryWorkSheet.Cells[row, 1, row, 100].Copy(SummaryWorkSheet.Cells[(row + 1), 1, (row + 1), 1]);

                        std_labor_min_1stTOTAL = std_labor_min_1stTOTAL + Convert.ToDecimal(SummarySheetData[SummaryCounter].B);
                        actl_labor_min_1stTOTAL = actl_labor_min_1stTOTAL + Convert.ToDecimal(SummarySheetData[SummaryCounter].C);
                        std_labor_min_2ndTOTAL = std_labor_min_2ndTOTAL + Convert.ToDecimal(SummarySheetData[SummaryCounter].F);
                        actl_labor_min_2ndTOTAL = actl_labor_min_2ndTOTAL + Convert.ToDecimal(SummarySheetData[SummaryCounter].G);
                        std_labor_min_allTOTAL = std_labor_min_allTOTAL + Convert.ToDecimal(SummarySheetData[SummaryCounter].J);
                        actl_labor_min_allTOTAL = actl_labor_min_allTOTAL + Convert.ToDecimal(SummarySheetData[SummaryCounter].K);

                        SummaryCounter++;
                    }
                    SummaryWorkSheet.Cells[rowCountTotal + 1, 2].Value = std_labor_min_1stTOTAL;
                    SummaryWorkSheet.Cells[rowCountTotal + 1, 3].Value = actl_labor_min_1stTOTAL;
                    SummaryWorkSheet.Cells[rowCountTotal + 1, 6].Value = std_labor_min_2ndTOTAL;
                    SummaryWorkSheet.Cells[rowCountTotal + 1, 7].Value = actl_labor_min_2ndTOTAL;
                    SummaryWorkSheet.Cells[rowCountTotal + 1, 10].Value = std_labor_min_allTOTAL;
                    SummaryWorkSheet.Cells[rowCountTotal + 1, 11].Value = actl_labor_min_allTOTAL;
                    #endregion

                    #region Detailed(Sheet2)
                    ExcelWorksheet DetailedWorkSheet = excelPackage.Workbook.Worksheets["Detailed"];
                    DetailedWorkSheet.Cells[3, 4].Value = "Run Date: " + GenerationDate;
                    int rowCountTotalDetails = DetailedSheetData.Count + 7;
                    int sheetrRow = 7;
                    int groupedWorkCenterClassificationCounter = 0;
                    foreach (var DetailedSheetList in groupedWorkCenterClassification)
                    {
                        //foreach (ExcelColumns DetailedList in DetailedSheetList)
                        //{
                        decimal QTY_Total = 0;
                        decimal Labor_Total = 0;
                        decimal TotalProduces_Total = 0;
                        int ClassificationHeaderRow = sheetrRow - 2;
                        for (int DetailsCounter = 0; DetailsCounter < DetailedSheetList.Count; DetailsCounter++)
                        {
                            DetailedWorkSheet.Cells[(ClassificationHeaderRow), 1].Value = DetailedSheetList[DetailsCounter].K;

                            DetailedWorkSheet.Cells[sheetrRow, 1].Value = DetailedSheetList[DetailsCounter].A;
                            DetailedWorkSheet.Cells[sheetrRow, 1].Style.WrapText = false;
                            DetailedWorkSheet.Cells[sheetrRow, 2].Value = DetailedSheetList[DetailsCounter].B;
                            DetailedWorkSheet.Cells[sheetrRow, 2].Style.WrapText = false;
                            DetailedWorkSheet.Cells[sheetrRow, 3].Value = DetailedSheetList[DetailsCounter].C;
                            DetailedWorkSheet.Cells[sheetrRow, 3].Style.WrapText = false;
                            DetailedWorkSheet.Cells[sheetrRow, 5].Value = DetailedSheetList[DetailsCounter].E;
                            DetailedWorkSheet.Cells[sheetrRow, 5].Style.WrapText = false;
                            DetailedWorkSheet.Cells[sheetrRow, 7].Value = Convert.ToDecimal(DetailedSheetList[DetailsCounter].G);
                            DetailedWorkSheet.Cells[sheetrRow, 7].Style.WrapText = false;
                            DetailedWorkSheet.Cells[sheetrRow, 9].Value = Convert.ToDecimal(DetailedSheetList[DetailsCounter].I);
                            DetailedWorkSheet.Cells[sheetrRow, 9].Style.WrapText = false;
                            DetailedWorkSheet.Cells[sheetrRow, 10].Value = Convert.ToDecimal(DetailedSheetList[DetailsCounter].J);
                            DetailedWorkSheet.Cells[sheetrRow, 10].Style.WrapText = false;
                            DetailedWorkSheet.Cells[sheetrRow, 11].Value = DetailedSheetList[DetailsCounter].K;
                            DetailedWorkSheet.Cells[sheetrRow, 11].Style.WrapText = false;

                            QTY_Total = QTY_Total + Convert.ToDecimal(DetailedSheetList[DetailsCounter].G);
                            Labor_Total = Labor_Total + Convert.ToDecimal(DetailedSheetList[DetailsCounter].I);
                            TotalProduces_Total = TotalProduces_Total + Convert.ToDecimal(DetailedSheetList[DetailsCounter].J);

                            if (DetailsCounter < DetailedSheetList.Count-1) {
                                DetailedWorkSheet.InsertRow((sheetrRow + 1), 1);
                                DetailedWorkSheet.Cells[sheetrRow, 1, sheetrRow, 100].Copy(DetailedWorkSheet.Cells[(sheetrRow + 1), 1, (sheetrRow + 1), 1]);
                                sheetrRow++;
                            }
                        }
                        int TotalRow = sheetrRow + 1;
                        DetailedWorkSheet.Cells[TotalRow, 7].Value = QTY_Total;
                        DetailedWorkSheet.Cells[TotalRow, 9].Value = Labor_Total;
                        DetailedWorkSheet.Cells[TotalRow, 10].Value = TotalProduces_Total;
                        

                        if (groupedWorkCenterClassificationCounter < groupedWorkCenterClassification.Count-1)
                        {
                            
                            sheetrRow++;
                            DetailedWorkSheet.InsertRow((sheetrRow + 1), 1);//Space
                            sheetrRow++;
                            //Classification Header
                            DetailedWorkSheet.InsertRow((sheetrRow + 1), 1);
                            DetailedWorkSheet.Cells[5, 1, 5, 100].Copy(DetailedWorkSheet.Cells[(sheetrRow + 1), 1, (sheetrRow + 1), 1]);
                            sheetrRow++;
                            //End Classification Header

                            //Details Header
                            DetailedWorkSheet.InsertRow((sheetrRow + 1), 1);
                            DetailedWorkSheet.Cells[6, 1, 6, 100].Copy(DetailedWorkSheet.Cells[(sheetrRow + 1), 1, (sheetrRow + 1), 1]);
                            sheetrRow++;
                            //End Details Header

                            //Details Content
                            DetailedWorkSheet.InsertRow((sheetrRow + 1), 1);
                            DetailedWorkSheet.Cells[7, 1, 7, 100].Copy(DetailedWorkSheet.Cells[(sheetrRow + 1), 1, (sheetrRow + 1), 1]);
                            sheetrRow++;
                            //End Details Content

                            //Details Total
                            DetailedWorkSheet.InsertRow((sheetrRow + 1), 1);
                            DetailedWorkSheet.Cells[TotalRow, 1, TotalRow, 100].Copy(DetailedWorkSheet.Cells[(sheetrRow + 1), 1, (sheetrRow + 1), 1]);
                            
                            //End Details Total
                        }
                        groupedWorkCenterClassificationCounter++;
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