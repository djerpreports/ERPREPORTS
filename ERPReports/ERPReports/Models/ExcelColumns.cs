using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;
using System.IO;
namespace ERPReports.Models
{
    public class ExcelColumns
    {
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }
        public string E { get; set; }
        public string F { get; set; }
        public string G { get; set; }
        public string H { get; set; }
        public string I { get; set; }
        public string J { get; set; }
        public string K { get; set; }
        public string L { get; set; }
        public string M { get; set; }
        public string N { get; set; }
        public string O { get; set; }
        public string P { get; set; }
        public string Q { get; set; }
        public string R { get; set; }
        public string S { get; set; }
        public string T { get; set; }
        public string U { get; set; }
        public string V { get; set; }
        public string W { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
    }
    public class ExcelData
    {
        public static List<ExcelColumns> Data;

        public static List<ExcelColumns> SheetData2;

        public static int Column;

        public static string FileName;

        public static byte[] Download(string Name, string SheetName1 = "", string SheetName2 = "")
        {
            try
            {
                string templateFilename = "PrintExcel.xlsx";
                string dir = Path.GetTempPath();
                string datetimeToday = DateTime.Now.ToString("yyMMddhhmmss");
                ExcelData.FileName = string.Format(Name + datetimeToday + ".xlsx");
                FileInfo newFile = new FileInfo(Path.Combine(dir, ExcelData.FileName));

                string apptemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Template", templateFilename);
                FileInfo templateFile = new FileInfo(apptemplatePath);
                using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
                {
                    ExcelWorksheet Sheet1 = package.Workbook.Worksheets["Sheet1"];
                    if (SheetName1 != "")
                    {
                        Sheet1.Name = SheetName1;
                    }
                    else
                    {
                        Sheet1.Name = Name;
                    }

                    for (int i = 1; i <= ExcelData.Column; i++)
                    {
                        string column = GetExcelColumnName(i);

                        var data = ExcelData.Data.Select(x => x.GetType().GetProperty(column).GetValue(x)).ToList();
                        int row = 1;
                        for (int x = 0; x <= data.Count - 1; x++)
                        {
                            string value = data[x].ToString();
                            int isInt = 0;

                            double isDouble = 0.00;
                            if (int.TryParse(value, out isInt))
                            {
                                Sheet1.Cells[column + row.ToString()].Value = Convert.ToInt32(value);
                            }
                            else if (Double.TryParse(value, out isDouble))
                            {
                                Sheet1.Cells[column + row.ToString()].Value = Convert.ToDouble(value);
                            }
                            else
                            {
                                Sheet1.Cells[column + row.ToString()].Value = value;
                            }
                            row++;
                        }
                    }

                    Sheet1.Cells["A1:" + GetExcelColumnName(ExcelData.Column) + "1"].Style.Font.Bold = true;
                    Sheet1.Cells["A:AZ"].AutoFitColumns();

                    //Sheet2 
                    if (SheetName2 != "")
                    {
                        ExcelWorksheet Sheet2 = package.Workbook.Worksheets.Add(SheetName2);
                        for (int i = 1; i <= ExcelData.Column; i++)
                        {
                            string column = GetExcelColumnName(i);

                            var data = ExcelData.SheetData2.Select(x => x.GetType().GetProperty(column).GetValue(x)).ToList();
                            int row = 1;
                            for (int x = 0; x <= SheetData2.Count - 1; x++)
                            {
                                string value = data[x].ToString();
                                int isInt = 0;

                                double isDouble = 0.00;
                                if (int.TryParse(value, out isInt))
                                {
                                    Sheet2.Cells[column + row.ToString()].Value = Convert.ToInt32(value);
                                }
                                else if (Double.TryParse(value, out isDouble))
                                {
                                    Sheet2.Cells[column + row.ToString()].Value = Convert.ToDouble(value);
                                }
                                else
                                {
                                    Sheet2.Cells[column + row.ToString()].Value = value;
                                }
                                row++;
                            }
                        }

                        Sheet2.Cells["A1:" + GetExcelColumnName(ExcelData.Column) + "1"].Style.Font.Bold = true;
                        Sheet2.Cells["A:AZ"].AutoFitColumns();
                    }


                    return package.GetAsByteArray();
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

        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }
    }
}