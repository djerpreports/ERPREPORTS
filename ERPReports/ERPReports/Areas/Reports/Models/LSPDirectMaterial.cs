using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ERPReports.Areas.Reports.Models
{
    public class DMAndLaborPercentageReport
    {
        public string item { get; set; }
        public string product_code { get; set; }
        public string fam_code { get; set; }
        public string trans_date { get; set; }
        public decimal qty_completed { get; set; }
        public decimal produced_amt { get; set; }
        public decimal actl_rm_cost { get; set; }
        public decimal std_rm_cost { get; set; }
        public decimal actl_lbr_cost { get; set; }
        public decimal std_lbr_cost { get; set; }
    }
    public class SlowMonitoringAnalysisReport
    {
        public string item { get; set; }
        public string description { get; set; }
        public string product_code { get; set; }
        public string Uf_location { get; set; }
        public string matl_stat { get; set; }
        public decimal QtyOnHand { get; set; }
        public decimal TotalMatlCostPHP { get; set; }
        public decimal TotalLandedCostPHP { get; set; }
        public decimal TotalPIFGProcessCostPHP { get; set; }
        public decimal TotalPIResinCostPHP { get; set; }
        public decimal TotalPIHiddenPHP { get; set; }
        public decimal TotalSFLbrCostPHP { get; set; }
        public decimal TotalCostPHP { get; set; }
        public decimal LatestPODate { get; set; }
        public decimal LatestIssueDate { get; set; }
        public string ItemRemarks { get; set; }

    }
}