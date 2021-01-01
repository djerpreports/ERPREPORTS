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
}