using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ERPReports.Areas.Reports.Models
{
    public class ProductivityIndex_Summary
    {
        public string WorkCenterClassification { get; set; }
        public decimal std_labor_min_1st { get; set; }
        public decimal actl_labor_min_1st { get; set; }
        public decimal std_labor_min_2nd { get; set; }
        public decimal actl_labor_min_2nd { get; set; }
        public decimal std_labor_min_all { get; set; }
        public decimal actl_labor_min_all { get; set; }
    }
}