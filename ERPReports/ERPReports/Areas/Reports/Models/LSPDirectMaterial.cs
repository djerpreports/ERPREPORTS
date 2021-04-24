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
    public class MiscellaneousTransactionReport
    {
        public string SummaryGroup { get; set; }
        public string TransType { get; set; }
        public string TransDesc { get; set; }
        public string MiscTransClass { get; set; }
        public string ReasonDesc { get; set; }
        public string JobOrLot { get; set; }
        public string TransDate { get; set; }
        public string Item { get; set; }
        public string ItemDesc { get; set; }
        public int QtyCompleted { get; set; }
        public int QtyScrapped { get; set; }
        public int Employee { get; set; }
        public string Wc { get; set; }
        public decimal MatlCost_PHP { get; set; }
        public decimal MatlLandedCost_PHP { get; set; }
        public decimal PIResin_PHP { get; set; }
        public decimal PIFGProcess_PHP { get; set; }
        public decimal PIHiddenProfit_PHP { get; set; }
        public decimal SFAddedCost_PHP { get; set; }
        public decimal FGAddedCost_PHP { get; set; }
        public decimal TotalCost_PHP { get; set; }
        public decimal TransQty { get; set; }

    }
    public class LSP_Rpt_DM_FinishedGoodsSalesReport
    {
        public string FGTransType { get; set; }
        public string TransDate { get; set; }
        public string PONum { get; set; }
        public string CustomerName { get; set; }
        public string JobOrder { get; set; }
        public string JobSuffix { get; set; }
        public string Item { get; set; }
        public string ItemDesc { get; set; }
        public string ProductCode { get; set; }
        public string Family { get; set; }
        public string FamilyDesc { get; set; }
        public decimal QtyCompleted { get; set; }
        public decimal StdMatlCost_PHP { get; set; }
        public decimal StdResinCost_PHP { get; set; }
        public decimal StdPIProcess_PHP { get; set; }
        public decimal StdHiddenProfit_PHP { get; set; }
        public decimal StdSFAdded_PHP { get; set; }
        public decimal StdFGAdded_PHP { get; set; }
        public decimal StdUnitCost_PHP { get; set; }
        public decimal ActlMatlUnitCost_PHP { get; set; }
        public decimal ActlLandedCost_PHP { get; set; }
        public decimal ActlResinCost_PHP { get; set; }
        public decimal ActlPIProcess_PHP { get; set; }
        public decimal ActlHiddenProfit_PHP { get; set; }
        public decimal ActlSFAdded_PHP { get; set; }
        public decimal ActlFGAdded_PHP { get; set; }
        public decimal ActlUnitCost_PHP { get; set; }
    }
    public class FinishedGoods_Sales_SampleJO
    {
      public string TransDate { get; set; }
      public string Item { get; set; }
	  public string ItemDesc { get; set; }
      public string ProductCode { get; set; }
	  public string Family { get; set; }
      public string FamilyDesc { get; set; }
	  public string PONum { get; set; }
	  public string LotNo { get; set; }
      public string JobOrder { get; set; }
	  public string JobSuffix { get; set; }
      public string CONum { get; set; }
	  public string COLine { get; set; }
      public string CustNum { get; set; }
	  public string ShipToCust { get; set; }
      public string CustomerName { get; set; }
	  public decimal QtyShipped { get; set; }
      public decimal SalesPrice { get; set; }
	  public decimal SalesPriceConv { get; set; }
      public decimal StdMatlCost_PHP { get; set; }
	  public decimal StdLandedCost_PHP { get; set; }
      public decimal StdResinCost_PHP { get; set; }
	  public decimal StdPIProcess_PHP { get; set; }
      public decimal StdHiddenProfit_PHP { get; set; }
	  public decimal StdSFAdded_PHP { get; set; }
      public decimal StdFGAdded_PHP { get; set; }
	  public decimal StdUnitCost_PHP { get; set; }
      public decimal ActlMatlUnitCost_PHP { get; set; }
	  public decimal ActlLandedCost_PHP { get; set; }
      public decimal ActlResinCost_PHP { get; set; }
	  public decimal ActlPIProcess_PHP { get; set; }
      public decimal ActlHiddenProfit_PHP { get; set; }
	  public decimal ActlSFAdded_PHP { get; set; }
      public decimal ActlFGAdded_PHP { get; set; }
	  public decimal ActlUnitCost_PHP { get; set; }
      public string ShipCategory { get; set; }
	  public string Recoverable { get; set; }
      public string JobRemarks { get; set; }
    }
    public class SalesSummary
    {
      public string inv_date { get; set; } 
      public string inv_num { get; set; } 
      public string ship_to_cust { get; set; } 
      public string inv_desc { get; set; } 
      public decimal amount { get; set; } 
      public decimal price { get; set; } 
      public decimal amount_php { get; set; } 
      public decimal exch_rate { get; set; } 
      public decimal eng_design { get; set; } 
    }
    public class RMBeginningBalanceReport {
        public string product_code { get; set; }
        public string item { get; set; }
        public string description { get; set; }
        public string name { get; set; }
        public string lot { get; set; }
        public string lot_create_date { get; set; }
        public string loc { get; set; }
        public string u_m { get; set; }
        public decimal qty_on_hand { get; set; }
        public decimal matl_unit_cost_php { get; set; }
        public decimal landed_cost_php { get; set; }
        public decimal resin_cost_php { get; set; }
        public decimal pi_process_cost_php { get; set; }
        public decimal pi_hidden_profit_php { get; set; }
        public decimal sf_added_value_php { get; set; }
        public decimal rm_cost_php { get; set; }
        public decimal rm_cost_usd { get; set; }

    }
    public class InventoryTurnOverReport
    {
        public string trans_date { get; set; }
        public string trans_dateMMYYYY { get; set; }
        public string trans_type { get; set; }
        public string reason_code { get; set; }
        public string reason_desc { get; set; }
        public decimal qty { get; set; }
        public decimal usage_matl { get; set; } 
        public decimal usage_landed { get; set; } 
        public string item { get; set; }
        public string item_desc { get; set; }
        public string product_code { get; set; }
        public string lot { get; set; }
        public string ref_num { get; set; }
        public int ref_line { get; set; }
        public decimal invty_matl_cost { get; set; } 
        public decimal invty_landed_cost { get; set; } 
        public decimal safety_matl_cost { get; set; } 
        public string report_group { get; set; }
        public decimal usage_M1 { get; set; }
        public decimal usage_L1 { get; set; }
        public decimal usage_M2 { get; set; }
        public decimal usage_L2 { get; set; }
        public decimal usage_M3 { get; set; }
        public decimal usage_L3 { get; set; }
        public decimal usage_M4 { get; set; }
        public decimal usage_L4 { get; set; }
        public decimal usage_M5 { get; set; }
        public decimal usage_L5 { get; set; }
        public decimal usage_M6 { get; set; }
        public decimal usage_L6 { get; set; }
        public decimal usage_M7 { get; set; }
        public decimal usage_L7 { get; set; }
        public decimal usage_M8 { get; set; }
        public decimal usage_L8 { get; set; }
        public decimal usage_M9 { get; set; }
        public decimal usage_L9 { get; set; }
        public decimal usage_M10 { get; set; }
        public decimal usage_L10 { get; set; }
        public decimal usage_M11 { get; set; }
        public decimal usage_L11 { get; set; }
        public decimal usage_M12 { get; set; }
        public decimal usage_L12 { get; set; }

        public decimal M1 { get; set; }
        public decimal L1 { get; set; }
        public decimal M2 { get; set; }
        public decimal L2 { get; set; }
        public decimal M3 { get; set; }
        public decimal L3 { get; set; }
        public decimal M4 { get; set; }
        public decimal L4 { get; set; }
        public decimal M5 { get; set; }
        public decimal L5 { get; set; }
        public decimal M6 { get; set; }
        public decimal L6 { get; set; }
        public decimal M7 { get; set; }
        public decimal L7 { get; set; }
        public decimal M8 { get; set; }
        public decimal L8 { get; set; }
        public decimal M9 { get; set; }
        public decimal L9 { get; set; }
        public decimal M10 { get; set; }
        public decimal L10 { get; set; }
        public decimal M11 { get; set; }
        public decimal L11 { get; set; }
        public decimal M12 { get; set; }
        public decimal L12 { get; set; }
        public decimal MAX_3Months { get; set; }
        public decimal L_MAX_3Months { get; set; }

    }

}