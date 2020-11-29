using System.Web.Mvc;

namespace ERPReports.Areas.MasterMaintenance
{
    public class MasterMaintenanceAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "MasterMaintenance";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "MasterMaintenance_default",
                "MasterMaintenance/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}