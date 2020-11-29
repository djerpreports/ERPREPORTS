using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ERPReports.Areas.Reports.Controllers
{
    public class LSPDirectMaterialController : Controller
    {
        // GET: Reports/LSPDirectMaterial
        public ActionResult Index()
        {
            return View("LSPDirectMaterial");
        }
    }
}