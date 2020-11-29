using System.Web.Mvc;

namespace ERPReports.Controllers
{
    [AllowAnonymous]
    public class ErrorController : Controller
    {
        // GET: Error
        public ActionResult Error404()
        {
            return View("Error404");
        }
        public ActionResult Error403()
        {
            return View("Error403");
        }
    }
}
