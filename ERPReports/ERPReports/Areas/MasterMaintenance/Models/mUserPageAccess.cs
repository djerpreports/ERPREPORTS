
using System.ComponentModel.DataAnnotations;

namespace ERPReports.Areas.MasterMaintenance.Models
{
    public class mUserPageAccess
    {
        public int ID { get; set; }
        [Required(ErrorMessage = "UserID value is required")]
        public int UserID { get; set; }

        [Required(ErrorMessage = "PageID value is required")]
        public int PageID { get; set; }
        public bool Status { get; set; }
        public int ReadAndWrite { get; set; }
        public int Delete { get; set; }
    }
}
