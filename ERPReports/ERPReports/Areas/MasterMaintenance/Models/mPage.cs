namespace ERPReports.Areas.MasterMaintenance.Models
{
    public class mPage
    {
        public int ID { get; set; }
        public string GroupLabel { get; set; }
        public string PageName { get; set; }
        public string PageLabel { get; set; }
        public string URL { get; set; }
        public int HasSub { get; set; }
        public string ParentMenu { get; set; }
        public int ParentOrder { get; set; }
        public int Order { get; set; }
        public string Icon { get; set; }
    }
}
