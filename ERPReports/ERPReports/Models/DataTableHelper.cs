namespace ERPReports.Models
{
    public class DataTableHelper
    {
        public object GetPropertyValue(object obj, string name)
        {
            return obj == null ? null : obj.GetType()
            .GetProperty(name)
            .GetValue(obj, null);
        }
    }
}
