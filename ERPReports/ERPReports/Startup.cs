using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ERPReports.Startup))]
namespace ERPReports
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
