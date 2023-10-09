using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(DAPM_TOURDL.Startup))]
namespace DAPM_TOURDL
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
