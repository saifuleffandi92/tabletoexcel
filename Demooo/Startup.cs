using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Demooo.Startup))]
namespace Demooo
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
