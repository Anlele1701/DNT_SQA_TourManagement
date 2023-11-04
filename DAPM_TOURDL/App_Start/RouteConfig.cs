using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace DAPM_TOURDL
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

            routes.MapRoute(
                name: "Login",
                url: "Login", // Đường dẫn đến trang đăng nhập
                defaults: new { controller = "KHACHHANGs", action = "DangNhap" }
            );
            routes.MapRoute(
                name: "LoginComplete",
                url: "", // Đường dẫn đến trang chủ sau khi đăng nhập
                defaults: new { controller = "Home", action = "HomePage", id = UrlParameter.Optional }
            );
            routes.MapRoute(
                name: "Default",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "HomePage", id = UrlParameter.Optional }
            );
        }
    }
}
