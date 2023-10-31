using DAPM_TOURDL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DAPM_TOURDL.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var db = new TourDLEntities();
            return View(db.TOURs.ToList());
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult HomePage()
        {
            var db = new TourDLEntities();
            return View(db.TOURs.ToList());
        }
    }
}