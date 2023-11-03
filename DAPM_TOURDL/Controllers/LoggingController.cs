using DAPM_TOURDL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DAPM_TOURDL.Controllers
{
    public class LoggingController : Controller
    {
        private TourDLEntities db = new TourDLEntities();

        // GET: Admin
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult LoginAdmin()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult LoginAdmin(NHANVIEN nhanvien)
        {
            var check = db.NHANVIENs.Where(s => s.Mail_NV.Equals(nhanvien.Mail_NV)
          && s.MatKhau.Equals(nhanvien.MatKhau)).FirstOrDefault();
            if (check == null)
            {
                ViewBag.Notification = "Tài khoản và mật khẩu không đúng";
                return View("Index", "HomeController");
            }
            else
            {
                db.Configuration.ValidateOnSaveEnabled = false;
                Session["IDUserAdmin"] = check.ID_NV;
                Session["HoTen"] = check.HoTen_NV;
                Session["Email"] = check.Mail_NV;
                return RedirectToAction
                    ("Index", "NHANVIENs");
            }
        }
    }
}