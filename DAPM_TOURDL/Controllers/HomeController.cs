using DAPM_TOURDL.Models;
using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace DAPM_TOURDL.Controllers
{
    public class HomeController : Controller
    {
        private TourDLEntities db = new TourDLEntities();

        public ActionResult Index()
        {
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
            return View(db.TOURs.ToList());
        }

        public ActionResult DanhMucSanPham(string id)
        {
            var data = db.SPTOURs.Where(s => s.ID_TOUR == id);
            return View(data.ToList());
        }

        public ActionResult ChiTietTour(string id)
        {
            var data = db.SPTOURs.Where(s => s.ID_SPTour == id);
            return View(data);
        }
        [HttpGet]
        public ActionResult DangNhap()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult DangNhap(KHACHHANG khachhang)
        {
            var kiemTraDangNhap = db.KHACHHANGs.Where(x => x.Mail_KH.Equals(khachhang.Mail_KH) && x.MatKhau.Equals(khachhang.MatKhau)).FirstOrDefault();
            if (kiemTraDangNhap != null)
            {
                Session["IDUser"] = kiemTraDangNhap.ID_KH;
                Session["EmailUserSS"] = kiemTraDangNhap.Mail_KH.ToString();
                Session["UsernameSS"] = kiemTraDangNhap.HoTen_KH.ToString();
                Session["GioiTinh"] = kiemTraDangNhap.GioiTinh_KH;
                Session["SDT"] = kiemTraDangNhap.SDT_KH.ToString();
                ViewBag.idkh = kiemTraDangNhap.ID_KH;
                return RedirectToAction
                    ("HomePage", "Home", new { id = khachhang.ID_KH });
            }
            else
            {
                ViewBag.Notification = "Tài khoản và mật khẩu không đúng";
            }
            return View();
        }
        public ActionResult DangXuat()
        {
            Session.Clear();
            return RedirectToAction("HomePage", "Home");
        }
        public ActionResult Profile(int id)
        {
            var data = db.KHACHHANGs.Find(id);
            return View(data);
        }
        [HttpGet]
        public ActionResult EditProfile(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            KHACHHANG course = db.KHACHHANGs.Find(id);
            if (course == null)
            {
                return HttpNotFound();
            }
            return View(course);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult EditProfile([Bind(Include = "ID_KH,HoTen_KH,GioiTinh_KH,NgaySinh_KH,MatKhau,CCCD,SDT_KH,Mail_KH,Diem")] KHACHHANG khachhang)
        {
            if (ModelState.IsValid)
            {
                db.Entry(khachhang).State = EntityState.Modified;
                db.SaveChanges();
                Session["UsernameSS"] = khachhang.HoTen_KH.ToString();
                Session["GioiTinh"] = khachhang.GioiTinh_KH;
                return RedirectToAction("Profile","Home", new { id = khachhang.ID_KH });
            }
            return View(khachhang);
        }
    }
}