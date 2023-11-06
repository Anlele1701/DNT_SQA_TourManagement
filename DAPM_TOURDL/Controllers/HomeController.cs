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
            KHACHHANG khachhang = db.KHACHHANGs.Find(id);
            if (khachhang == null)
            {
                return HttpNotFound();
            }
            return View(khachhang);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult EditProfile([Bind(Include = "ID_KH,HoTen_KH,GioiTinh_KH,NgaySinh_KH,MatKhau,CCCD,SDT_KH,Mail_KH,Diem")] KHACHHANG khachhang)
        {
            DateTime ngayTruocKhiDu16Tuoi = DateTime.Now.AddYears(-16);
            if (!(khachhang.GioiTinh_KH == "Nam" || khachhang.GioiTinh_KH == "Nữ"))
            {
                ViewBag.Notification = "Giới tính chỉ có thể là 'Nam' hoặc 'Nữ'";
            }
            else if (khachhang.NgaySinh_KH > ngayTruocKhiDu16Tuoi)
            {
                ViewBag.Notification = "Ngày sinh phải đủ 16 tuổi";
            }
            else if (khachhang.CCCD.Length != 12 || !Regex.IsMatch(khachhang.CCCD, @"^[0-9]+$"))
            {
                ViewBag.Notification = "Căn Cước Công Dân vui lòng nhập đủ 12 số và không bao gồm chữ,kí tự";
            }
            else if (khachhang.SDT_KH.Length != 10 || !Regex.IsMatch(khachhang.SDT_KH, @"^[0-9]+$"))
            {
                ViewBag.Notification = "Số điện thoại phải có 10 số và không bao gồm chữ,kí tự";
            }
            else if (db.KHACHHANGs.Any(x => x.CCCD == khachhang.CCCD && x.ID_KH != khachhang.ID_KH))
            {
                ViewBag.Notification = "Căn cước công dân này đã được đăng ký";
            }
            else if (db.KHACHHANGs.Any(x => x.SDT_KH == khachhang.SDT_KH && x.ID_KH != khachhang.ID_KH))
            {
                ViewBag.Notification = "Số điện thoại này đã có người sử dụng";
            }
            else if (khachhang.HoTen_KH.Length > 32)
            {
                ViewBag.Notification = "Tên không được quá 64 ký tự";
            }
            else if (db.KHACHHANGs.Any(x => x.MatKhau != khachhang.MatKhau && x.ID_KH == khachhang.ID_KH))
            {
                ViewBag.Notification = "Mật khẩu xác nhận không chính xác";
            }
            else if (ModelState.IsValid)
            {
                db.Entry(khachhang).State = EntityState.Modified;
                db.SaveChanges();
                Session["UsernameSS"] = khachhang.HoTen_KH.ToString();
                Session["GioiTinh"] = khachhang.GioiTinh_KH;
                return RedirectToAction("Profile", "Home", new { id = khachhang.ID_KH });
            }

            return View(khachhang);
        }
    }
}