using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using DAPM_TOURDL.Models;
using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace DAPM_TOURDL.Controllers
{
    public class KHACHHANGsController : Controller
    {
        private TourDLEntities db = new TourDLEntities();

        public ActionResult ExportToExcel()
        {
            var khS = db.KHACHHANGs;
            //var khS = db.HOADONs.Include(h => h.KHACHHANG).Include(h => h.SPTOUR);
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("HOADON");
                var currentrow = 1;
                worksheet.Cell(currentrow, 1).Value = "ID Khách hàng";
                worksheet.Cell(currentrow, 2).Value = "Tên khách hàng";
                worksheet.Cell(currentrow, 3).Value = "Giới tính";
                worksheet.Cell(currentrow, 4).Value = "SĐT";
                worksheet.Cell(currentrow, 5).Value = "Email";
                worksheet.Cell(currentrow, 6).Value = "Điểm";
                foreach (var hoadon in khS)
                {
                    currentrow++;
                    worksheet.Cell(currentrow, 1).Value = hoadon.ID_KH;
                    worksheet.Cell(currentrow, 2).Value = hoadon.HoTen_KH;
                    worksheet.Cell(currentrow, 3).Value = hoadon.GioiTinh_KH;
                    worksheet.Cell(currentrow, 4).Value = hoadon.SDT_KH;
                    worksheet.Cell(currentrow, 5).Value = hoadon.Mail_KH;
                    worksheet.Cell(currentrow, 6).Value = hoadon.Diem;
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "DanhSachKhachHang.xlsx"
                        );
                }
            }
        }

        // GET: KHACHHANGs
        public ActionResult Index(string SearchString)
        {
            var kh = db.KHACHHANGs.ToList();
            if (!string.IsNullOrEmpty(SearchString))
            {
                kh = kh.Where(s => s.HoTen_KH.Contains(SearchString) || s.Mail_KH.Contains(SearchString)).ToList();
            }
            return View(kh);
        }

        // GET: KHACHHANGs/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            KHACHHANG kHACHHANG = db.KHACHHANGs.Find(id);
            if (kHACHHANG == null)
            {
                return HttpNotFound();
            }
            return View(kHACHHANG);
        }

        // GET: KHACHHANGs/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: KHACHHANGs/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID_KH,HoTen_KH,GioiTinh_KH,NgaySinh_KH,MatKhau,CCCD,SDT_KH,Mail_KH,Diem")] KHACHHANG kHACHHANG)
        {
            kHACHHANG.Diem = 0;
            if (db.KHACHHANGs.Any(x => x.Mail_KH == kHACHHANG.Mail_KH) || db.NHANVIENs.Any(x => x.Mail_NV == kHACHHANG.Mail_KH))
            {
                ModelState.AddModelError("Mail_KH", "Email này đã tồn tại");
                return View(kHACHHANG);
            }
            if (db.KHACHHANGs.Any(x => x.CCCD == kHACHHANG.CCCD))
            {
                ModelState.AddModelError("CCCD", "CCCD đã tồn tại");
            }
            if (kHACHHANG.CCCD.Length != 12 || !Regex.IsMatch(kHACHHANG.CCCD, @"^[0-9]+$"))
            {
                ModelState.AddModelError("CCCD", "CCCD không đúng 12 số");
            }
            if (db.KHACHHANGs.Any(x=>x.SDT_KH == kHACHHANG.SDT_KH) || db.NHANVIENs.Any(x=>x.SDT_NV == kHACHHANG.SDT_KH))
            {
                ModelState.AddModelError("SDT_KH", "Số điện thoại đã tồn tại");
            }
            if (kHACHHANG.SDT_KH.Length != 10 || !Regex.IsMatch(kHACHHANG.SDT_KH, @"^[0-9]+$"))
            {
                ModelState.AddModelError("SDT_KH", "Số điện thoại không đúng 10 số");
            }
            DateTime ngaySinh18 = DateTime.Now.AddYears(-16);
            if(kHACHHANG.NgaySinh_KH > ngaySinh18)
            {
                ModelState.AddModelError("NgaySinh_KH", "Ngày sinh phải lớn hơn 16 tuổi");
            }
            if (!MatKhauManh(kHACHHANG.MatKhau))
            {
                ModelState.AddModelError("MatKhau", "Mật khẩu phải có ít nhất 8 ký tự, bao gồm ít nhất 1 số, 1 chữ thường, 1 chữ hoa, 1 ký tự đặc biệt");
            }
            if (ModelState.IsValid)
            {
                db.KHACHHANGs.Add(kHACHHANG);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(kHACHHANG);
        }
        private bool MatKhauManh(string password)
        {
            return password.Length >= 8 &&
                Regex.IsMatch(password, @"^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[!@#$%^&*()_+{}\[\]:;<>,.?~\\-]).*$");
        }

        // GET: KHACHHANGs/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            KHACHHANG kHACHHANG = db.KHACHHANGs.Find(id);
            if (kHACHHANG == null)
            {
                return HttpNotFound();
            }
            return View(kHACHHANG);
        }

        // POST: KHACHHANGs/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID_KH,HoTen_KH,GioiTinh_KH,NgaySinh_KH,MatKhau,CCCD,SDT_KH,Mail_KH,Diem")] KHACHHANG kHACHHANG)
        {
            // Check Mail
            if (db.KHACHHANGs.Any(x => x.Mail_KH == kHACHHANG.Mail_KH && x.ID_KH != kHACHHANG.ID_KH) || db.NHANVIENs.Any(x => x.Mail_NV == kHACHHANG.Mail_KH))
            {
                ModelState.AddModelError("Mail_KH", "Email này đã tồn tại");
                return View(kHACHHANG);
            }

            // Check CCCD
            if (db.KHACHHANGs.Any(x => x.CCCD == kHACHHANG.CCCD && x.ID_KH != kHACHHANG.ID_KH))
            {
                ModelState.AddModelError("CCCD", "CCCD đã tồn tại");
            }
            if (kHACHHANG.CCCD.Length != 12 || !Regex.IsMatch(kHACHHANG.CCCD, @"^[0-9]+$"))
            {
                ModelState.AddModelError("CCCD", "CCCD không đúng 12 số");
            }
            //
            // Check SĐT
            if (db.KHACHHANGs.Any(x => x.SDT_KH == kHACHHANG.SDT_KH && x.ID_KH != kHACHHANG.ID_KH) || db.NHANVIENs.Any(x => x.SDT_NV == kHACHHANG.SDT_KH))
            {
                ModelState.AddModelError("SDT_KH", "Số điện thoại đã tồn tại");
            }
            if (kHACHHANG.SDT_KH.Length != 10 || !Regex.IsMatch(kHACHHANG.SDT_KH, @"^[0-9]+$"))
            {
                ModelState.AddModelError("SDT_KH", "Số điện thoại không đúng 10 số");
            }
            //
            // Check Ngày sinh
            DateTime ngaySinh18 = DateTime.Now.AddYears(-16);
            if (kHACHHANG.NgaySinh_KH > ngaySinh18)
            {
                ModelState.AddModelError("NgaySinh_KH", "Ngày sinh phải lớn hơn 16 tuổi");
            }
            //
            // Check mật khẩu
            if (!MatKhauManh(kHACHHANG.MatKhau))
            {
                ModelState.AddModelError("MatKhau", "Mật khẩu phải có ít nhất 8 ký tự, bao gồm ít nhất 1 số, 1 chữ thường, 1 chữ hoa, 1 ký tự đặc biệt");
            }
            //
            // Check GIOITINH
            if (!(kHACHHANG.GioiTinh_KH == "Nam" || kHACHHANG.GioiTinh_KH == "Nữ"))
            {
                ModelState.AddModelError("GioiTinh_KH", "giới tính chỉ Nam hoặc Nữ");
            }
            //
            if (ModelState.IsValid)
            {
                db.Entry(kHACHHANG).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(kHACHHANG);
        }

        // GET: KHACHHANGs/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            KHACHHANG kHACHHANG = db.KHACHHANGs.Find(id);
            if (kHACHHANG == null)
            {
                return HttpNotFound();
            }
            return View(kHACHHANG);
        }

        // POST: KHACHHANGs/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            KHACHHANG kHACHHANG = db.KHACHHANGs.Find(id);
            db.KHACHHANGs.Remove(kHACHHANG);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        
    }
}