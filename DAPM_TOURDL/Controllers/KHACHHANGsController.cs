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
        public ActionResult Index()
        {
            return View(db.KHACHHANGs.ToList());
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
            if (ModelState.IsValid)
            {
                db.KHACHHANGs.Add(kHACHHANG);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(kHACHHANG);
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