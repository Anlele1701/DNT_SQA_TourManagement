using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using DAPM_TOURDL.Models;

namespace DAPM_TOURDL.Controllers
{
    public class HOADONsController : Controller
    {
        private TourDLEntities db = new TourDLEntities();

        public ActionResult ExportToExcel()
        {
            var hOADONs = db.HOADONs.Include(h => h.KHACHHANG).Include(h => h.SPTOUR);
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("HOADON");
                var currentrow = 1;
                worksheet.Cell(currentrow, 1).Value = "ID Hóa đơn";
                worksheet.Cell(currentrow, 2).Value = "ID Khách hàng";
                worksheet.Cell(currentrow, 3).Value = "Tên khách hàng";
                worksheet.Cell(currentrow, 4).Value = "Tình trạng";
                worksheet.Cell(currentrow, 5).Value = "Ngày đặt";
                foreach (var hoadon in hOADONs)
                {
                    currentrow++;
                    worksheet.Cell(currentrow, 1).Value = hoadon.ID_HoaDon;
                    worksheet.Cell(currentrow, 2).Value = hoadon.ID_KH;
                    worksheet.Cell(currentrow, 3).Value = hoadon.KHACHHANG.HoTen_KH;
                    worksheet.Cell(currentrow, 4).Value = hoadon.TinhTrang;
                    worksheet.Cell(currentrow, 5).Value = hoadon.NgayDat;
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "DanhSachHoaDon.xlsx"
                        );
                }
            }
        }

        // GET: HOADONs
        public ActionResult Index()
        {
            var hOADONs = db.HOADONs.Include(h => h.KHACHHANG).Include(h => h.SPTOUR);
            return View(hOADONs.ToList());
        }

        // GET: HOADONs/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            HOADON hOADON = db.HOADONs.Find(id);
            if (hOADON == null)
            {
                return HttpNotFound();
            }
            return View(hOADON);
        }

        // GET: HOADONs/Create
        public ActionResult Create()
        {
            ViewBag.ID_KH = new SelectList(db.KHACHHANGs, "ID_KH", "HoTen_KH");
            ViewBag.ID_SPTour = new SelectList(db.SPTOURs, "ID_SPTour", "TenSPTour");
            return View();
        }

        // POST: HOADONs/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID_HoaDon,SLTreEm,TongTienTour,NgayDat,TinhTrang,SLNguoiLon,TienKhuyenMai,TienPhaiTra,ID_SPTour,ID_KH")] HOADON hOADON)
        {
            if (ModelState.IsValid)
            {
                db.HOADONs.Add(hOADON);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.ID_KH = new SelectList(db.KHACHHANGs, "ID_KH", "HoTen_KH", hOADON.ID_KH);
            ViewBag.ID_SPTour = new SelectList(db.SPTOURs, "ID_SPTour", "TenSPTour", hOADON.ID_SPTour);
            return View(hOADON);
        }

        // GET: HOADONs/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            HOADON hOADON = db.HOADONs.Find(id);
            if (hOADON == null)
            {
                return HttpNotFound();
            }
            ViewBag.ID_KH = new SelectList(db.KHACHHANGs, "ID_KH", "HoTen_KH", hOADON.ID_KH);
            ViewBag.ID_SPTour = new SelectList(db.SPTOURs, "ID_SPTour", "TenSPTour", hOADON.ID_SPTour);
            return View(hOADON);
        }

        // POST: HOADONs/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID_HoaDon,SLTreEm,TongTienTour,NgayDat,TinhTrang,SLNguoiLon,TienKhuyenMai,TienPhaiTra,ID_SPTour,ID_KH")] HOADON hOADON)
        {
            if (ModelState.IsValid)
            {
                db.Entry(hOADON).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.ID_KH = new SelectList(db.KHACHHANGs, "ID_KH", "HoTen_KH", hOADON.ID_KH);
            ViewBag.ID_SPTour = new SelectList(db.SPTOURs, "ID_SPTour", "TenSPTour", hOADON.ID_SPTour);
            return View(hOADON);
        }

        // GET: HOADONs/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            HOADON hOADON = db.HOADONs.Find(id);
            if (hOADON == null)
            {
                return HttpNotFound();
            }
            return View(hOADON);
        }

        // POST: HOADONs/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            HOADON hOADON = db.HOADONs.Find(id);
            db.HOADONs.Remove(hOADON);
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