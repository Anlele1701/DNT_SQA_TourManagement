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
    public class SPTOURsController : Controller
    {
        private TourDLEntities db = new TourDLEntities();

        // GET: SPTOURs
        public ActionResult Index(string SearchString)
        {
            var sptour = db.SPTOURs.ToList();
            if (!string.IsNullOrEmpty(SearchString))
            {
                sptour = sptour.Where(s => s.TenSPTour.Contains(SearchString) || s.ID_SPTour.Contains(SearchString)).ToList();
            }
            return View(sptour);
        }

        // GET: SPTOURs/Details/5
        public ActionResult Details(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            SPTOUR sPTOUR = db.SPTOURs.Find(id);
            if (sPTOUR == null)
            {
                return HttpNotFound();
            }
            return View(sPTOUR);
        }

        // GET: SPTOURs/Create
        public ActionResult Create()
        {
            ViewBag.ID_NV = new SelectList(db.NHANVIENs, "ID_NV", "HoTen_NV");
            ViewBag.ID_TOUR = new SelectList(db.TOURs, "ID_TOUR", "TenTour");
            return View();
        }

        // POST: SPTOURs/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID_SPTour,TenSPTour,GiaNguoiLon,NgayKhoiHanh,NgayKetThuc,MoTa,DiemTapTrung,DiemDen,SoNguoi,HinhAnh,GiaTreEm,ID_NV,ID_TOUR")] SPTOUR sPTOUR)
        {
            if (ModelState.IsValid)
            {
                db.SPTOURs.Add(sPTOUR);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.ID_NV = new SelectList(db.NHANVIENs, "ID_NV", "HoTen_NV", sPTOUR.ID_NV);
            ViewBag.ID_TOUR = new SelectList(db.TOURs, "ID_TOUR", "TenTour", sPTOUR.ID_TOUR);
            return View(sPTOUR);
        }

        // GET: SPTOURs/Edit/5
        public ActionResult Edit(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            SPTOUR sPTOUR = db.SPTOURs.Find(id);
            if (sPTOUR == null)
            {
                return HttpNotFound();
            }
            ViewBag.ID_NV = new SelectList(db.NHANVIENs, "ID_NV", "HoTen_NV", sPTOUR.ID_NV);
            ViewBag.ID_TOUR = new SelectList(db.TOURs, "ID_TOUR", "TenTour", sPTOUR.ID_TOUR);
            return View(sPTOUR);
        }

        // POST: SPTOURs/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID_SPTour,TenSPTour,GiaNguoiLon,NgayKhoiHanh,NgayKetThuc,MoTa,DiemTapTrung,DiemDen,SoNguoi,HinhAnh,GiaTreEm,ID_NV,ID_TOUR")] SPTOUR sPTOUR)
        {
            if (ModelState.IsValid)
            {
                db.Entry(sPTOUR).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.ID_NV = new SelectList(db.NHANVIENs, "ID_NV", "HoTen_NV", sPTOUR.ID_NV);
            ViewBag.ID_TOUR = new SelectList(db.TOURs, "ID_TOUR", "TenTour", sPTOUR.ID_TOUR);
            return View(sPTOUR);
        }

        // GET: SPTOURs/Delete/5
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            SPTOUR sPTOUR = db.SPTOURs.Find(id);
            if (sPTOUR == null)
            {
                return HttpNotFound();
            }
            return View(sPTOUR);
        }

        // POST: SPTOURs/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            SPTOUR sPTOUR = db.SPTOURs.Find(id);
            db.SPTOURs.Remove(sPTOUR);
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