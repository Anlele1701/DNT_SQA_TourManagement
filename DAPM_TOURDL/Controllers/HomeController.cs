using DAPM_TOURDL.Models;
using DAPM_TOURDL.Models.Payment;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.Owin.Security.OAuth;
using PagedList;
using System;
using System.Configuration;
using System.Data.Entity;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Web.Mvc;

namespace DAPM_TOURDL.Controllers
{
    public class HomeController : Controller
    {
        private TourDLEntities db = new TourDLEntities();
        //https://localhost:44385/

        public ActionResult Index()
        {
            return View(db.TOURs.ToList());
        }

        public ActionResult About()
        {
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
            var data = db.TOURs.ToList();
            return View(data);
        }

        public ActionResult ChiTietTour(string id)
        {
            var data = db.SPTOURs.Where(s => s.ID_SPTour == id);
            return View(data);
        }
        public ActionResult DangKy()
        {

            return View();
        }
        [HttpPost]
        public ActionResult DangKy(KHACHHANG khachhang)
        {
            DateTime ngayHienTai = DateTime.Now;
            DateTime ngaySinh18 = ngayHienTai.AddYears(-16);
            if (db.KHACHHANGs.Any(x => x.Mail_KH == khachhang.Mail_KH))
            {
                ViewBag.Notification = "Tài khoản đã tồn tại";
            }
            else if (string.IsNullOrEmpty(khachhang.Mail_KH) || khachhang.GioiTinh_KH == null || khachhang.NgaySinh_KH == null || string.IsNullOrEmpty(khachhang.MatKhau) || string.IsNullOrEmpty(khachhang.CCCD) || string.IsNullOrEmpty(khachhang.SDT_KH) || string.IsNullOrEmpty(khachhang.HoTen_KH))
            {
                ViewBag.Notification = "Vui lòng không bỏ trống field nào nhé ^_^";
            }
            else if (!(khachhang.GioiTinh_KH == "Nam" || khachhang.GioiTinh_KH == "Nữ"))
            {
                ViewBag.Notification = "Giới tính chỉ có thể là 'Nam' hoặc 'Nữ'";
            }
            else if (khachhang.NgaySinh_KH > ngaySinh18)
            {
                ViewBag.Notification = "Yêu cầu lớn hơn 16+";
            }
            else if (khachhang.CCCD.Length != 12 || !Regex.IsMatch(khachhang.CCCD, @"^[0-9]+$"))
            {
                ViewBag.Notification = "Căn Cước Công Dân vui lòng nhập đủ 12 số và không bao gồm chữ,kí tự";
            }
            else if (khachhang.SDT_KH.Length != 10 || !Regex.IsMatch(khachhang.SDT_KH, @"^[0-9]+$"))
            {
                ViewBag.Notification = "Số điện thoại phải có 10 số và không bao gồm chữ,kí tự";
            }
            else if(db.KHACHHANGs.Any(x=>x.CCCD == khachhang.CCCD))
            {
                ViewBag.Notification = "Căn Cước Công Dân này đã tồn tại";
            }
            else if(db.KHACHHANGs.Any(x=>x.SDT_KH == khachhang.SDT_KH))
            {
                ViewBag.Notification = "Số Điện Thoại này đã tồn tại";
            }
            else if (!MatKhauManh(khachhang.MatKhau))
            {
                ViewBag.Notification = "Mật khẩu phải có ít nhất 8 ký tự, bao gồm ít nhất 1 số, 1 chữ thường, 1 chữ hoa, 1 ký tự đặc biệt";
            }
            else
            {
                db.KHACHHANGs.Add(khachhang);
                db.SaveChanges();
                return RedirectToAction("DangNhap", "Home");
            }
            return View();
        }
        private bool MatKhauManh(string password)
        {
            return password.Length >= 8 &&
                Regex.IsMatch(password, @"^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[!@#$%^&*()_+{}\[\]:;<>,.?~\\-]).*$");
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
                if(kiemTraDangNhap.HoTen_KH != null)
                {
                    Session["UsernameSS"] = kiemTraDangNhap.HoTen_KH.ToString();
                }
                Session["IDUser"] = kiemTraDangNhap.ID_KH;
                /*Session["EmailUserSS"] = kiemTraDangNhap.Mail_KH.ToString();
                Session["GioiTinh"] = kiemTraDangNhap.GioiTinh_KH;
                Session["SDT"] = kiemTraDangNhap.SDT_KH.ToString();
                Session["CCCD"] = kiemTraDangNhap.CCCD.ToString();
                ViewBag.idkh = kiemTraDangNhap.ID_KH;*/
                // Kiểm tra xem có thông tin tour trong Session không
                if (Session["TourInfo"] != null)
                {
                    var tour = Session["TourInfo"];
                    Session.Remove("TourInfo");

                    return RedirectToAction("DatTour", new { id = tour });
                }

                return RedirectToAction("HomePage", "Home", new { id = kiemTraDangNhap.ID_KH });
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
        public ActionResult ThongTinCaNhan(int id)
        {
            var data = db.KHACHHANGs.Find(id);
            if (TempData["SuccessMessage"] != null)
            {
                ViewBag.SuccessMessage = TempData["SuccessMessage"];
            }
            return View(data);
        }
        [HttpGet]
        public ActionResult ChinhSuaThongTinCaNhan(int? id)
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
        public ActionResult ChinhSuaThongTinCaNhan([Bind(Include = "ID_KH,HoTen_KH,GioiTinh_KH,NgaySinh_KH,MatKhau,CCCD,SDT_KH,Mail_KH,Diem")] KHACHHANG khachhang)
        {
            DateTime ngayTruocKhiDu16Tuoi = DateTime.Now.AddYears(-16);
            if ((khachhang.GioiTinh_KH != "Nam" && khachhang.GioiTinh_KH != "Nữ"))
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
                TempData["SuccessMessage"] = "Chỉnh sửa thông tin cá nhân thành công";
                //Session["UsernameSS"] = khachhang.HoTen_KH.ToString();
                return RedirectToAction("ThongTinCaNhan", "Home", new { id = khachhang.ID_KH });
            }
            return View(khachhang);
        }
        [HttpGet]
        public ActionResult DoiMatKhau(int id)
        {
            var data = db.KHACHHANGs.Find(id);
            return View(data);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult DoiMatKhau(int id,string currentPassword, string newPassword, string confirmPassword)
        {
            var data = db.KHACHHANGs.Find(id);
            if(currentPassword != data.MatKhau)
            {
                ModelState.AddModelError("currentPassword", "Mật khẩu hiện tại không chính xác");
            }
            if(currentPassword == newPassword)
            {
                ModelState.AddModelError("newPassword", "Vui lòng đổi mật khẩu mới không, đổi lại mật khẩu cũ");
            }
            if(newPassword != confirmPassword)
            {
                ModelState.AddModelError("confirmPassword", "Mật khẩu xác nhận không trùng khớp");
            }
            if (!MatKhauManh(newPassword))
            {
                ModelState.AddModelError("newPassword", "Mật khẩu phải có ít nhất 8 ký tự, bao gồm ít nhất 1 số, 1 chữ thường, 1 chữ hoa, 1 ký tự đặc biệt");
            }
            if (!ModelState.IsValid)
            {
                TempData["currentPassword"] = currentPassword;
                TempData["newPassword"] = newPassword;
                TempData["confirmPassword"] = confirmPassword;
                return View(data);
            }
            else
            {
                data.MatKhau = newPassword;
                db.SaveChanges();
                TempData["SuccessMessage"] = "Đổi mật khẩu thành công";
                return RedirectToAction("ThongTinCaNhan", "Home", new { id = data.ID_KH });
            } 
        }
        public ActionResult LichSuDatTour(int id)
        {
            var data = db.HOADONs.Where(t => t.ID_KH == id).ToList();
            if(data.Count == 0)
            {
                ViewBag.Mes = "Bạn chưa đặt tour nào cả";
            }
            return View(data);
        }
        public ActionResult HuyTourDaDat(int? id)
        {
            if (Request.IsAjaxRequest())
            {
                Session["idhoadon"] = id;
                return PartialView();
            }
            else
            {
                return PartialView("Error");
            }
        }
        public ActionResult HuyTour(int id)
        {
            HOADON data = db.HOADONs.Find(id);
            db.HOADONs.Remove(data);
            db.SaveChanges();
            TempData["SuccessMessage"] = "Hủy tour thành công";
            return RedirectToAction("LichSuDatTour", new { id = Session["IDUser"] });
        }
        
        [HttpGet]
        public ActionResult DatTour(string id)
        {
            var data = db.SPTOURs.Find(id);
            if (data == null)
            {
                return HttpNotFound();
            }
            var idSPTOUR = data.ID_SPTour;
            Session["TourInfo"] = idSPTOUR;
            if (Session["IDUser"] != null)
            {
                int idUser = (int)Session["IDUser"];
                var user = db.KHACHHANGs.FirstOrDefault(u => u.ID_KH == idUser);
                ViewBag.KhachHang = user;   
            }
            else if (data.SoNguoi <= 0)
            {
                ViewBag.Noti = "Hết số lượng chỗ ngồi";
            }
            return View(data);
        }
        [HttpPost]
        public ActionResult DatTour(FormCollection form, string id)
        {
            HOADON hOADON = new HOADON();
            var sptour = db.SPTOURs.FirstOrDefault(s => s.ID_SPTour == id);
            if (Session["IDUser"] == null)
            {
                return View(sptour);
            }
            else
            {
                int currentIDUser = (int)Session["IDUser"];
                var KHang = db.KHACHHANGs.FirstOrDefault(u => u.ID_KH == currentIDUser);
                hOADON.ID_KH = currentIDUser;

                hOADON.ID_SPTour = sptour.ID_SPTour;
                hOADON.NgayDat = DateTime.Now; //Default
                hOADON.TinhTrang = "Chưa TT"; //Default 

                hOADON.SLNguoiLon = int.Parse(form["songuoilon"]);
                hOADON.SLTreEm = int.Parse(form["sotreem"]);

                int slnguoilon = int.Parse(form["songuoilon"]);
                int sltreem = int.Parse(form["sotreem"]);
                int giaguoilon = sptour.GiaNguoiLon;
                int giatreem = sptour.GiaTreEm;
                /*int giaguoilon = int.Parse(form["gianguoilon"]);
                int giatreem = int.Parse(form["giatreem"]);*/

                //so sánh ...
                int soluong = slnguoilon + sltreem;
                int SoLuongSPTOUR = (int)sptour.SoNguoi;
                if (SoLuongSPTOUR < soluong)
                {
                    ViewBag.KhachHang = KHang;
                    ViewBag.SoLuong = "Vui lòng đặt ít hơn hoặc bằng số lượng chỗ hiện có !";
                    return View(sptour);
                }
                else
                {
                    if(KHang.Diem != null)
                    {
                        hOADON.TienKhuyenMai = (int)(KHang.Diem * 0.05);
                    }
                    else
                    {
                        hOADON.TienKhuyenMai = 0;
                    }
                    var tienkhuyenmai = hOADON.TienKhuyenMai;
                    int tongtien = (slnguoilon * giaguoilon) + (sltreem * giatreem);
                    int VAT = (int)(tongtien * 0.05);
                    hOADON.TongTienTour = tongtien;
                    hOADON.TienPhaiTra = tongtien - tienkhuyenmai + VAT;

                    SoLuongSPTOUR -= soluong;
                    sptour.SoNguoi = SoLuongSPTOUR;
                    db.Entry(sptour).State = EntityState.Modified;
                    db.HOADONs.Add(hOADON);
                    db.SaveChanges();
                }
            }
            return RedirectToAction("HoaDon", "Home", new { id = hOADON.ID_HoaDon });
        }
        public ActionResult PartialDanhMucTour()
        {
            return PartialView();
        }
        [HttpGet]
        public ActionResult DanhMucTour(string name, int? to, int? from, int page = 1)
        {
            ViewBag.TourNameList = db.TOURs.ToList();

            page = page < 1 ? 1 : page;/////
            int pageSize = 9;/////////
            var tours = from t in db.SPTOURs select t;

            if (!string.IsNullOrEmpty(name))
            {
                tours = tours.Where(x => x.TenSPTour.Contains(name));
                if (to != null)
                {
                    tours = tours.Where(x => x.TenSPTour.Contains(name) && x.GiaNguoiLon >= to);
                }
                if (from != null)
                {
                    tours = tours.Where(x => x.TenSPTour.Contains(name) && x.GiaNguoiLon <= from);
                }
            }
            else
            {
                if (to != null || from != null)
                {
                    if (to != null && from != null)
                    {
                        tours = tours.Where(x => x.GiaNguoiLon >= to && x.GiaNguoiLon <= from);
                    }
                    else if (to != null)
                    {
                        tours = tours.Where(x => x.GiaNguoiLon >= to);
                    }
                    else if (from != null)
                    {
                        tours = tours.Where(x => x.GiaNguoiLon <= from);
                    }
                }
            }
            //sắp xếp theo giá tiền
            tours = tours.OrderBy(item => item.GiaNguoiLon);

            //kiểm tra để thông báo lỗi 
            if(tours.Count() == 0)
            {
                ViewBag.MesSearch = "Không tìm thấy tour nào với lựa chọn của bạn !";
            }
            var toursPage = tours.ToPagedList(page, pageSize);
            return View(toursPage);
        }
        public ActionResult HoaDon(int id)
        {
            var data = db.HOADONs.Find(id);
            return View(data);
        }
        public ActionResult ThanhToanMomo(int id)
        {
            var idHD = db.HOADONs.Find(id);
            var url = UrlPayment(idHD);
            return View(url);
        }
        //VNPAY
        public string UrlPayment(HOADON hoadon)
        {
            var urlPayment = "";
            //Get Config info
            string vnp_Returnurl = ConfigurationManager.AppSettings["vnp_Returnurl"];
            string vnp_Url = ConfigurationManager.AppSettings["vnp_Url"];
            string vnp_TmnCode = ConfigurationManager.AppSettings["vnp_TmnCode"];
            string vnp_HashSerect = ConfigurationManager.AppSettings["vnp_HashSerect"];

            VnPayLibrary vnpay = new VnPayLibrary();

            vnpay.AddRequestData("vnp_Version", VnPayLibrary.VERSION);
            vnpay.AddRequestData("vnp_Command", "pay");
            vnpay.AddRequestData("vnp_TmnCode", vnp_TmnCode);
            vnpay.AddRequestData("vnp_Amount", (hoadon.TienPhaiTra*100).ToString());
            vnpay.AddRequestData("vnp_BankCode", "VNPAYQR");
            /*if(TypePaymentVN == 1)
            {
                vnpay.AddRequestData("vnp_BankCode", "VNPAYQR");
            }else if(TypePaymentVN == 2)
            {
                vnpay.AddRequestData("vnp_BankCode", "VNBANK");
            }else if(TypePaymentVN == 3)
            {
                vnpay.AddRequestData("vnp_BankCode", "INTCARD");
            }*/
            vnpay.AddRequestData("vnp_CreateDate", hoadon.NgayDat.ToString("yyyyMMddHHmmss"));
            vnpay.AddRequestData("vnp_CurrCode", "VND");
            vnpay.AddRequestData("vnp_IpAddr", Utils.GetIpAddress());
            vnpay.AddRequestData("vnp_Locale", "vn");
            vnpay.AddRequestData("vnp_OrderInfo", "Thanh toán đơn hàng :" + hoadon.ID_HoaDon);
            vnpay.AddRequestData("vnp_OrderType", "other");
            vnpay.AddRequestData("vnp_Returnurl", vnp_Returnurl);
            vnpay.AddRequestData("vnp_TxnRef", hoadon.ID_HoaDon.ToString());

            urlPayment = vnpay.CreateRequestUrl(vnp_Url, vnp_HashSerect);
            return urlPayment;

        }
        //test

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