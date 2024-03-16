using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Client
{
    internal class CLEditProfile : Tests
    {
        [Test]
        public void FUNC_CL_EditProfile()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfClientExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("CL - Chỉnh sửa TT Cá Nhân");
            int worksheetCount = worksheet.UsedRangeRowMax;
            for (int i = 2; i <= worksheetCount; i++)
            {
                string cellValues = worksheet.Cell(i, 3).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                string expected = worksheet.Cell(i, 4).Value.ToString();
                string actual = "";
                string a = "";
                string errorMes= "";
                //VIẾT CÂU LỆNH LOGIC CỦA TEST SCRIPT
                driver.Navigate().GoToUrl(localHost + "/Home/LoginAndRegister");
                driver.Manage().Window.Maximize();
                driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys("lethanhduyan@gmail.com");
                driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys("17012003");
                driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click(); Thread.Sleep(800);
                driver.FindElement(By.CssSelector(".nav-link.no-hover")).Click();
                Thread.Sleep(800);
                driver.FindElement(By.XPath("//i[@id='pen-icon']")).Click();
                driver.FindElement(By.CssSelector("#HoTen_KH")).Clear();
                driver.FindElement(By.CssSelector("#HoTen_KH")).SendKeys(newString[0]);
                driver.FindElement(By.CssSelector("#GioiTinh_KH")).Click();
                SelectElement GenderSelect = new SelectElement(driver.FindElement(By.CssSelector("#GioiTinh_KH")));
                try {           
                    GenderSelect.SelectByValue(newString[1].Trim());
                    }
                catch(NoSuchElementException)
                {
                    errorMes = "\"Giới tính chỉ Nam hoặc Nữ\"";

                }
                driver.FindElement(By.CssSelector("#SDT_KH")).Clear();
                driver.FindElement(By.CssSelector("#SDT_KH")).SendKeys(newString[3]);
                driver.FindElement(By.XPath("//input[@id='NgaySinh_KH']")).Clear();
                driver.FindElement(By.XPath("//input[@id='NgaySinh_KH']")).SendKeys(newString[4]);
                driver.FindElement(By.CssSelector("#Mail_KH")).SendKeys(newString[5]);
                if (newString[6] != "")
                {
                    driver.FindElement(By.CssSelector("#CCCD")).Clear();
                    driver.FindElement(By.CssSelector("#CCCD")).SendKeys(newString[6]);
                }
                Thread.Sleep(500);
                driver.FindElement(By.CssSelector("#MatKhau")).Clear();
                driver.FindElement(By.CssSelector("#MatKhau")).SendKeys(newString[2]);
                ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, document.body.scrollHeight - 150)");
                Thread.Sleep(500);
                driver.FindElement(By.CssSelector("input[value='Save']")).Click();
                Thread.Sleep(500);
                if (driver.Url.Contains(localHost + "/Home/ThongTinCaNhan/2"))
                {
                    actual = "Chỉnh sửa thành công và chuyển về trang Thông tin cá nhân với thông tin vừa cập nhật";
                }
                else
                {
                    a = "thất bại";
                    actual = "Chỉnh sửa thất bại và hiển thị thông báo lỗi " + errorMes.Trim();
                }
                worksheet.Cell(i, 5).Value = actual;
                if (CompareExpectedAndActual(expected, a)) worksheet.Cell(i, 6).Value = "Passed";
                else worksheet.Cell(i, 6).Value = "Failed";
                spreadsheet.SaveAs(pathOfClientExcel);
                spreadsheet.Close();
            }
        }
    }
}