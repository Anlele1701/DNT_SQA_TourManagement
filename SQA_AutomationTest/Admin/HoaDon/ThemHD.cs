using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin.HoaDon
{
    internal class ThemHD : Tests
    {
        [Test]
        [TestCase("bngoc.hi4103@gmail.com", "17012003")]
        public void TestXoaKH(string username, string password)
        {
            driver.Manage().Window.Maximize();
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Thêm hóa đơn");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            for (int i = 2; i <= worksheetCount; i++)
            {
                string expected = worksheet.Cell(i, 4).Value.ToString();
                string cellValues = worksheet.Cell(i, 3).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.Navigate().GoToUrl(localHost + "/Logging/LoginAdmin");
                driver.FindElement(By.Id("Mail_NV")).SendKeys(username);
                driver.FindElement(By.Id("MatKhau")).SendKeys(password);
                driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl("https://localhost:44385/HOADONs/Index");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//button[@class='btnCreate btn btn-dark '])[1]")).Click();
                Thread.Sleep(1000);
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight)");
                driver.FindElement(By.XPath("(//input[@id='SLTreEm'])[1]")).SendKeys(newString[0]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='TongTienTour'])[1]")).SendKeys(newString[1]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='NgayDat'])[1]")).SendKeys(newString[2]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='TinhTrang'])[1]")).SendKeys(newString[6]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='SLNguoiLon'])[1]")).SendKeys(newString[4]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='TienKhuyenMai'])[1]")).SendKeys(newString[5]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='TienPhaiTra'])[1]")).SendKeys(newString[3]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//select[@id='ID_SPTour'])[1]")).Click();
                Thread.Sleep(1000);
                if (newString[7].Contains("Sài Gòn"))
                {
                    js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight)");
                    driver.FindElement(By.XPath("(//option[@value='0002'])[1]")).Click();
                }
                else if (newString[7].Contains("Đà Lạt"))
                {
                    driver.FindElement(By.XPath("(//option[@value='0005'])[1]")).Click();
                }
                else if (newString[7].Contains("SaPa"))
                {
                    driver.FindElement(By.XPath("(//option[@value='0009'])[1]")).Click();
                }
                else if (newString[7].Contains("Vũng Tàu"))
                {
                    driver.FindElement(By.XPath("(//option[@value='0011'])[1]")).Click();
                }
                else if (newString[7].Contains("Hà Nội"))
                {
                    driver.FindElement(By.XPath("(//option[@value='0003'])[1]")).Click();
                }
                else
                {
                    driver.FindElement(By.XPath("(//option[@value='0002'])[1]")).Click();
                }
                
                driver.FindElement(By.XPath("(//select[@id='ID_KH'])[1]")).Click();
                Thread.Sleep(1000);
                if (newString[8].Contains("Lê Thoại Bảo Ngọc"))
                {
                    js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight)");
                    driver.FindElement(By.XPath("(//option[@value='2'])[1]")).Click();
                }
                else if (newString[8].Contains("Lê Thành Duy Ân"))
                {
                    js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight)");
                    driver.FindElement(By.XPath("(//option[@value='6'])[1]")).Click();
                }
                else
                {
                    driver.FindElement(By.XPath("(//option[@value='2'])[1]")).Click();
                }
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@value='Create'])[1]")).Click();
                Thread.Sleep(1000);
                if (driver.Url == "https://localhost:44385/KHACHHANGs/Index")
                {
                    string actual = "Hệ thống tạo hóa đơn mới thành công và trở về trang Index";
                    if (CompareExpectedAndActual(expected, actual))
                    {
                        worksheet.Cell(i, 5).Value = actual;
                        worksheet.Cell(i, 6).Value = "Passed";
                    }
                    else worksheet.Cell(i, 6).Value = "Failed";
                }
                else
                {
                    string actual = "Hệ thống thông báo dữ liệu không hợp lệ, yêu cầu người dùng nhập lại";
                    worksheet.Cell(i, 5).Value = actual;
                    worksheet.Cell(i, 6).Value = "Failed";
                }
                Thread.Sleep(2000);
                // Save document
                spreadsheet.SaveAs(pathOfExcel);
                spreadsheet.Close();
            }
        }
    }
}
