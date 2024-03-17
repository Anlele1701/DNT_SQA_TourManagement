using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin.KhachHang
{
    internal class ThemKH :  Tests
    {
        [Test]
        [TestCase("bngoc.hi4103@gmail.com", "17012003")]
        public void TestThemKH(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Thêm KH ");
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
                driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/ul[1]/li[5]")).Click();
                //driver.Navigate().GoToUrl("https://localhost:44385/KHACHHANGs/Index");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//button[@class='btnCreate btn btn-dark '])[1]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//input[@id='HoTen_KH']")).SendKeys(newString[0]);
                Thread.Sleep(1000);
                if (newString[1].Contains("Nam"))
                {
                    driver.FindElement(By.XPath("(//input[@id='GioiTinh_KH'])[1]")).SendKeys(newString[1]);
                }
                else if (newString[1].Contains ("Nữ"))
                {
                    driver.FindElement(By.XPath("(//input[@id='GioiTinh_KH'])[2]")).SendKeys(newString[1]);
                }
                else { }
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='NgaySinh_KH'])[1]")).SendKeys(newString[2]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//input[@id='MatKhau']")).SendKeys(newString[3]);
                driver.FindElement(By.XPath("//input[@id='CCCD']")).SendKeys(newString[4]);
                driver.FindElement(By.XPath("//input[@id='SDT_KH']")).SendKeys(newString[5]);
                driver.FindElement(By.XPath("//input[@id='Mail_KH']")).SendKeys(newString[6]);
                driver.FindElement(By.XPath("(//input[@value='Create'])[1]")).Click();
                if (driver.Url.Contains(localHost + "KHACHHANGs/Index"))
                {
                    string actual = "Hệ thống tạo khách hàng thành công và trả về trang Index";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else if (ElementExists(By.XPath("/html/body/span/h1")))
                {
                    string actual = driver.FindElement(By.XPath("/html/body/span/h1")).Text;
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else
                {
                    string actual = "Hệ thống báo lỗi không đủ dữ liệu và không tạo khách hàng mới";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                // Save document
                spreadsheet.SaveAs(pathOfExcel);
                spreadsheet.Close();
            }
        }
    }
}
