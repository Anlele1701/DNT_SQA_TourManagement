using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin.KhachHang
{
    internal class SuaTTKH : Tests
    {
        private IWebElement ClearAndSetValue(string xpath)
        {
            IWebElement elementclean = driver.FindElement(By.XPath(xpath));
            elementclean.Clear();
            return elementclean;
        }
        [Test]
        [TestCase("bngoc.hi4103@gmail.com", "17012003")]
        public void TestSuaThongTinKH(string username, string password)
        {
            driver.Manage().Window.Size = new System.Drawing.Size(1080, 1920);
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD -Sửa tài khoản KH");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            for (int i = 2; i <= worksheetCount; i++)
            {
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.Navigate().GoToUrl(localHost + "/Logging/LoginAdmin");
                driver.FindElement(By.Id("Mail_NV")).SendKeys(username);
                driver.FindElement(By.Id("MatKhau")).SendKeys(password);
                driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/ul[1]/li[5]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//a[contains(text(),'Edit')])[2]")).Click();
           
                Thread.Sleep(1000);
                ClearAndSetValue("(//input[@id='HoTen_KH'])[1]");
                ClearAndSetValue("(//input[@id='GioiTinh_KH'])[1]");
                ClearAndSetValue("(//input[@id='NgaySinh_KH'])[1]");
                ClearAndSetValue("(//input[@id='CCCD'])[1]");
                ClearAndSetValue("(//input[@id='SDT_KH'])[1]");
                ClearAndSetValue("(//input[@id='Mail_KH'])[1]");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='HoTen_KH'])[1]")).SendKeys(newString[0]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='GioiTinh_KH'])[1]")).SendKeys(newString[1]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='NgaySinh_KH'])[1]")).SendKeys(newString[2]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='MatKhau'])[1]")).SendKeys(newString[3]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='CCCD'])[1]")).SendKeys(newString[4]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='SDT_KH'])[1]")).SendKeys(newString[5]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@id='Mail_KH'])[1]")).SendKeys(newString[6]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//div[@class='col-md-offset-2 col-md-10'])[1]")).Click();
                Thread.Sleep(1000);
                if (driver.Url.Contains(localHost + "KHACHHANGs/Index"))
                {
                    string actual = "Hệ thống chỉnh sửa khách hàng thành công và trả về trang Index";
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
                    string actual = "Hệ thống báo lỗi sai format và không chỉnh sửa khách hàng ";
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
