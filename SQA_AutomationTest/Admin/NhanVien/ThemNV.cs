using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin.NhanVien
{
    internal class ThemNV : BaseTest
    {


        public void DangNhap()
        {
            driver.Navigate().GoToUrl(localHost + "/Logging/LoginAdmin");
            driver.FindElement(By.Id("Mail_NV")).SendKeys("bngoc.hi4103@gmail.com");
            driver.FindElement(By.Id("MatKhau")).SendKeys("17012003");
            driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
        }

        [Test]
        public void TestThemNV()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Thêm NV");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            DangNhap();
            for (int i = 2; i <= worksheetCount; i++)
            {
                driver.Navigate().GoToUrl(localHost + "/NHANVIENs/Index");
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.FindElement(By.XPath("/html/body/div[1]/div[2]/ul/li[2]/a")).Click();
                driver.FindElement(By.XPath("/html/body/div[2]/div/div[1]/button")).Click();
                driver.FindElement(By.XPath("//*[@id=\"HoTen_NV\"]")).SendKeys(newString[0]);
                driver.FindElement(By.XPath("//*[@id=\"GioiTinh_NV\"]")).SendKeys(newString[1]);
                driver.FindElement(By.XPath("//*[@id=\"NgaySinh_NV\"]")).SendKeys(newString[2]);
                driver.FindElement(By.XPath("//*[@id=\"MatKhau\"]")).SendKeys(newString[3]);
                driver.FindElement(By.XPath("//*[@id=\"Mail_NV\"]")).SendKeys(newString[4]);
                driver.FindElement(By.XPath("//*[@id=\"ChucVu\"]")).SendKeys(newString[5]);
                driver.FindElement(By.XPath("//*[@id=\"SDT_NV\"]")).SendKeys(newString[6]);
                driver.FindElement(By.CssSelector("body > div.main-content > div > form > div > div.form-group > div > input")).Click();
                if (driver.Url.Contains(localHost + "NHANVIENs/Index"))
                {
                    string actual = "Hệ thống tạo nhân viên thành công và trả về trang Index";
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
                    string actual = "Hệ thống báo lỗi không đủ dữ liệu và không tạo nhân viên mới";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
            }
            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();
        }
    }
}