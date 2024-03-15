using AutoItX3Lib;
using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace SQA_AutomationTest.Admin.SPTour
{
    internal class ThemSPTour:BaseTest
    {
        private string localHost = "https://localhost:44385";
        private string pathOfExcel;
        private string[] newString;
        [SetUp]
        public void Setup()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            pathOfExcel = "FILETEST/Admin.xlsx";
            string currentDirectory = Directory.GetCurrentDirectory();
            pathOfExcel = Path.Combine(currentDirectory, pathOfExcel); //đường dẫn tuyệt đối
            Console.WriteLine(pathOfExcel);
        }
        public void DangNhap()
        {
            driver.Navigate().GoToUrl(localHost + "/Logging/LoginAdmin");
            driver.FindElement(By.Id("Mail_NV")).SendKeys("bngoc.hi4103@gmail.com");
            driver.FindElement(By.Id("MatKhau")).SendKeys("17012003");
            driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
        }

        [Test]
        public void TestThemSPTour()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Thêm SPTour");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            DangNhap();
            for (int i = 2; i <= worksheetCount; i++)
            {
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.Navigate().GoToUrl("https://localhost:44385/SPTOURs/Index");
                driver.FindElement(By.XPath("/html/body/div[2]/div/div[1]/button")).Click();
                driver.FindElement(By.XPath("//*[@id=\"ID_SPTour\"]")).SendKeys(newString[0]);
                driver.FindElement(By.XPath("//*[@id=\"TenSPTour\"]")).SendKeys(newString[1]);
                driver.FindElement(By.XPath("//*[@id=\"GiaNguoiLon\"]")).SendKeys(newString[2]);
                driver.FindElement(By.XPath("//*[@id=\"NgayKhoiHanh\"]")).SendKeys(newString[3]);
                driver.FindElement(By.XPath("//*[@id=\"NgayKetThuc\"]")).SendKeys(newString[4]);
                driver.FindElement(By.XPath("//*[@id=\"MoTa\"]")).SendKeys(newString[5]);
                driver.FindElement(By.XPath("//*[@id=\"DiemTapTrung\"]")).SendKeys(newString[6]);
                driver.FindElement(By.XPath("//*[@id=\"DiemDen\"]")).SendKeys(newString[7]);
                driver.FindElement(By.XPath("//*[@id=\"SoNguoi\"]")).SendKeys(newString[8]);
                driver.FindElement(By.XPath("//*[@id=\"HinhAnh\"]")).SendKeys(newString[9]);
                driver.FindElement(By.XPath("//*[@id=\"GiaTreEm\"]")).SendKeys(newString[10]);
                driver.FindElement(By.XPath("/html/body/div[2]/div/form/div/div[8]/div/input")).Click();


                if (driver.Url.Contains(localHost + "SPTOURs/Index"))

                {
                    string actual = "Hệ thống tạo sản phẩm tour thành công và trả về trang Index";
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
                    string actual = "Hệ thống báo lỗi không đủ dữ liệu và không tạo sp tour mới";
                    worksheet.Cell(i, 4).Value = actual;
                    if (expected.Contains("Hệ thống báo lỗi"))
                    {
                        worksheet.Cell(i, 5).Value = "Passed";
                    }
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
            }
            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();
        }

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
            driver.Dispose();
        }
        public bool ElementExists(By locator)
        {
            try
            {
                driver.FindElement(locator);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
    }
}
