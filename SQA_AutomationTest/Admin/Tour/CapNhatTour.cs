using AutoItX3Lib;
using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace SQA_AutomationTest.Admin.Tour
{
    internal class CapNhatTour:BaseTest
    {
        private string localHost = "https://localhost:44385";
        private string pathAn;
        private string pathOfExcel;
        private string[] newString;
        string patfOfFolderImages;
        [SetUp]
        public void Setup()
        {
            patfOfFolderImages = "C:\\Bài tập\\Bảo đảm chất lượng phần mềm\\DNT_SQA_TourManagement\\DAPM_TOURDL\\Images\\Admin\\TOUR\\";
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
        public void TestCapNhatTour()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Chỉnh sửa Tour");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            DangNhap();
            for (int i = 2; i <= worksheetCount; i++)
            {
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.Navigate().GoToUrl(localHost + "/TOURs/Index");
                driver.FindElement(By.XPath("/html/body/div[1]/div[2]/ul/li[3]/a")).Click();
                driver.FindElement(By.XPath("//*[@id=\"tourBox\"]/div[1]/div[2]/a/i")).Click();
                driver.FindElement(By.XPath("//*[@id=\"tourBox\"]/div[1]/div[2]/ul/li[1]/a")).Click();
                driver.FindElement(By.XPath("//*[@id=\"TenTour\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"GiaTour\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"MoTa\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"HinhTour\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"TinhThanh\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"LoaiTour\"]")).Clear();

                driver.FindElement(By.XPath("//*[@id=\"TenTour\"]")).SendKeys(newString[0]);
                driver.FindElement(By.XPath("//*[@id=\"GiaTour\"]")).SendKeys(newString[1]);
                driver.FindElement(By.XPath("//*[@id=\"MoTa\"]")).SendKeys(newString[2]);
                driver.FindElement(By.XPath("//*[@id=\"HinhTour\"]")).SendKeys(newString[3]);
                driver.FindElement(By.XPath("//*[@id=\"TinhThanh\"]")).SendKeys(newString[4]);
                driver.FindElement(By.XPath("//*[@id=\"LoaiTour\"]")).SendKeys(newString[5]);
                driver.FindElement(By.XPath("/html/body/div[2]/div/form/div/div[7]/div/input")).Click();
                if (driver.Url.Contains(localHost + "TOURs/Index"))
                {
                    string actual = "Hệ thống chỉnh sửa tour thành công và trả về trang Index";
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
                    string actual = "Hệ thống báo lỗi không đủ dữ liệu và không tạo tour mới";
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
