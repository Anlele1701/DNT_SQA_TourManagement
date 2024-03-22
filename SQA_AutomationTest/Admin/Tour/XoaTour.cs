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
    internal class XoaTour:Tests
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
        public void TestXoaTour()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Xóa Tour");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            CL_LoggedInValidWithPara("bngoc.hi4103@gmail.com", "17012003");
            for (int i = 2; i <= worksheetCount; i++)
            {
                driver.Navigate().GoToUrl(localHost + "/TOURs/Index");
                string expected = worksheet.Cell(i, 3).Value.ToString();
                driver.FindElement(By.XPath("/html/body/div[1]/div[2]/ul/li[3]/a")).Click();
                driver.FindElement(By.XPath("//*[@id=\"tourBox\"]/div[1]/div[2]/a/i")).Click();
                driver.FindElement(By.XPath("//*[@id=\"tourBox\"]/div[1]/div[2]/ul/li[3]/a")).Click();
                driver.FindElement(By.XPath("/html/body/div[2]/div/div/form/div/input")).Click();
                if (driver.Url.Contains(localHost + "TOURs/Index"))
                {
                    string actual = "Hệ thống xóa tour thành công và trả về trang Index";
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
                    string actual = "Hệ thống báo lỗi không thể xóa tour";
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
    }
}
