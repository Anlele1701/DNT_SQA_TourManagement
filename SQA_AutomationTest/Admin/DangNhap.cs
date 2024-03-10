using Bytescout.Spreadsheet;
using DAPM_TOURDL;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin
{
    internal class DangNhap
    {
        private string localHost = "https://localhost:44385";
        private IWebDriver driver;
        private string pathAn;
        private string pathOfExcel;
        private string[] newString;
        Compare convert; //tách test data thành từng chuỗi nhỏ
        [SetUp]
        public void Setup()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            convert = new Compare();
            pathOfExcel = "FILETEST/Admin.xlsx";
            string currentDirectory = Directory.GetCurrentDirectory();
            pathOfExcel = Path.Combine(currentDirectory, pathOfExcel); //đường dẫn tuyệt đối
            Console.WriteLine(pathOfExcel);
            driver = new EdgeDriver();
        }

        [Test]
        public void TestDangNhapAdmin()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Đăng Nhập");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            for (int i = 2; i <= worksheetCount; i++)
            {
                Console.WriteLine(i);
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = convert.ConvertToArray(parts);
                driver.Navigate().GoToUrl(localHost + "/Logging/LoginAdmin");
                driver.FindElement(By.Id("Mail_NV")).SendKeys(newString[0]);
                driver.FindElement(By.Id("MatKhau")).SendKeys(newString[1]);
                driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
                string expected = worksheet.Cell(i, 3).Value.ToString();
                //
                if (driver.Url == localHost || driver.Url.Contains(localHost + "/NHANVIENs/GetData"))
                {
                    string actual = "Hệ thống xác thực người dùng thành công và chuyển vào trang Admin";
                    worksheet.Cell(i, 4).Value = actual;
                    if (convert.CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else
                {
                    if (ElementExists(By.XPath("/html/body/span/h1")))
                    {
                        string actual = driver.FindElement(By.XPath("/html/body/span/h1")).Text;
                        worksheet.Cell(i, 4).Value = actual;
                        if (convert.CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                        else worksheet.Cell(i, 5).Value = "Failed";
                    }
                    else
                    {
                        string actual = "Hệ thống báo lỗi sai thông tin đăng nhập và yêu cầu nhập lại";
                        worksheet.Cell(i, 4).Value = actual;
                        if (convert.CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                        else worksheet.Cell(i, 5).Value = "Failed";
                    }
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
