﻿using NUnit.Framework.Internal.Execution;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using Bytescout.Spreadsheet;
using DAPM_TOURDL;

namespace SQA_AutomationTest
{
    public class Tests
    {
        private string localHost = "https://localhost:44385";
        private IWebDriver driver;
        private string pathOfExcel;
        private Compare convert; //tách test data thành từng chuỗi nhỏ

        [SetUp]
        public void Setup()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            convert = new Compare();
            pathOfExcel = "FILETEST/Test.xlsx";
            string currentDirectory = Directory.GetCurrentDirectory();
            pathOfExcel = Path.Combine(currentDirectory, pathOfExcel); //đường dẫn tuyệt đối
            Console.WriteLine(pathOfExcel);
            driver = new EdgeDriver();
        }

        [Test]
        public void FUNC_CL_Login()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Sheet2");
            int worksheetCount = worksheet.UsedRangeRowMax;
            for (int i = 2; i <= worksheetCount; i++)
            {
                string cellValues = worksheet.Cell(i, 3).Value.ToString();
                Console.WriteLine(cellValues);
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = convert.ConvertToArray(parts);
                //Viet cau lenh ra
                driver.Navigate().GoToUrl(localHost + "/Home/LoginAndRegister");
                driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys(newString[0]);
                driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys(newString[1]);
                driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
                string expected = worksheet.Cell(i, 4).Value.ToString();
                if (driver.Url == (localHost + "/Home/HomePage/2"))
                {
                    string actual = "Thông báo đăng nhập thành công, điều hướng sang trang HomePage";
                    worksheet.Cell(i, 5).Value = actual;
                    if (convert.CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 6).Value = "Passed";
                    else worksheet.Cell(i, 6).Value = "Failed";
                }
                else
                {
                    try
                    {
                        string errorMsg = driver.FindElement(By.XPath("//span[@class='field-validation-error text-danger']")).Text;
                        worksheet.Cell(i, 5).Value = errorMsg;
                    }
                    catch (NoSuchElementException)
                    {
                        worksheet.Cell(i, 5).Value = "Chưa điền đầy đủ thông tin đăng nhập";
                        Console.WriteLine("Không tìm thấy thông báo lỗi, có thể đăng nhập thành công hoặc có lỗi khác xảy ra");
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
    }
}