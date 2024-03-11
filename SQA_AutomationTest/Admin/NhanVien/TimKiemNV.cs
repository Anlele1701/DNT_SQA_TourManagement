﻿using Bytescout.Spreadsheet;
using DAPM_TOURDL;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin.NhanVien
{
    internal class TimKiemNV
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
        [TestCase("bngoc.hi4103@gmail.com", "17012003")]
        public void TestTimKiemNV(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Tìm Kiếm NV");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            for(int i=2; i<=worksheetCount; i++)
            {
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = convert.ConvertToArray(parts);
                driver.Navigate().GoToUrl(localHost + "/Logging/LoginAdmin");
                driver.FindElement(By.Id("Mail_NV")).SendKeys(username);
                driver.FindElement(By.Id("MatKhau")).SendKeys(password);
                driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
                driver.FindElement(By.XPath("/html/body/div[1]/div[2]/ul/li[2]/a")).Click();
                IWebElement element = driver.FindElement(By.XPath("/html/body/div[2]/div/div[1]/form/input"));
                element.SendKeys(newString[0]);
                element.SendKeys(Keys.Enter);
                if (newString[0].Length > 50)
                {
                    string actual = "Hệ thống báo lỗi ký tự > 50";
                    if (convert.CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                if (ElementExists(By.XPath("//*[@id=\"listBox\"]/div/table/tbody/tr[1]")))
                {
                    string actual = "Hệ thống trả về dữ liệu nhân viên tìm kiếm";
                    worksheet.Cell(i, 4).Value = actual;
                    if (convert.CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else
                {
                    string actual = "Hệ thống trả về dữ liệu trống và báo không có dữ liệu";
                    worksheet.Cell(i, 4).Value = actual;
                    if (convert.CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                // Save document
                spreadsheet.SaveAs(pathOfExcel);
                spreadsheet.Close();
            }
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