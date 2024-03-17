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
    internal class XoaSPTour:Tests
    {
        

        [Test]
        public void TestXoaSPTour()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Xóa SPTour");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            CL_LoggedInValidWithPara("bngoc.hi4103@gmail.com", "17012003");
            for (int i = 2; i <= worksheetCount; i++)
            {
                string expected = worksheet.Cell(i, 3).Value.ToString();
                driver.Navigate().GoToUrl("https://localhost:44385/SPTOURs/Index");
                driver.FindElement(By.XPath("//*[@id=\"listBox\"]/div/table/tbody/tr[1]/td[10]/a[3]")).Click();
                Thread.Sleep(5000);
                driver.FindElement(By.XPath("/html/body/div[2]/div/div/form/div/input")).Click();
                if (driver.Url.Contains(localHost + "SPTOURs/Index"))

                {
                    string actual = "Hệ thống xóa sản phẩm tour thành công và trả về trang Index";
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
                    string actual = "Hệ thống ngừng xóa sản phẩm";
                    worksheet.Cell(i, 4).Value = actual;
                    if (expected.Contains("Hệ thống ngừng xóa sản phẩm"))
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
