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
    internal class TimKiemSPTour:Tests
    {
        

        [Test]
        public void TestTimKiemTour()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Tìm kiếm SPTour");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            CL_LoggedInValidWithPara("bngoc.hi4103@gmail.com", "17012003");
            for (int i = 2; i <= worksheetCount; i++)
            {
                driver.Navigate().GoToUrl("https://localhost:44385/SPTOURs/Index");
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string inputData = worksheet.Cell(i, 2).Value.ToString();
                IWebElement element = driver.FindElement(By.XPath("/html/body/div[2]/div/div[1]/form/input"));
                element.SendKeys(inputData);
                element.SendKeys(Keys.Enter);
                string actual;
                IWebElement table = driver.FindElement(By.XPath("//*[@id=\"listBox\"]/div/table/tbody"));
                IList<IWebElement> elements = table.FindElements(By.TagName("tr"));
                if (elements.Count > 0)
                {
                    actual = "Hệ thống trả về danh sách các tour tìm kiếm";
                    worksheet.Cell(i, 4).Value = actual;
                }
                else
                {
                    actual = "Hệ thống báo lỗi không tìm thấy tour trong hệ thống và không xuất dữ liệu tour nào ra";
                    worksheet.Cell(i, 4).Value = actual;
                }
                if (expected.Contains("Hệ thống trả về") && actual.Contains("Hệ thống trả về")||
                    expected.Contains("không tìm thấy") && actual.Contains("không tìm thấy"))
                {
                    worksheet.Cell(i, 5).Value = "Passed";
                }
                else worksheet.Cell(i, 5).Value = "Failed";
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
