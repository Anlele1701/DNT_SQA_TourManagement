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
    internal class TimKiemTour:Tests
    {
        
        [Test]
        public void TestTimKiemTour()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Tìm kiếm Tour");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            CL_LoggedInValidWithPara("bngoc.hi4103@gmail.com", "17012003");
            for (int i = 2; i <= worksheetCount; i++)
            {
                driver.Navigate().GoToUrl("https://localhost:44385/TOURs/Index");
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string inputData=worksheet.Cell(i,2).Value.ToString();
                IWebElement element= driver.FindElement(By.XPath("/html/body/div[2]/div/div[1]/form/input"));
                element.SendKeys(inputData);
                element.SendKeys(Keys.Enter);
                string actual="";
                IList<IWebElement> elements = driver.FindElements(By.ClassName("card"));
                if (elements.Count > 0)
                {
                    actual = "Hệ thống trả về danh sách các tour tìm kiếm";
                    worksheet.Cell(i, 4).Value = actual;
                }
                else
                {
                    actual = "Hệ thống trả về danh sách trống";
                    worksheet.Cell(i, 4).Value = actual;
                }
                if (expected.Contains("Hệ thống trả về"))
                {
                    worksheet.Cell(i, 5).Value = "Passed";
                }
                else worksheet.Cell(i, 5).Value = "Failed";
            }
            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();
        }
    }
}
