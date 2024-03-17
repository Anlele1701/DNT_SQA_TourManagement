using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin.NhanVien
{
    internal class TimKiemNV : Tests
    {

        [Test]
        public void TestTimKiemNV()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Tìm Kiếm NV");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            CL_LoggedInValidWithPara("bngoc.hi4103@gmail.com", "17012003");
            for (int i = 2; i <= worksheetCount; i++)
            {
                driver.Navigate().GoToUrl(localHost+"/NHANVIENs/Index");
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.FindElement(By.XPath("/html/body/div[1]/div[2]/ul/li[2]/a")).Click();
                IWebElement element = driver.FindElement(By.XPath("/html/body/div[2]/div/div[1]/form/input"));
                element.SendKeys(newString[0]);
                element.SendKeys(Keys.Enter);
                string actual="";
                if (newString[0].Length > 50)
                {
                    actual = "Hệ thống báo lỗi ký tự > 50";
                }
                if (ElementExists(By.XPath("//*[@id=\"listBox\"]/div/table/tbody/tr[1]")))
                {
                    actual = "Hệ thống trả về dữ liệu nhân viên tìm kiếm";
                    worksheet.Cell(i, 4).Value = actual;
                }
                else
                {
                    actual = "Hệ thống trả về dữ liệu trống và báo không có dữ liệu";
                    worksheet.Cell(i, 4).Value = actual;
                }

                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                else worksheet.Cell(i, 5).Value = "Failed";
            }
            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();
        }
    }
}