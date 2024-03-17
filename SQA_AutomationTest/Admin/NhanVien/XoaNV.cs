using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin.NhanVien
{
    internal class XoaNV:Tests
    {
        [Test]
        public void TestXoaNV()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Xóa NV");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            CL_LoggedInValidWithPara("bngoc.hi4103@gmail.com", "17012003");
            for (int i = 2; i <= worksheetCount; i++)
            {
                driver.Navigate().GoToUrl(localHost+"/NHANVIENs/Index");
                string expected = worksheet.Cell(i, 3).Value.ToString();
                driver.FindElement(By.XPath("/html/body/div[1]/div[2]/ul/li[2]/a")).Click();
                driver.FindElement(By.XPath("//*[@id=\"listBox\"]/div/table/tbody/tr[1]/td[8]/a[3]")).Click();
                driver.FindElement(By.XPath("/html/body/div[2]/div/div/form/div/input")).Click();
                string actual = "";
                if(driver.Url.Contains(localHost + "NHANVIENs/Index")){
                    actual = "Hệ thống xóa nhân viên thành công và trả về trang Index";
                    worksheet.Cell(i, 4).Value = actual;
                }
                else if (ElementExists(By.XPath("/html/body/span/h1")))
                {
                    actual = driver.FindElement(By.XPath("/html/body/span/h1")).Text;
                    worksheet.Cell(i, 4).Value = actual;
                }

                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
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
