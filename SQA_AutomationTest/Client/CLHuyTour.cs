using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Client
{
    internal class CLHuyTour : Tests
    {
        [Test]
        public void FUNC_CL_CANCELTOUR()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfClientExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("CL - Hủy tour đã đặt");
            int worksheetCount = worksheet.UsedRangeRowMax;
            for (int i = 2; i <= worksheetCount; i++)
            {
                string cellValues = worksheet.Cell(i, 3).Value.ToString();
                Console.WriteLine(cellValues);
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                string expected = worksheet.Cell(i, 4).Value.ToString();
                string actual = "";

                //VIẾT CÂU LỆNH LOGIC CỦA TEST SCRIPT
                driver.Navigate().GoToUrl(localHost + "/Home/LichSuDatTour/2");
                Thread.Sleep(1000);

                string idHD = driver.FindElement(By.ClassName("fw-semibold")).GetAttribute("text");
                if (newString[0] == idHD)
                {
                    driver.FindElement(By.ClassName("btn btn-danger")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("//a[contains(text(),'Xác nhận hủy')]")).Click();
                }

                if(ElementExists(By.XPath("(//h1[normalize-space()=\"Server Error in '/' Application.\"])[1]")))
                {
                    actual = "Hệ thống bị lỗi, hủy tour thất bại";
                    worksheet.Cell(i, 5).Value = actual;
                }
                else
                {
                    actual = "Hủy tour thất bại";
                    worksheet.Cell(i, 5).Value = actual;
                }

                if (CompareExpectedAndActual(expected, actual))
                    worksheet.Cell(i, 6).Value = "Passed";
                else
                    worksheet.Cell(i, 6).Value = "Failed";
            }
            // Save document
            spreadsheet.SaveAs(pathOfClientExcel);
            spreadsheet.Close();
        }
    }
}
