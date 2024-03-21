using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Client
{
    internal class CLTimKiem : Tests
    {
        [Test]
        public void FUNC_CL_SEARCH()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfClientExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("CL - Tìm kiếm tour");
            int worksheetCount = worksheet.UsedRangeRowMax;
            for (int i = 2; i <= worksheetCount; i++)
            {
                string cellValues = worksheet.Cell(i, 3).Value.ToString();
                Console.WriteLine(cellValues);
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                string expected = worksheet.Cell(i, 4).Value.ToString();
                string actual = "";
                string errorMes1 = "";
                string errorMes2 = "";
                string errorMes3 = "";

                //VIẾT CÂU LỆNH LOGIC CỦA TEST SCRIPT
                driver.Navigate().GoToUrl(localHost + "/Home/DanhMucTour");

                //tên tour
                SelectElement tourSelect = new SelectElement(driver.FindElement(By.CssSelector("#tourNames")));
                try
                {
                    tourSelect.SelectByValue(newString[0].Trim());
                }
                catch (NoSuchElementException)
                {
                    errorMes1 = "không có tour phù hợp";
                }

                //giá khởi điểm
                SelectElement firstPriceSelect = new SelectElement(driver.FindElement(By.CssSelector("select[name='to']")));
                try
                {
                    firstPriceSelect.SelectByValue(newString[1].Trim());
                }
                catch (NoSuchElementException)
                {
                    errorMes2 = "không có mức giá phù hợp";
                }

                //giá kết thúc
                SelectElement secondPriceSelect = new SelectElement(driver.FindElement(By.CssSelector("select[name='from']")));
                try
                {
                    secondPriceSelect.SelectByValue(newString[2].Trim());
                }
                catch (NoSuchElementException)
                {
                    errorMes3 = "không có mức giá phù hợp";
                }

                //ấn nút Search
                driver.FindElement(By.CssSelector("input[value='Search']")).Click();

                //verify
                Thread.Sleep(500);

                if (ElementExists(By.CssSelector(".alert.alert-warning.text-center")))
                {
                    actual = driver.FindElement(By.CssSelector(".alert.alert-warning.text-center")).Text + errorMes1 + errorMes2 + errorMes3;
                    worksheet.Cell(i, 5).Value = actual;
                }
                else
                {
                    actual = "Hiển thị các tour phù hợp với thông tin được nhập trong phần Tìm Kiếm";
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
