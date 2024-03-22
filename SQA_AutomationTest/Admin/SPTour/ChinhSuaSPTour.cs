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
    internal class ChinhSuaSPTour:Tests
    {

        [Test]
        public void TestChinhSuaSPTour()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Chỉnh sửa SPTour");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            CL_LoggedInValidWithPara("bngoc.hi4103@gmail.com", "17012003");
            for (int i = 2; i <= worksheetCount; i++)
            {
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.Navigate().GoToUrl("https://localhost:44385/SPTOURs/Index");
                driver.FindElement(By.XPath("//*[@id=\"listBox\"]/div/table/tbody/tr[1]/td[10]/a[1]")).Click();

                driver.FindElement(By.XPath("//*[@id=\"TenSPTour\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"GiaNguoiLon\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"NgayKhoiHanh\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"NgayKetThuc\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"MoTa\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"DiemTapTrung\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"DiemDen\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"SoNguoi\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"HinhAnh\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"GiaTreEm\"]")).Clear();


                driver.FindElement(By.XPath("//*[@id=\"TenSPTour\"]")).SendKeys(newString[0]);
                driver.FindElement(By.XPath("//*[@id=\"GiaNguoiLon\"]")).SendKeys(newString[1]);
                driver.FindElement(By.XPath("//*[@id=\"NgayKhoiHanh\"]")).SendKeys(newString[2]);
                driver.FindElement(By.XPath("//*[@id=\"NgayKetThuc\"]")).SendKeys(newString[3]);
                driver.FindElement(By.XPath("//*[@id=\"MoTa\"]")).SendKeys(newString[4]);
                driver.FindElement(By.XPath("//*[@id=\"DiemTapTrung\"]")).SendKeys(newString[5]);
                driver.FindElement(By.XPath("//*[@id=\"DiemDen\"]")).SendKeys(newString[6]);
                driver.FindElement(By.XPath("//*[@id=\"SoNguoi\"]")).SendKeys(newString[7]);
                driver.FindElement(By.XPath("//*[@id=\"HinhAnh\"]")).SendKeys(newString[8]);
                driver.FindElement(By.XPath("//*[@id=\"GiaTreEm\"]")).SendKeys(newString[9]);
                driver.FindElement(By.XPath("/html/body/div[2]/div/form/div/div[8]/div/input")).Click();

                string actual = "";
                if (driver.Url.Contains(localHost + "SPTOURs/Index"))

                {
                    actual = "Hệ thống chỉnh sửa sản phẩm tour thành công và trả về trang Index";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else if (ElementExists(By.XPath("/html/body/span/h1")))
                {
                    actual = driver.FindElement(By.XPath("/html/body/span/h1")).Text;
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else
                {
                    actual = "Hệ thống báo lỗi không đủ dữ liệu và không tạo sp tour mới";
                    worksheet.Cell(i, 4).Value = actual;
                    if (expected.Contains("Hệ thống báo lỗi"))
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
    }
}
