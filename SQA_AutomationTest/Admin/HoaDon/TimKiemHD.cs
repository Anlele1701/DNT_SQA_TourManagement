using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin.HoaDon
{
    internal class TimKiemHD : Tests
    {
        [Test]
        [TestCase("bngoc.hi4103@gmail.com", "17012003")]
        public void TestTimKiemHD(string username, string password)
        {
            driver.Manage().Window.Size = new System.Drawing.Size(1080, 1920);
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Tìm kiếm hóa đơn");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            for (int i = 2; i <= worksheetCount; i++)
            {
                string expected = worksheet.Cell(i, 4).Value.ToString();
                string cellValues = worksheet.Cell(i, 3).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.Navigate().GoToUrl(localHost + "/Logging/LoginAdmin");
                driver.FindElement(By.Id("Mail_NV")).SendKeys(username);
                driver.FindElement(By.Id("MatKhau")).SendKeys(password);
                driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
                //CaseDangNhap();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl("https://localhost:44385/HOADONs/Index");
                Thread.Sleep(1000);
                IWebElement element = driver.FindElement(By.XPath("(//input[@placeholder='Tìm kiếm qua gmail hoặc tên'])[1]"));
                Thread.Sleep(1000);
                element.SendKeys(newString[1]);
                element.SendKeys(Keys.Enter);
                if (newString[0].Length > 50)
                {
                    string actual = "Hệ thống báo lỗi ký tự > 50";
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 6).Value = "Passed";
                    else worksheet.Cell(i, 6).Value = "Failed";
                }
                if (ElementExists(By.XPath("//*[@id=\"listBox\"]/div/table/tbody/tr[1]")))
                {
                    string actual = "Hệ thống trả về dữ liệu khách hàng tìm kiếm";
                    worksheet.Cell(i, 5).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 6).Value = "Passed";
                    else worksheet.Cell(i, 6).Value = "Failed";
                }
                else
                {
                    string actual = "Hệ thống trả về dữ liệu trống và báo không có dữ liệu";
                    worksheet.Cell(i, 5).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 6).Value = "Passed";
                    else worksheet.Cell(i, 6).Value = "Failed";
                }
                // Save document
                spreadsheet.SaveAs(pathOfExcel);
                spreadsheet.Close();
            }
        }
    }
}
