using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Client
{
    internal class CLDatTour : Tests
    {
        [Test]
        public void FUNC_CLDatTour()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfClientExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("CL - Đặt Tour");
            int worksheetCount = worksheet.UsedRangeRowMax;
            for (int i = 2; i <= worksheetCount; i++)
            {
                string cellValues = worksheet.Cell(i, 3).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                string expected = worksheet.Cell(i, 4).Value.ToString();
                string actual = "";
                string a = "";
                string errorMes = "";
                driver.Manage().Window.Maximize();
                if (newString.Length > 2)
                {
                    CL_LoggedInValidWithPara(newString[0], newString[1]);
                    driver.Manage().Window.Maximize();
                    driver.Navigate().GoToUrl(localHost + "/Home/DatTour/0010");
                    driver.FindElement(By.CssSelector("input[placeholder='Số lượng người lớn']")).SendKeys(newString[2]);
                    driver.FindElement(By.CssSelector("input[value='0']")).SendKeys(newString[3]);
                    ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, document.body.scrollHeight - 150)");
                    Thread.Sleep(1000);
                    driver.FindElement(By.CssSelector("#myBtn")).Click();
                }
                else
                {
                    driver.Navigate().GoToUrl(localHost+ "/Home/DatTour/0010");
                    driver.FindElement(By.CssSelector("input[placeholder='Số lượng người lớn']")).SendKeys(newString[0]);
                    driver.FindElement(By.CssSelector("input[value='0']")).SendKeys(newString[1]);
                    ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, document.body.scrollHeight - 150)");
                    Thread.Sleep(1000);
                    driver.FindElement(By.CssSelector("#myBtn")).Click();
                }
                if (ElementExists(By.XPath("//div[@id='myModal']")))
                {
                    actual = "Hiển thị Popup yêu cầu người dùng đăng nhập trước khi đặt tour";
                }
                else if(ElementExists(By.XPath("//div[@class='alert alert-warning']")))
                {
                    actual = "Đặt tour không thành công và hiển thị thông báo vui lòng đặt ít hơn số lượng hiện có";
                }
                else if (driver.Url.Contains("sandbox.vnpayment.vn/paymentv2/Transaction"))
                {
                    actual = "Đặt tour thành công và chuyển đến trang hóa đơn với các thông tin của người dùng khi đăng nhập đã được khởi tạo và thông tin đặt tour liên quan khác";
                }
                else
                {
                    actual = "Hiển thị thông báo cảnh cáo số lượng nhập không hợp lệ";
                }
                worksheet.Cell(i, 5).Value = actual;
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 6).Value = "Passed";
                else worksheet.Cell(i, 6).Value = "Failed";
                spreadsheet.SaveAs(pathOfClientExcel);
                spreadsheet.Close();
            }
        }
    }
}
