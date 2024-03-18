using Bytescout.Spreadsheet;
using NUnit.Framework.Internal;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin.KhachHang
{
    internal class XoaKH : Tests
    {
        [Test]
        [TestCase("bngoc.hi4103@gmail.com", "17012003")]
        public void TestXoaKH(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Xóa KH");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            for (int i = 2; i <= worksheetCount; i++)
            {
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.Navigate().GoToUrl(localHost + "/Logging/LoginAdmin");
                driver.FindElement(By.Id("Mail_NV")).SendKeys(username);
                driver.FindElement(By.Id("MatKhau")).SendKeys(password);
                driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl("https://localhost:44385/KHACHHANGs/Index");
                Thread.Sleep(1000);
                IWebElement element = driver.FindElement(By.XPath("//input[@placeholder='Tìm kiếm qua gmail hoặc tên']"));
                Thread.Sleep(2000);
                element.SendKeys(newString[0]);
                element.SendKeys(Keys.Enter);
                driver.FindElement(By.XPath("(//a[contains(text(),'Delete')])[1]")).Click();
                Thread.Sleep(2000);
                driver.FindElement(By.XPath("((//input[@value='Delete'])[1])")).Click();
                Thread.Sleep(1000);
                if(driver.Url == "https://localhost:44385/KHACHHANGs/Index")
                {
                    string actual = "Hệ thống xóa khách hàng thành công và trả về trang Index";
                    if (CompareExpectedAndActual(expected, actual)) 
                    {   worksheet.Cell(i, 4).Value = actual; 
                        worksheet.Cell(i, 5).Value = "Passed"; 
                    }
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else
                {
                    string actual = "Hệ thống ngừng xóa nhân viên và trả về trang Index";
                    worksheet.Cell(i, 4).Value = actual;
                    worksheet.Cell(i, 5).Value = "Failed";
                }
                Thread.Sleep(2000);
                // Save document
                spreadsheet.SaveAs(pathOfExcel);
                spreadsheet.Close();
            }
        }
    }
}
