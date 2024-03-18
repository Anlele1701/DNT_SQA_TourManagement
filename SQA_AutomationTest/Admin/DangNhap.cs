using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin
{
    internal class DangNhap : Tests
    {
        public void TestDangNhapAdmin()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Đăng Nhập");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            for (int i = 2; i <= worksheetCount; i++)
            {
                Console.WriteLine(i);
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.Navigate().GoToUrl(localHost + "/Logging/LoginAdmin");
                driver.FindElement(By.Id("Mail_NV")).SendKeys(newString[0]);
                driver.FindElement(By.Id("MatKhau")).SendKeys(newString[1]);
                driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
                string expected = worksheet.Cell(i, 3).Value.ToString();
                //
                string actual = "";
                if (driver.Url == localHost || driver.Url.Contains(localHost + "/NHANVIENs/GetData"))
                {
                    actual = "Hệ thống xác thực người dùng thành công và chuyển vào trang Admin";
                    worksheet.Cell(i, 4).Value = actual;
                }
                else
                {
                    if (ElementExists(By.XPath("/html/body/span/h1")))
                    {
                        actual = driver.FindElement(By.XPath("/html/body/span/h1")).Text;
                        worksheet.Cell(i, 4).Value = actual;
                    }
                    else
                    {
                        actual = "Hệ thống báo lỗi sai thông tin đăng nhập và yêu cầu nhập lại";
                        worksheet.Cell(i, 4).Value = actual;
                    }

                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";

                }
            }
            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();
        }
    }
}