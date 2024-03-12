using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Client
{
    internal class CLDangNhap : Tests
    {
        [Test]
        public void FUNC_CL_Login()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfClientExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("CL - Đăng Nhập");
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
                driver.Navigate().GoToUrl(localHost + "/Home/LoginAndRegister");
                driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys(newString[0]);
                driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys(newString[1]);
                driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
                if (driver.Url == (localHost + "/Home/HomePage/2"))
                {
                    actual = "Thông báo đăng nhập thành công, điều hướng sang trang HomePage";
                    worksheet.Cell(i, 5).Value = actual;
                }
                else
                {
                    try
                    {
                        actual = driver.FindElement(By.XPath("//span[@class='field-validation-error text-danger']")).Text;
                        worksheet.Cell(i, 5).Value = actual;
                    }
                    catch (NoSuchElementException)
                    {
                        actual = "Chưa điền đầy đủ thông tin đăng nhập";
                        worksheet.Cell(i, 5).Value = actual;
                    }
                }
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 6).Value = "Passed";
                else worksheet.Cell(i, 6).Value = "Failed";
            }
            // Save document
            spreadsheet.SaveAs(pathOfClientExcel);
            spreadsheet.Close();
        }
    }
}