using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Client
{
    internal class CLDangKy : Tests
    {
        [Test]
        public void FUNC_CL_Register()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfClientExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("CL - Đăng Ký");
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
                //chuyen qua form dang ky
                driver.FindElement(By.XPath("//button[@id='overlayBtn']")).Click();

                //nhap ho ten
                driver.FindElement(By.CssSelector("#HoTen_KH")).SendKeys(newString[0]);

                //chon gioi tinh
                string nam = driver.FindElement(By.XPath("(//input[@id='GioiTinh_KH'])[1]")).GetAttribute("value");
                string nu = driver.FindElement(By.XPath("(//input[@id='GioiTinh_KH'])[2]")).GetAttribute("value");
                Thread.Sleep(500);
                if (newString[1] == nam)
                    driver.FindElement(By.XPath("(//input[@id='GioiTinh_KH'])[1]")).Click();
                else if (newString[1] == nu)
                    driver.FindElement(By.XPath("(//input[@id='GioiTinh_KH'])[2]")).Click();
                else
                    driver.FindElement(By.XPath("(//input[@id='GioiTinh_KH'])[1]")).Click();
                //
                
                //dien ngay sinh
                driver.FindElement(By.CssSelector("#NgaySinh_KH")).SendKeys(newString[2]);

                //dien mat khau
                driver.FindElement(By.XPath("//form[@action='/Home/DangKy']//input[@id='MatKhau']")).SendKeys(newString[3]);

                //dien cccd
                driver.FindElement(By.XPath("//input[@id='CCCD']")).SendKeys(newString[4]);

                //dien sdt
                driver.FindElement(By.XPath("//input[@id='SDT_KH']")).SendKeys(newString[5]);

                //dien email
                driver.FindElement(By.XPath("//form[@action='/Home/DangKy']//input[@id='Mail_KH']")).SendKeys(newString[6]);

                //an nut dang ky
                driver.FindElement(By.CssSelector("div[class='sign-up-btn'] button[type='submit']")).Click();

                //verify
                Thread.Sleep(500);
                if (ElementExists(By.CssSelector(".message")))
                {
                    actual = driver.FindElement(By.CssSelector(".message")).Text;
                    worksheet.Cell(i, 5).Value = actual;
                }
                else 
                {
                    try
                    {
                        actual = driver.FindElement(By.CssSelector("div[class='validation-summary-errors text-danger'] ul li")).Text;
                        worksheet.Cell(i, 5).Value = actual;
                    }
                    catch (NoSuchElementException)
                    {
                        actual = "Thông báo thiếu thông tin";
                        worksheet.Cell(i, 5).Value = actual;
                    }
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
