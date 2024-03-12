using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Client
{
    internal class CLEditProfile : Tests
    {
        [Test]
        public void FUNC_CL_EditProfile()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfClientExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("CL - Chỉnh sửa TT Cá Nhân");
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
                CL_LoggedInValid();
                driver.FindElement(By.CssSelector(".nav-link.no-hover")).Click();
                driver.FindElement(By.XPath("//i[@id='pen-icon']")).Click();
                //if (convert.CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 6).Value = "Passed";
                //else worksheet.Cell(i, 6).Value = "Failed";
            }
            // Save document
            spreadsheet.SaveAs(pathOfClientExcel);
            spreadsheet.Close();
        }

        public void CL_LoggedInValid()
        {
            driver.Navigate().GoToUrl(localHost + "/Home/LoginAndRegister");
            driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys("lethanhduyan@gmail.com");
            driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys("17012003");
            driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
        }
    }
}