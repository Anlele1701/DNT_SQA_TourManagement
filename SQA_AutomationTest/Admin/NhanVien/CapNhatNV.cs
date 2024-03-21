using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin.NhanVien
{
    internal class CapNhatNV:Tests
    {

        [Test]
        public void AD_TestCapNhatNV()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("AD - Sửa NV");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            CL_LoggedInValidWithPara("bngoc.hi4103@gmail.com","17012003");
            for (int i = 2; i <= worksheetCount; i++)
            {
                CL_LoggedInValidWithPara("bngoc.hi4103@gmail.com", "17012003");
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.Navigate().GoToUrl(localHost + "/NHANVIENs/Index");
                driver.FindElement(By.XPath("/html/body/div[1]/div[2]/ul/li[2]/a")).Click();
                driver.FindElement(By.XPath("//*[@id=\"listBox\"]/div/table/tbody/tr[1]/td[8]/a[1]")).Click();
                driver.FindElement(By.XPath("//*[@id=\"HoTen_NV\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"GioiTinh_NV\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"NgaySinh_NV\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"MatKhau\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"Mail_NV\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"ChucVu\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"SDT_NV\"]")).Clear();
                driver.FindElement(By.XPath("//*[@id=\"HoTen_NV\"]")).SendKeys(newString[0]);
                driver.FindElement(By.XPath("//*[@id=\"GioiTinh_NV\"]")).SendKeys(newString[1]);
                driver.FindElement(By.XPath("//*[@id=\"NgaySinh_NV\"]")).SendKeys(newString[2]);
                driver.FindElement(By.XPath("//*[@id=\"MatKhau\"]")).SendKeys(newString[3]);
                driver.FindElement(By.XPath("//*[@id=\"Mail_NV\"]")).SendKeys(newString[4]);
                driver.FindElement(By.XPath("//*[@id=\"ChucVu\"]")).SendKeys(newString[5]);
                driver.FindElement(By.XPath("//*[@id=\"SDT_NV\"]")).SendKeys(newString[6]);
                driver.FindElement(By.XPath("/html/body/div[2]/div/form/div/div[8]/div/input")).Click();
                string actual = "";
                if (driver.Url.Contains(localHost + "NHANVIENs/Index"))
                {
                    actual = "Hệ thống chỉnh sửa nhân viên thành công và trả về trang Index";
                    worksheet.Cell(i, 4).Value = actual;
                }
                else if (ElementExists(By.XPath("/html/body/span/h1")))
                {
                    actual = driver.FindElement(By.XPath("/html/body/span/h1")).Text;
                    worksheet.Cell(i, 4).Value = actual;
                }
                else
                {
                    actual = "Hệ thống báo lỗi không đủ dữ liệu";
                    worksheet.Cell(i, 4).Value = actual;
                }
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                else worksheet.Cell(i, 5).Value = "Failed";
            }
            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();
        }
    }
}
