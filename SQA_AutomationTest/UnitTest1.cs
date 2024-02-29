using NUnit.Framework.Internal.Execution;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using Bytescout.Spreadsheet;
using Range = Bytescout.Spreadsheet.Range;
using Bytescout.Spreadsheet.COM;

namespace SQA_AutomationTest
{
    public class Tests
    {
        private IWebDriver driver;
        private string pathAn;

        [SetUp]
        public void Setup()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            pathAn = "C:/Users/ADMIN/Documents/HUFLIT/NAM 3/HK2/BDCL/ĐỒ ÁN/Test.xlsx";
            driver = new EdgeDriver();
        }

        #region FunctionTest

        //public void UpdateValue(int i, string part, Worksheet worksheet)
        //{
        //    Console.WriteLine(part.Trim());
        //    worksheet.Cell(i, 5).Value = "Fat pig";
        //    //Code logic
        //    //Login

        //    //Actual result so sanh expected Result
        //}

        //public void ReadData(Worksheet worksheet, int worksheetCount)
        //{
        //    for (int i = 2; i < worksheetCount; i++)
        //    {
        //        if (worksheet.Cell(i, 3).Merged)
        //        {
        //            Range objRange = worksheet.Cell(i, 3).MergedWith;
        //            string mergedCellValue = Convert.ToString(worksheet.Cell(objRange.Row, objRange.LeftColumnIndex).Value).Trim();
        //            string[] parts = mergedCellValue.Split('\n');
        //            string a = null;
        //            foreach (string part in parts)
        //            {
        //                if (objRange.Row == i)
        //                {
        //                    UpdateValue(i, part, worksheet);
        //                }
        //            }
        //        }
        //        else
        //        {
        //            string cellValues = worksheet.Cell(i, 3).Value.ToString();
        //            string[] parts = cellValues.Split('\n');
        //            string a = null;
        //            foreach (string part in parts)
        //            {
        //                UpdateValue(i, part, worksheet);
        //            }
        //        }
        //        Console.WriteLine("---");
        //    }
        //}

        //public void SaveExcel(Spreadsheet spreadsheet)
        //{
        //    spreadsheet.SaveAs(@$"{pathAn}");
        //    spreadsheet.Close();
        //}
        [Test]
        public void AutomationTest()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathAn}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName($"CL - Đăng Nhập");
            Console.WriteLine(worksheet.Name);
            int worksheetCount = worksheet.UsedRangeRowMax;

            for (int i = 2; i < worksheetCount; i++)
            {
                string mail = null;
                string password = null;
                if (worksheet.Cell(i, 3).Merged)
                {
                    Range objRange = worksheet.Cell(i, 3).MergedWith;
                    string mergedCellValue = Convert.ToString(worksheet.Cell(objRange.Row, objRange.LeftColumnIndex).Value).Trim();
                    string[] parts = mergedCellValue.Split('\n');
                    foreach (string part in parts)
                    {
                        if (objRange.Row == i)
                        {
                            mail = parts[0];
                            password = parts[1];
                            Console.WriteLine(part.Trim());
                            driver.Navigate().GoToUrl("https://localhost:44385/Home/LoginAndRegister");
                            driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys(mail);
                            driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys(password);
                            driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
                            worksheet.Cell(i, 5).Value = "peongox";
                        }
                    }
                }
                else
                {
                    string cellValues = worksheet.Cell(i, 3).Value.ToString();

                    string[] parts = cellValues.Split('\n');
                    string[] newString = new string[parts.Length];
                    for (int j = 0; j < parts.Length; j++)
                    {
                        if (parts[j] == " ")
                        {
                            newString[j] = "";
                        }
                        else
                        {
                            newString[j] = parts[j];
                        }
                    }
                    //0 - email
                    //1 - null;

                    driver.Navigate().GoToUrl("https://localhost:44385/Home/LoginAndRegister");
                    driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys(newString[j]);
                    driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys(newString[j + 1]);
                    driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
                    worksheet.Cell(i, 5).Value = "peongox";

                    //foreach (string part in parts)
                    //{
                    //    if (parts[0] == " ")
                    //    {
                    //        mail = "";
                    //    }
                    //    else
                    //    {
                    //        mail = parts[0];
                    //    }
                    //    if (parts[1] == " ")
                    //    {
                    //        password = "";
                    //    }
                    //    else password = parts[1];
                    //    Console.WriteLine(part.Trim());
                    //    driver.Navigate().GoToUrl("https://localhost:44385/Home/LoginAndRegister");
                    //    driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys(mail);
                    //    driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys(password);
                    //    driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
                    //    worksheet.Cell(i, 5).Value = "peongox";
                    //}
                }
                Console.WriteLine("---");
            }

            // Save document
            spreadsheet.SaveAs(@"C:\Users\ADMIN\Documents\HUFLIT\NAM 3\HK2\BDCL\ĐỒ ÁN\Test.xlsx");
            spreadsheet.Close();
        }

        #endregion FunctionTest

        [Test]
        public void LoginTest()
        {
            //AutomationTest("CL - Đăng Nhập");
            //driver.Navigate().GoToUrl("https://localhost:44385/Home/LoginAndRegister");
            //driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys();
        }

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
            driver.Dispose();
        }
    }
}