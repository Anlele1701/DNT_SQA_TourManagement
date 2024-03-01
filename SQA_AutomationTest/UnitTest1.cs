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

        public void TestCompare(ref string[] newString, string[] parts)
        {
            for (int j = 0; j < parts.Length; j++)
            {
                if (parts[j] == "null")
                {
                    newString[j] = "";
                }
                else
                {
                    newString[j] = parts[j];
                }
                Console.WriteLine(newString[j]);
            }
        }

        [Test]
        public void FUNC_CL_Login()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathAn}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Sheet2");
            Console.WriteLine(worksheet.Name);
            int worksheetCount = worksheet.UsedRangeRowMax;
            for (int i = 2; i <= worksheetCount; i++)
            {
                if (worksheet.Cell(i, 3).Merged)
                {
                    Range objRange = worksheet.Cell(i, 3).MergedWith;
                    string mergedCellValue = Convert.ToString(worksheet.Cell(objRange.Row, objRange.LeftColumnIndex).Value);
                    string[] parts = mergedCellValue.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                    string[] newString = new string[parts.Length];
                    if (objRange.Row == i)
                    {
                        TestCompare(ref newString, parts);
                        driver.Navigate().GoToUrl("https://localhost:44385/Home/LoginAndRegister");
                        driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys(newString[0]);
                        driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys(newString[1]);
                        driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
                        //
                        if (driver.Url == ("https://localhost:44385") || driver.Url.Contains("Home/HomePage/2"))
                        {
                            Console.WriteLine("Đăng nhập thành công");
                            worksheet.Cell(i, 5).Value = "Thông báo đăng nhập thành công";
                        }
                        else
                        {
                            try
                            {
                                string errorMsg = driver.FindElement(By.XPath("//span[@class='field-validation-error text-danger']")).Text;
                                worksheet.Cell(i, 5).Value = errorMsg;
                            }
                            catch (NoSuchElementException)
                            {
                                Console.WriteLine("Chưa điền đủ thông tin đăng nhập");
                                worksheet.Cell(i, 5).Value = "Chưa điền đầy đủ thông tin đăng nhập";
                            }
                        }
                    }
                }
                else
                {
                    string cellValues = worksheet.Cell(i, 3).Value.ToString();
                    string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                    string[] newString = new string[parts.Length];
                    TestCompare(ref newString, parts);
                    driver.Navigate().GoToUrl("https://localhost:44385/Home/LoginAndRegister");
                    driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys(newString[0]);
                    driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys(newString[1]);
                    driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
                    Thread.Sleep(1000);
                    //
                    if (driver.Url == ("https://localhost:44385") || driver.Url.Contains("https://localhost:44385/Home/HomePage/2"))
                    {
                        Console.WriteLine("Đăng nhập thành công");
                        worksheet.Cell(i, 5).Value = "Thông báo đăng nhập thành công";
                    }
                    else
                    {
                        try
                        {
                            string errorMsg = driver.FindElement(By.XPath("//span[@class='field-validation-error text-danger']")).Text;
                            worksheet.Cell(i, 5).Value = errorMsg;
                        }
                        catch (NoSuchElementException)
                        {
                            worksheet.Cell(i, 5).Value = "Chưa điền đầy đủ thông tin đăng nhập";
                            Console.WriteLine("Không tìm thấy thông báo lỗi, có thể đăng nhập thành công hoặc có lỗi khác xảy ra");
                        }
                    }
                }
            }
            Console.WriteLine("---");
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