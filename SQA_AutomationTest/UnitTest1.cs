using NUnit.Framework.Internal.Execution;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using Bytescout.Spreadsheet;
using Range = Bytescout.Spreadsheet.Range;
using Bytescout.Spreadsheet.COM;
using DAPM_TOURDL;

namespace SQA_AutomationTest
{
    public class Tests
    {
        private string localHost = "https://localhost:44385";
        private IWebDriver driver;
        private string pathAn;
        private string pathOfExcel;
        private string[] newString;
        Compare convert; //tách test data thành từng chuỗi nhỏ

        [SetUp]
        public void Setup()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            convert=new Compare();
            pathAn = "C:/Users/ADMIN/Documents/HUFLIT/NAM 3/HK2/BDCL/ĐỒ ÁN/Test.xlsx";
            pathOfExcel = "FILETEST/Test.xlsx";
            string currentDirectory = Directory.GetCurrentDirectory();
            pathOfExcel = Path.Combine(currentDirectory, pathOfExcel); //đường dẫn tuyệt đối
            Console.WriteLine(pathOfExcel);
            driver = new EdgeDriver();
        }

        #region FunctionTest

        [Test]
        public void FUNC_CL_Login()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Sheet2");
            int worksheetCount = worksheet.UsedRangeRowMax;
            for (int i = 2; i <= worksheetCount; i++)
            {
                    string cellValues = worksheet.Cell(i, 3).Value.ToString();
                    string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                    string[] newString = new string[parts.Length];
                    newString=convert.ConvertToArray(parts);
                    driver.Navigate().GoToUrl("https://localhost:44385/Home/LoginAndRegister");
                    driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys(newString[0]);
                    driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys(newString[1]);
                    driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
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
            Console.WriteLine("---");
            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();
        }

        [Test]
        public void TestLogin2()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Sheet2");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);
            for (int i = 2; i <= worksheetCount; i++)
            {
                Console.WriteLine(i); 
                string cellValues = worksheet.Cell(i, 3).Value.ToString();
                 string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                 string[] newString = convert.ConvertToArray(parts);
                 driver.Navigate().GoToUrl(localHost+"/Home/LoginAndRegister");
                 driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys(newString[0]);
                 driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys(newString[1]);
                 driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
                string expected = worksheet.Cell(i, 4).Value.ToString();
                 //
                 if (driver.Url == localHost || driver.Url.Contains(localHost+"/Home/HomePage/2"))
                 {
                    string actual = "Thông báo đăng nhập thành công, điều hướng sang trang HomePage";
                    worksheet.Cell(i, 5).Value = actual;
                    if (convert.CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 6).Value = "Passed";
                    else worksheet.Cell(i, 6).Value = "Failed";

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
            // Save document
            spreadsheet.SaveAs(pathOfExcel);
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