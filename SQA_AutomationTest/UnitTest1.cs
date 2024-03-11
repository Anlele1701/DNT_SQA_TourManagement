using NUnit.Framework.Internal.Execution;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using Bytescout.Spreadsheet;
using DAPM_TOURDL;

namespace SQA_AutomationTest
{
    public class Tests
    {
        private string localHost = "https://localhost:44385";
        private IWebDriver driver;
        private string pathOfClientExcel;
        private string pathOfAdminExcel;
        private Compare convert; //tách test data thành từng chuỗi nhỏ

        [SetUp]
        public void Setup()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            convert = new Compare();
            string currentDirectory = Directory.GetCurrentDirectory();
            pathOfClientExcel = "FILETEST/ClientTest.xlsx";
            pathOfAdminExcel = "FILETEST/AdminTest.xlsx";
            pathOfClientExcel = Path.Combine(currentDirectory, pathOfClientExcel); //đường dẫn tuyệt đối
            pathOfAdminExcel = Path.Combine(currentDirectory, pathOfAdminExcel);
            Console.WriteLine(pathOfClientExcel);
            driver = new EdgeDriver();
        }

        #region CLIENT

        #region CL Shared Function

        public void CL_LoggedInValid()
        {
            driver.Navigate().GoToUrl(localHost + "/Home/LoginAndRegister");
            driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='Mail_KH']")).SendKeys("lethanhduyan@gmail.com");
            driver.FindElement(By.XPath("//form[@action='/Login']//input[@id='MatKhau']")).SendKeys("17012003");
            driver.FindElement(By.XPath("//button[@type='submit'][contains(text(),'Đăng Nhập')]")).Click();
        }

        #endregion CL Shared Function

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
                string[] newString = convert.ConvertToArray(parts);
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
                if (convert.CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 6).Value = "Passed";
                else worksheet.Cell(i, 6).Value = "Failed";
            }
            // Save document
            spreadsheet.SaveAs(pathOfClientExcel);
            spreadsheet.Close();
        }

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
                string[] newString = convert.ConvertToArray(parts);
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

        #endregion CLIENT

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
            driver.Dispose();
        }
    }
}