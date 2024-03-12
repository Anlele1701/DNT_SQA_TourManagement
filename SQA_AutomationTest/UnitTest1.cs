using NUnit.Framework.Internal.Execution;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using Bytescout.Spreadsheet;

namespace SQA_AutomationTest
{
    public class Tests
    {
        protected string localHost = "https://localhost:44385";
        protected IWebDriver driver;
        protected string pathOfClientExcel;
        protected string pathOfAdminExcel;
        protected string pathOfExcel;
        protected string[] newString;

        [SetUp]
        public void Setup()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            string currentDirectory = Directory.GetCurrentDirectory();
            pathOfExcel = "FILETEST/Admin.xlsx";
            pathOfClientExcel = "FILETEST/ClientTest.xlsx";
            pathOfAdminExcel = "FILETEST/AdminTest.xlsx";
            pathOfClientExcel = Path.Combine(currentDirectory, pathOfClientExcel); //đường dẫn tuyệt đối
            pathOfAdminExcel = Path.Combine(currentDirectory, pathOfAdminExcel);
            pathOfExcel = Path.Combine(currentDirectory, pathOfExcel);
            Console.WriteLine(pathOfClientExcel);
            driver = new EdgeDriver();
        }

        public string[] ConvertToArray(string[] parts)
        {
            string[] newString = new string[parts.Length];
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
            return newString;
        }

        public bool CompareExpectedAndActual(string expected, string actual)
        {
            if (expected == actual) return true;
            else return false;
        }

        public bool ElementExists(By locator)
        {
            try
            {
                driver.FindElement(locator);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
            driver.Dispose();
        }
    }
}