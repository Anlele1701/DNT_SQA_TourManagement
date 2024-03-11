using DAPM_TOURDL;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest
{
    internal class BaseTest
    {
        protected string localHost = "https://localhost:44385";
        protected IWebDriver driver;
        protected string pathAn;
        protected string pathOfExcel;
        protected string[] newString;
        protected Compare convert;

        [SetUp]
        public void Setup()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            convert = new Compare();
            pathOfExcel = "FILETEST/Admin.xlsx";
            string currentDirectory = Directory.GetCurrentDirectory();
            pathOfExcel = Path.Combine(currentDirectory, pathOfExcel);
            Console.WriteLine(pathOfExcel);
            driver = new EdgeDriver();
        }

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
            driver.Dispose();
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
    }
}