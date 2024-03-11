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

        [SetUp]
        public void Setup()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            pathOfExcel = "FILETEST/Admin.xlsx";
            string currentDirectory = Directory.GetCurrentDirectory();
            pathOfExcel = Path.Combine(currentDirectory, pathOfExcel);
            Console.WriteLine(pathOfExcel);
            driver = new EdgeDriver();
        }

        public string[] ConvertToArray(string[] parts)
        {
            string[] newString = new string[parts.Length];
            for (int j = 0; j < parts.Length; j++)
            {
                if (parts[j].Contains("null"))
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