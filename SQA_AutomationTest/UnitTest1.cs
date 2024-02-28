using OpenQA.Selenium;
using OpenQA.Selenium.Edge;

namespace SQA_AutomationTest
{
    public class Tests
    {
        private IWebDriver driver;

        [SetUp]
        public void Setup()
        {
            driver = new EdgeDriver();
        }

        [Test]
        public void TestURL()
        {
            driver.Navigate().GoToUrl("https://localhost:44385/");
        }

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
            driver.Dispose();
        }
    }
}