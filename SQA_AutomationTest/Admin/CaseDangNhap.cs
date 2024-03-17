using NUnit.Framework.Internal;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin
{
    internal class CaseDangNhap : Tests
    {
        [Test]
        [TestCase("bngoc.hi4103@gmail.com", "17012003")]
        public void DangNhapAdmin(string username, string password)
        {
            driver.Navigate().GoToUrl("https://localhost:44385/Logging/LoginAdmin");
            driver.FindElement(By.Id("Mail_NV")).SendKeys(username);
            driver.FindElement(By.Id("MatKhau")).SendKeys(password);
            driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
            Thread.Sleep(1000);
        }
    }
}
