using Newtonsoft.Json;
using OpenQA.Selenium;
using System.Net.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQA_AutomationTest.Admin
{
    internal class TestResponsive : Tests
    {
        [Test]
        [TestCase("bngoc.hi4103@gmail.com", "17012003")]
        public void TestRes(string username, string password)
        {
            driver.Navigate().GoToUrl(localHost + "/Logging/LoginAdmin");
            driver.FindElement(By.Id("Mail_NV")).SendKeys(username);
            driver.FindElement(By.Id("MatKhau")).SendKeys(password);
            driver.FindElement(By.XPath("/html/body/form/div/div/button")).Click();
            Thread.Sleep(1000);
            driver.Url.Contains("https://localhost:44385/NHANVIENs/GetData");
            SetWindowSize(1920, 1080);
            Thread.Sleep(2000);
            SetWindowSize(768, 1024);
            Thread.Sleep(2000);
            TestResponsiveSizes();
            IWebElement bookedToursCountElement = driver.FindElement(By.Id("bookedToursCount"));
            Assert.That(bookedToursCountElement.Displayed, Is.True);
            IWebElement employeeCountElement = driver.FindElement(By.Id("employCount"));
            Assert.That(employeeCountElement.Displayed, Is.True);
            IWebElement totalAmountElement = driver.FindElement(By.Id("totalAmount"));
            Assert.That(totalAmountElement.Displayed, Is.True);
            TestJavaScriptCode();
        }
        private async Task<string> GetTourData()
        {
            using (var client = new HttpClient())
            {
                var response = await client.GetAsync("URL_API");
                if (response.IsSuccessStatusCode)
                {
                    var data = await response.Content.ReadAsStringAsync();
                    return data;
                }
                else
                {
                    throw new Exception("Failed to retrieve data from the server.");
                }
            }
        }
        private async void TestJavaScriptCode()
        {
            try
            {
                string chartData = await GetTourData();

                string script1 = @"
            var chartData = " + chartData + @";
            var chartSeries = [];

            for (var i = 0; i < chartData.length; i++) {
                var item = chartData[i];
                chartSeries.push({
                    name: item.name,
                    data: [item.count]
                });
            }

            Highcharts.chart('linegraph', {
                chart: {
                    type: 'bar'
                },
                title: {
                    text: 'Số lượng sản phẩm tour'
                },
                xAxis: {
                    categories: ['Số lượng']
                },
                yAxis: {
                    title: {
                        text: 'Số lượng'
                    }
                },
                plotOptions: {
                    bar: {
                        dataLabels: {
                            enabled: true
                        }
                    }
                },
                series: chartSeries
            });";

                ((IJavaScriptExecutor)driver).ExecuteScript(script1);
                Thread.Sleep(3000);
                IWebElement chartElement = driver.FindElement(By.Id("linegraph"));
                Assert.That(chartElement.Displayed, Is.True);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        private void SetWindowSize(int width, int height)
        {
            driver.Manage().Window.Size = new System.Drawing.Size(width, height);
        }
        private void TestResponsiveSizes()
        {
            TestResponsiveSize(360, 640);  
            TestResponsiveSize(1920, 1080); 
            TestResponsiveSize(1440, 900);
        }

        private void TestResponsiveSize(int width, int height)
        {
            driver.Manage().Window.Size = new System.Drawing.Size(width, height);

            System.Threading.Thread.Sleep(2000);
        }
        [Test]
        public void TestResponsiveMiniSZ()
        {
            SetWindowSize(360, 640);
            Thread.Sleep(1000);
        }

        [Test]
        public void TestResponsiveNormalSZ()
        {
            SetWindowSize(768, 1024);
            Thread.Sleep(1000);

        }
        [Test]
        public void TestResponsiveBigSZ()
        {
            SetWindowSize(1440, 900);
            Thread.Sleep(1000);
        }
    }
}
