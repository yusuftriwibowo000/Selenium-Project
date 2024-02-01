using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using System;
using System.Collections.ObjectModel;
using System.IO;
using SeleniumNew;
using LibraryPDF;
namespace SeleniumNew
{

    public class Tests
    {
        public static IWebDriver driver = SingletonDriver.GetDriver();
        Actions action = new Actions(driver);
        public static List<string> screenshotPaths = new List<string>();

        [OneTimeSetUp]
        public void SetUp()
        {
            LibPDF.InitializeDocument(); // Initialize document before tests
            LibPDF.CreateCover();
        }

        [Test]
        public void TestLogin()
        {
            driver.Navigate().GoToUrl("https://www.google.com/");
            Thread.Sleep(1000);
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);
            LibPDF.CaptureScreen(screenshotPaths, "Open Google.com", "Passed");
            Thread.Sleep(2000);
            IWebElement searchGoogle = driver.FindElement(By.Id("APjFqb"));
            searchGoogle.SendKeys("Christian Lenglolo");
            Thread.Sleep(2000);
            LibPDF.CaptureScreen(screenshotPaths, "Search di Google", "Passed");
            searchGoogle.SendKeys(Keys.Enter);
            Thread.Sleep(2000);
            LibPDF.CaptureScreen(screenshotPaths, "Hasil dari Search", "Passed");
            Thread.Sleep(1000);
            action.Click().Perform();
            Thread.Sleep(1000);
            action.SendKeys(Keys.PageDown).Perform();
            Thread.Sleep(1000);
            LibPDF.CaptureScreen(screenshotPaths, "Hasil dari Search (2)", "Passed");
        }

       [TearDown]
       public void Close()
       {
            LibPDF.GeneratePDF();
            
            // Close Browser
            driver.Quit();
        }
    }
}