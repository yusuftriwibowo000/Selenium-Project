using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using System;
using System.Collections.ObjectModel;
using System.IO;
using SeleniumNew;
using LibraryPDF;
using LibraryExcel;
namespace SeleniumNew
{
    public class Tests
    {
        public static IWebDriver driver = SingletonDriver.GetDriver();
        IWebElement element;
        Actions action = new Actions(driver);
        public static List<string> screenshotPaths = new List<string>();
        public static string excelFilePath = LibPDF.projectDir + "/Excel/TC001_Login.xlsx";
        public static string excelSheetName = "TC001";

        [OneTimeSetUp]
        public void SetUp()
        {
            LibPDF.InitializeDocument(excelFilePath, excelSheetName); // Initialize document before tests
            LibPDF.CreateCover();
        }

        [Test]
        public void TestLogin()
        {
            driver.Navigate().GoToUrl(LibExcel.GetDataExcel(LibPDF.globalExcelFilePath,"URL","Global"));
            Thread.Sleep(1000);
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);
            LibPDF.CaptureScreen(screenshotPaths, "Open Google.com", "Passed");
            Thread.Sleep(2000);
            element = driver.FindElement(By.Id("APjFqb"));
            element.SendKeys(LibExcel.GetDataExcel(excelFilePath, "DATA_SEARCH", excelSheetName));
            Thread.Sleep(2000);
            LibPDF.CaptureScreen(screenshotPaths, "Search di Google", "Done");
            element.SendKeys(Keys.Enter);
            Thread.Sleep(2000);
            LibPDF.CaptureScreen(screenshotPaths, "Hasil dari Search", "Passed");
            Thread.Sleep(1000);

            //bool isElementExist = false;
            element = driver.FindElement(By.XPath("//cite[@class='qLRx3b tjvcx GvPZzd cHaqb' and text()='https://en.wikipedia.org']"));
            if (element.Displayed)
            {
                element.Click();
                //isElementExist = true;
            }
            //else
            //{
            //    action.Click().Perform();
            //    Thread.Sleep(1000);

            //    int i = 2;
            //    while (!isElementExist)
            //    {
            //        action.SendKeys(Keys.PageDown).Perform();
            //        Thread.Sleep(1000);
            //        LibPDF.CaptureScreen(screenshotPaths, "Hasil dari Search "+i, "Passed");
            //        i++;

            //        if(element.Displayed)
            //        {
            //            element.Click();
            //            isElementExist = true;
            //        }
            //    }
            //}
            Thread.Sleep(2000);
            element = driver.FindElement(By.XPath("//span[@class='mw-page-title-main']"));
            if (element.Displayed)
            {
                LibPDF.CaptureScreen(screenshotPaths, "Halaman Profil " + LibExcel.GetDataExcel(excelFilePath, "DATA_SEARCH", excelSheetName), "Passed");
                action.Click().Perform();
                Thread.Sleep(1000);
                action.SendKeys(Keys.PageDown).Perform();
                Thread.Sleep(1000);
                LibPDF.CaptureScreen(screenshotPaths, "Halaman Profil " + LibExcel.GetDataExcel(excelFilePath, "DATA_SEARCH", excelSheetName) + "(2)", "Passed");
            }
            //action.Click().Perform();
            //Thread.Sleep(1000);
            //action.SendKeys(Keys.PageDown).Perform();
            //Thread.Sleep(1000);
            //LibPDF.CaptureScreen(screenshotPaths, "Hasil dari Search (2)", "Passed");
        }

        [TearDown]
        public void Close()
        {
            LibPDF.GeneratePDF(excelFilePath, excelSheetName);

            // Close Browser
            driver.Quit();
        }
    }
}