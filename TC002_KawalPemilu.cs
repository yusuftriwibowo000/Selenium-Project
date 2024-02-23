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
using DocumentFormat.OpenXml.Presentation;
namespace KawalPemilu
{
    public class KawalPemilu
    {
        public static IWebDriver driver = SingletonDriver.GetDriver();
        IWebElement element;
        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
        Actions action = new Actions(driver);
        public static List<string> screenshotPaths = new List<string>();
        public static string excelFilePath = LibPDF.projectDir + "/Excel/TC002_KawalPemilu.xlsx";
        public static string excelSheetName = "TC002";
        private static string dt_Provinsi = LibExcel.GetDataExcel(excelFilePath, "PROVINSI", excelSheetName);

        [OneTimeSetUp]
        public void SetUp()
        {
            LibPDF.InitializeDocument(excelFilePath, excelSheetName); // Initialize document before tests
            LibPDF.CreateCover();
        }

        [Test]
        public void TestLogin()
        {
            // Buka URL
            driver.Navigate().GoToUrl(LibExcel.GetDataExcel(LibPDF.globalExcelFilePath,"URL","Global"));
            Thread.Sleep(1000);
            driver.Manage().Window.Maximize();
            Thread.Sleep(3000);
            element = driver.FindElement(By.XPath("//span[@class='kp' and text()='KawalPemilu']"));
            
            // Jika Header "KawalPemilu" Muncul
            if(element.Displayed)
            {
                LibPDF.CaptureScreen(screenshotPaths, "Halaman kawalpemilu.org", "Passed");
                Thread.Sleep(5000);
                element = driver.FindElement(By.XPath("//a[text()='" + dt_Provinsi + "']"));
                // Jika Provinsi yang dicari muncul
                if(element.Displayed)
                {
                    js.ExecuteScript("arguments[0].scrollIntoView({block:'center'});", element);
                    Thread.Sleep(2000);
                    LibPDF.CaptureScreen(screenshotPaths, "Pilih Provinsi " + dt_Provinsi, "Done");
                    js.ExecuteScript("arguments[0].click();", element);
                }
                else
                {
                    LibPDF.CaptureScreen(screenshotPaths, "Provinsi " + dt_Provinsi + " Tidak Terdapat di Halaman" , "Failed");
                }

                Thread.Sleep(4000);
                element = driver.FindElement(By.XPath("//a[@class='breadcrumb-link active' and text()='" + dt_Provinsi+ "']"));

                // Jika sudah masuk halaman Provinsi
                if (element.Displayed)
                {
                    element.Click();
                    Thread.Sleep(1000);
                    element.SendKeys(Keys.PageUp);
                    Thread.Sleep(1000);
                    LibPDF.CaptureScreen(screenshotPaths, "Halaman Rekapitulasi Suara Pilpres di Provinsi " + dt_Provinsi, "Passed");
                    element.SendKeys(Keys.PageDown);
                    Thread.Sleep(1000);
                    LibPDF.CaptureScreen(screenshotPaths, "Halaman Rekapitulasi Suara Pilpres di Provinsi " + dt_Provinsi + " (2)", "Passed");
                    Thread.Sleep(2000);
                    element = driver.FindElement(By.XPath("//td[div[@class='lokasi']/a[@class='hierarchy' and text()='HULU SUNGAI UTARA']]"));
                    int targetRow = Convert.ToInt32(js.ExecuteScript("return arguments[0].parentNode.rowIndex;", element));
                    //int targetColumn = Convert.ToInt32(js.ExecuteScript("return arguments[0].cellIndex;", element));
                    element = driver.FindElement(By.XPath($"//table[@class='collapsed_border sticky_table']/tbody[@class='data']/tr[{targetRow}]/td[5]/app-percent/span"));
                    string cakupanTPS = element.Text;
                    LibPDF.CaptureScreen(screenshotPaths, $"Kabupaten HULU SUNGAI UTARA Memiliki Cakupan TPS : " + cakupanTPS, "Done");
                }
                else
                {
                    LibPDF.CaptureScreen(screenshotPaths, "Halaman Rekapitulasi Suara Pilpres di Provinsi " + dt_Provinsi + " Gagal Dibuka", "Passed");
                }
            }
            else
            {
                LibPDF.CaptureScreen(screenshotPaths, "Halaman kawalpemilu.org Gagal Dibuka", "Failed");
                Thread.Sleep(2000);
            }
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