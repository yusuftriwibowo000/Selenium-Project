using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject
{
    public sealed class SingletonDriver
    {
        private static IWebDriver driver;

        private SingletonDriver() { }

        public static IWebDriver GetDriver()
        {
            if (driver == null)
            {
                driver = new FirefoxDriver("C:/Users/HP/source/repos/TestProject/Drivers/");
            }
            return driver;
        }
    }
}