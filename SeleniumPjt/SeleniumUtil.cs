using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace SeleniumPjt
{
    internal class SeleniumUtil
    {

        private SeleniumUtil() { }
        private static SeleniumUtil instance = null;
        private IWebDriver driver;
        private DefaultWait<IWebDriver> fluentWait;
        public static SeleniumUtil GetInstance()
        {
            if (instance == null)
            {
                instance = new SeleniumUtil();
                
            }
            return instance;
    
        }

        public void SwitchTo()
        {
            driver.SwitchTo().Window(driver.WindowHandles.Last());
        }

        public IWebDriver LoadChromeDriver()
        {

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            driver = new ChromeDriver(options);
            FluentWait();
            return driver;
        }

        public void QuitChromeDriver()
        {
            
            driver.Close();
            driver.Quit();
        }

        public void GoToTargetURL(string URL)
        {
            driver.Navigate().GoToUrl(URL);
        }

        public string GetTitle()
        {
            return driver.Title;
        }

        public void ClickElement(IWebElement element)
        {
            element.Click();
        }

        public void ClickElement(By locator)
        {
            driver.FindElement(locator).Click();
        }

        public IWebElement FindElement(By locator)
        {
            IWebElement element = driver.FindElement(locator);
            return element;
        }

        public IReadOnlyCollection<IWebElement> FindElements(By locator)
        {
            IReadOnlyCollection<IWebElement> elements = driver.FindElements(locator);
            return elements;
        }

        public void WaitElementTillVisible(By locator)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
            wait.Until(ExpectedConditions.ElementIsVisible(locator));
        }

        public void WaitElementToBeClickable(By locator)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
            wait.Until(ExpectedConditions.ElementToBeClickable(locator));
        }

        public void FluentWait()
        {
            fluentWait = new DefaultWait<IWebDriver>(driver);
            fluentWait.Timeout = TimeSpan.FromSeconds(5);
            fluentWait.PollingInterval = TimeSpan.FromMilliseconds(500);
            fluentWait.IgnoreExceptionTypes(typeof(NoSuchElementException), 
                                            typeof(ElementNotInteractableException));
        }

        public object JSExecute(string script, IWebElement element)
        {
            return ((IJavaScriptExecutor)driver).ExecuteScript(script, element);
        }

    }
}
