using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActualizarReqSM.NavigatorSM
{
    class Locator
    {
        public void ExplicitWait(ChromeDriver driver, string elementXpaht)
        {
            WebDriverWait wait = new(driver, TimeSpan.FromSeconds(10));

            wait.Until(ExpectedConditions.ElementExists(By.XPath("")));
            
        }

        public static IWebElement FluentWait(ChromeDriver driver, string elementXpath)
        {
            DefaultWait<IWebDriver> fluentWait = new DefaultWait<IWebDriver>(driver){
            Timeout = TimeSpan.FromSeconds(10),
            PollingInterval = TimeSpan.FromMilliseconds(250),
            Message = "Elemento no encontrado"
            };

            fluentWait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            
            IWebElement iwelement = fluentWait.Until(x => x.FindElement(By.XPath(elementXpath)));

            return iwelement;

        }
    }
}
