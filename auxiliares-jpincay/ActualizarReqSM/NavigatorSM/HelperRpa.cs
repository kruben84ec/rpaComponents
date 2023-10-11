using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Text;
using Serilog;
using OpenQA.Selenium.Chrome;
using System.Runtime;

namespace ActualizarReqSM.NavigatorSM
{
    internal class DriverWeb 
    {
        public DriverWeb() { }

        public ChromeDriver GetChromeDriver()
        {
            ChromeOptions options = new ChromeOptions();
            //Set the argument 
            //options.AddArgument("--headless");
            options.AddArgument("--silent");
            options.AddArgument("--log-level=3");
            options.AddArguments("--start-maximized");
            options.AddArguments("--ignore-certificate-errors");
            options.AddArguments("--ignore-ssl-errors");
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            options.AddExcludedArgument("enable-automation");
            options.AddAdditionalChromeOption("useAutomationExtension", false);
            //Set Chrome to work with headless mode
            ChromeDriver web = new ChromeDriver(options);
            
            return web;

        }
    }

    internal class HelperRpa
    {
        private WebDriver driverInterface;

        public HelperRpa(WebDriver driver) => driverInterface = driver;

        internal void findButtonClick(string field, int timeInit, int timefisnish)
        {
            var enviarPeticionButton = driverInterface.FindElement(By.XPath(field));
            Thread.Sleep(timeInit);
            enviarPeticionButton.Click();
            Thread.Sleep(timefisnish);
        }

        internal void findFieldClickWait(string field, int secondsWait)
        {
            WebDriverWait wait = new WebDriverWait(driverInterface, TimeSpan.FromSeconds(secondsWait));
            var fieldSearch = field;
            var optionOperation = driverInterface.FindElement(By.XPath(fieldSearch));
            wait.Until(driverInterface => optionOperation);
            optionOperation.Click();
        }


        internal void ClickWaitField(string field, double secondsWait)
        {
            bool element = false;

            WebDriverWait wait = new WebDriverWait(driverInterface, TimeSpan.FromSeconds(secondsWait))
            {
                PollingInterval = TimeSpan.FromMilliseconds(250)
            };

            wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(WebDriverTimeoutException));

            var waitedElement = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(field)));

            waitedElement.Click();

            
            //element = wait.Until(drv =>
            //drv.FindElement(By.XPath(field)).Displayed);

            //if(element)
            //{
            //    driverInterface.FindElement(By.XPath(field)).Click();
            //}


        }

        internal void findFieldClick(string field)
        {
            var fieldSearch = field;
            var optionOperation = driverInterface.FindElement(By.XPath(fieldSearch));
            optionOperation.Click();
        }

        internal void FindFieldClearSetText(string field, string valueField) {

            var fieldSearch = field;
            var optionOperation = driverInterface.FindElement(By.XPath(fieldSearch));
            optionOperation.Click();
            optionOperation.Clear();
            optionOperation.SendKeys(valueField);
        }

        internal void findFieldSetText(string field, string valueField)
        {
            var fieldSearch = field;
            var optionOperation = driverInterface.FindElement(By.XPath(fieldSearch));
            optionOperation.SendKeys(valueField);
        }
        internal string findFieldGetText(string field)
        {
            var fieldSearch = field;
            var optionOperation = driverInterface.FindElement(By.XPath(fieldSearch));
            return optionOperation.GetAttribute("value");
        }

        internal string cleanString(string imputString)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char charterInput in imputString.Trim())
            {
                if ((charterInput >= '0' && charterInput <= '9') || (charterInput >= 'A' && charterInput <= 'Z') || (charterInput >= 'a' && charterInput <= 'z') || charterInput == '.' || charterInput == '_')
                {
                    sb.Append(charterInput);
                }
            }
            return sb.ToString();
        }

        public (string,  string) getFechasReporte()
        {
            string fechaReport = $@"{DateTime.Now:d/M/yyyy}";
            string fechaInicio = fechaReport + " 00:00:00";
            string fechaFin = fechaReport + " 23:59:00";
            return (fechaInicio, fechaFin);
        }

        public void ConfigLog(string path)
        {
            string fecha_log = $@"{DateTime.Now:yyyy-M-d}\";
            string logPathFinal = Path.Combine(path, fecha_log);
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File($"{logPathFinal}{System.AppDomain.CurrentDomain.FriendlyName}_{DateTime.Now:yyyyMMdd-HHmm}.log",
                                 outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Log configurado...");
        }


        public static string ValidateInputFilePath(String path)
        {
            /* 
             * valida ruta, obtiene archivo mas reciente y retorna  ruta completa
             */

            string recentFilePath = "";

            try
            {
                //Lee directorio en busqueda de archivo mas reciente
                var directory = new DirectoryInfo(path);
                recentFilePath = path;
                return recentFilePath;

            }
            catch (Exception)
            {
                throw;
            }

        }
    }
}
