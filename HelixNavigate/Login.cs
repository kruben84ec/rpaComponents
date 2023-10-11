using OpenQA.Selenium;
using Serilog;


namespace helixIntegration
{
     class Login
    {

        internal bool access(IWebDriver driver, string urlAccess ="")
        {
            HelperRpa help = new HelperRpa(driver);
            //help.ConfigLog(@"E:\RECURSOS ROBOT\LOGS\MESA_SERVICIO\GESTIONDEUSUARIOS\\");

            String user = "usrbotrunner";
            String password = "BotInterdin.2002";
            if(string.IsNullOrEmpty(urlAccess))
            {
                driver.Navigate().GoToUrl(urlAccess);
            }

            Thread.Sleep(3000);
            Log.Information("Ingresando al Login");
            try
            {
                string loginUrl = "https://or-rsso1.onbmc.com/rsso/start";
                string currectUrl = driver.Url;
                Console.WriteLine(currectUrl);
                if (currectUrl == loginUrl)
                {
                    IWebElement usuarioInput = driver.FindElement(By.Id("user_login"));
                    usuarioInput.SendKeys(user);

                    IWebElement passwordInput = driver.FindElement(By.Id("login_user_password"));
                    passwordInput.SendKeys(password);

                    IWebElement bottonLogin = driver.FindElement(By.Id("login-jsp-btn"));
                    bottonLogin.Click();
                    Thread.Sleep(3000);

                    return true;

                }else
                {
                    Console.WriteLine("No pudo accesder");

                }
            }
            catch (WebDriverException error)
            {

                Console.WriteLine(error.ToString());
                return false;
            }
            return false;

        }


    }
}
