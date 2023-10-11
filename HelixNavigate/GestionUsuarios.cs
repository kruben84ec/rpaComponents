using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;



namespace helixIntegration
{
    class GestionUsuarios : HelperRpa
    {
        private IWebDriver driverInterface;

        public GestionUsuarios(IWebDriver driver) : base(driver) => driverInterface = driver;

        internal bool gestionUsuario(IWebDriver driver, bool isAccess)
        {
            if (isAccess)
            {
                Console.WriteLine("Abrir el catalogo");
                //var tramiteUrl = @"https://dceservice-dwp.onbmc.com/dwp/app/#/catalog/section/332;type=SBE;providerSourceName=SBE";
                
                // 2023-05-22
                var tramiteUrl = @"https://dceservice-dwp.onbmc.com/dwp/app/#/catalog/section/933;type=SBE;providerSourceName=SBE";
                driver.Navigate().GoToUrl(tramiteUrl);

                try
                {
                    Thread.Sleep(5000);

                    // var pageCreateService = @"/html/body/dwp-root/dwp-main-layout/div/main/dwp-immersive/div/div/div[2]/section/div/div[1]/dwp-tombstone-card/dwp-large-card/div/div[1]/dwp-icon-media/div/div/div[2]/dwp-card-title/div";
                    var pageCreateService = @"//*[contains(text(),'Gestión de usuarios bancos asociados (AS400)')]";

                    driver.FindElement(By.XPath(pageCreateService)).Click();
                    Thread.Sleep(4000);
                    return true;

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return false;
                }


            }
            return false;
        }
    }
}
