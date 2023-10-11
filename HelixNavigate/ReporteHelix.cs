
using helixIntegration;
using OpenQA.Selenium;
using Serilog;
namespace HelixNavigate
{
     class ReporteHelix:HelperRpa
    {
        private IWebDriver driverInterface;
        public ReporteHelix(IWebDriver driver) : base(driver) => driverInterface = driver;
    
        internal bool accessReport(IWebDriver web)
        {
            try
            {
                Program execute = new Program();
                HelperRpa help = new HelperRpa(web);
                //help.ConfigLog(@"E:\RECURSOS ROBOT\LOGS\MESA_SERVICIO\GESTIONDEUSUARIOS\\");

                (string fechaInicio, string fechaFin) = help.getFechasReporte();
                string reportPageDowload = "https://dceservice.onbmc.com/arsys/smartreporting";


                Log.Information("Login smart reporting");
                web.Navigate().GoToUrl(reportPageDowload);
                Thread.Sleep(5000);
                Login login = new Login();
                login.access(web, "https://or-rsso1.onbmc.com/rsso/start");
                Log.Information("Abriendo la pagina de reporte");
                Thread.Sleep(10000);

                Log.Information("Click boton fav panel de reporte");
                findFieldClickWait("//li[@id='fave']/div[1]", 10);
                Thread.Sleep(5000);
                Log.Information("Click de myFavouritesScrollableArea");
                findFieldClick("//div[@id='myFavouritesScrollableArea']/div[1]/div[1]/ul[1]/li[1]/p[1]");
                Thread.Sleep(10000);

                findFieldClickWait("//*[@id=\"108566\"]/div/div[4]/div/div/div[1]/div/input", 20);
                Log.Information("Ingresar la fecha");
                Thread.Sleep(4000);

                FindFieldClearSetText("/html/body/div[7]/div[2]/div[2]/div[1]/div[1]/input[1]", fechaInicio);
                Thread.Sleep(3000);

                FindFieldClearSetText("/html/body/div[7]/div[2]/div[2]/div[1]/div[2]/input[1]", fechaFin);
                Thread.Sleep(4000);

                Log.Information("Click de aplicar");
                findFieldClickWait("/html/body/div[7]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr/td/div/table/tbody/tr/td[2]/span", 20);
                Thread.Sleep(5000);

                findFieldClickWait("//*[@id=\"pagecontent\"]/div[2]/div[3]/div[1]/div/div/div[3]/div[1]/table/tbody/tr/td/div/table/tbody/tr/td[2]", 20);
                Log.Information("Click de abrir");
                Thread.Sleep(5000);
                findFieldClickWait("//*[@id=\"reportexport\"]/img", 20);
                Log.Information("Click de reporte exportar");
                Thread.Sleep(5000);
                findFieldClickWait("//*[@id=\"rptDataOverlayPanelContent\"]/div/div[1]/table/tbody/tr[1]/td[2]/a", 20);
                Log.Information("Click selecionar CSV");
                Thread.Sleep(5000);
                findFieldClickWait("//*[@id=\"csvExportBtnContainer\"]/button", 20);
                Log.Information("Click de export");
                Thread.Sleep(5000);
                findFieldClickWait("/html/body/div[2]/div/table/tbody/tr/td[2]/table/tbody/tr/td/a", 20);
                Thread.Sleep(7000);
                Log.Information("Cerrar ventana");

                return true;
            }
            catch(Exception ex)
            {
                Log.Error(ex.ToString());
            }
            return false;

        }
    }
}
