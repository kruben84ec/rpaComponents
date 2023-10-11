using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Threading;
using System.Collections.Generic;
using System.Diagnostics;
using Serilog;
using HelixNavigate;


namespace helixIntegration
{
    internal class Program
    {
        public void ConfigLog()
        {
            string path = @"E:\RECURSOS ROBOT\LOGS\MESA_SERVICIO\GESTIONDEUSUARIOS\\";
            string fecha_log = $@"{DateTime.Now:yyyy-M-d}\";
            string logPathFinal = Path.Combine(path, fecha_log);
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File($"{logPathFinal}{System.AppDomain.CurrentDomain.FriendlyName}_{DateTime.Now:yyyyMMdd-HHmm}.log",
                                 outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Log configurado...");
        }

        static void Main(string[] args)
        {
            string message_tikets = "";


            Program execute = new Program();

            tikets tikets = new tikets();

 
           

            Process[] processes = Process.GetProcessesByName("cmd");
            if (args.Length > 0)
            {
                List<tikets> tickets = new List<tikets>();

                Login login = new Login();
                IWebDriver web = execute.initWeb();

                ReporteHelix reporte = new ReporteHelix(web);
                HelperRpa help = new HelperRpa(web);
                help.ConfigLog(@"E:\RECURSOS ROBOT\LOGS\MESA_SERVICIO\GESTIONDEUSUARIOS\\");

                try
                {
                    String pathReport = @args[0];

                    tickets = execute.getTickets(web, pathReport);

                    Ticket ticketHelix = new Ticket(web);
                    GestionUsuarios servicePageTicket = new GestionUsuarios(web);

                    String catalogoUrl = "https://dceservice-dwp.onbmc.com/dwp/app/#/catalog";
                    web.Navigate().GoToUrl(catalogoUrl);
                    Thread.Sleep(3000);
                    bool isAccess = login.access(web, "https://or-rsso1.onbmc.com/rsso/start");
                    Thread.Sleep(5000);

                    int numeroTicktes = tickets.Count;
                    message_tikets = "Ticktes a procesar: " + numeroTicktes.ToString();

                    Log.Information($"{message_tikets}");


                    
                    foreach (var ticket in tickets)
                    {

                        Thread.Sleep(10000);
                        message_tikets = "Procesar ticket: " + ticket.idOdt.ToString();

                        Log.Information($"{message_tikets}");

                        servicePageTicket.gestionUsuario(web, isAccess);
                        ticketHelix.crear(ticket);
                        ticketHelix.modificar(ticket);
                        ticketHelix.eliminar(ticket);
                        Thread.Sleep(5000);
                        Log.Information("siguiente ticket");

                        web.Navigate().GoToUrl(catalogoUrl);

                        message_tikets = "Fin del ticket: " + ticket.idOdt.ToString();
                        Log.Information($"{message_tikets}");


                    }
                    
                    Thread.Sleep(3000);
                    Log.Information("Descargando reporte de tickets gestionados");

                    reporte.accessReport(web);
                                       

                }
                catch (Exception e)
                {
                    Log.Error($"{e.Message}");
                }
                Thread.Sleep(2000);

  
                web.Close();
                web.Quit();
                

                //Fin del tiempo de proceso

            }
            else
            {
                message_tikets = "No ingreso la ruta del archivo base";
                Log.Error($"{message_tikets}");
            }
        }
        
        internal List<tikets> getTickets(IWebDriver driver, string pathFile)
        {
            Ticket ticketHelix = new Ticket(driver);
            List<tikets> tickets = new List<tikets>();
            string pathFileExcel = pathFile;

            Log.Information($"Ruta del archivo fuente de tickets: {pathFileExcel}");

            tickets = ticketHelix.GetSMTickets(pathFileExcel);
            return tickets;
        }

        internal IWebDriver initWeb()
        {
            ChromeOptions options = new ChromeOptions();
            //Set the argument 
            options.AddArguments("--start-maximized");
            options.AddArguments("--ignore-certificate-errors");
            options.AddArguments("--ignore-ssl-errors");
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            options.AddExcludedArgument("enable-automation");
            options.AddAdditionalChromeOption("useAutomationExtension", false);
            //Set Chrome to work with headless mode
            IWebDriver web = new ChromeDriver(options);
            return web;
        }
    }
}
