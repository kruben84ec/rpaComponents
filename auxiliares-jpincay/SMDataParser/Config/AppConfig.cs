using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMDataParser.Config
{
    internal class AppConfig
    {
        public string inputPath;
        public string outputPath;
        public string odtNoGestionados;
        public string logPath;
        public string inputFileName;

        public AppConfig() {

            /*
             * Rutas pruebas
             */
            //inputPath = @"C:\Users\Jay\Desktop\Diners\4 TicketParser ServiceManagerHelix\input\";
            //outputPath = @"C:\Users\Jay\Desktop\Diners\4 TicketParser ServiceManagerHelix\output\";
            //logPath = Path.Combine(@"C:\Users\Jay\Desktop\Diners\4 TicketParser ServiceManagerHelix\input\",
            //              new string($@"{DateTime.Now:yyyy-M-d}\"));
            //odtNoGestionados = Path.Combine(logPath, new string($@"ODTNoGestionados_{DateTime.Now:yyyy-M-d_HH}.csv"));

            /*
             * Rutas produccion
            */

            inputPath = @"E:\RECURSOS ROBOT\DATA\MESA_SERVICIO\GESTIONDEUSUARIOS\ARCHBASE\";
            outputPath = @"E:\RECURSOS ROBOT\DATA\MESA_SERVICIO\GESTIONDEUSUARIOS\ARCHBASE\";
            logPath = Path.Combine(@"E:\RECURSOS ROBOT\LOGS\MESA_SERVICIO\GESTIONDEUSUARIOS\", new string($@"{DateTime.Now:yyyy-M-d}\"));
            odtNoGestionados = Path.Combine(logPath, new string($@"ODTNoGestionados_{DateTime.Now:yyyy-M-d_HH}.csv"));

            inputFileName = "export.csv";

        }


        public List<String> estandardInput = new() {
            "accion",
            "identificacion",
            "perfil a asignar",
            "usuario",
            "nombres",
            "correo"
        };

        public List<String> cabeceraFinal = new() { 
            "idodt", 
            "operacion", 
            "nombres apellidos", 
            "identificacion", 
            "correo", 
            "perfil",
            "opcionSistema",
            "usuario", 
            "idpeticionhelix" 
        };


        public void configureLog()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File($"{logPath}{System.AppDomain.CurrentDomain.FriendlyName}_{DateTime.Now:yyyyMMdd-HHmm}.log",
                                 outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Log configurado...");
        }

    }
}
