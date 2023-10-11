using Serilog;
using System.Globalization;

namespace HelixTicketsReportParser.Config
{
    internal class AppConfig
    {
        public string inputPath;
        public string outputPath;
        public string archivoFinalPath;
        public string logPath;
        public string inputFileName;
        public string outputFileName;
        public int colIdPeticionHelix;


        public AppConfig() {

            /*
             * Rutas pruebas 
             *
             */
            //inputPath = @"C:\Users\Jay\Desktop\Diners\6 HelixTicketsReportParser\input\";
            //outputPath = @"C:\Users\Jay\Desktop\Diners\6 HelixTicketsReportParser\output\";
            //archivoFinalPath = @"C:\Users\Jay\Desktop\Diners\6 HelixTicketsReportParser\output\";
            //logPath = @"C:\Users\Jay\Desktop\Diners\6 HelixTicketsReportParser\input\";


            /*
             * Rutas produccion
             * 
             */
            inputPath = @"E:\RECURSOS ROBOT\DATA\MESA_SERVICIO\GESTIONDEUSUARIOS\HELIX\";
            outputPath = @"E:\RECURSOS ROBOT\DATA\MESA_SERVICIO\GESTIONDEUSUARIOS\ARCHBASE\";
            archivoFinalPath = @"E:\RECURSOS ROBOT\DATA\MESA_SERVICIO\GESTIONDEUSUARIOS\ARCHIVOFINAL\";
            logPath = @"E:\RECURSOS ROBOT\LOGS\MESA_SERVICIO\GESTIONDEUSUARIOS\";


            inputFileName = "helixticketsreport.csv";
            outputFileName = "ArchivoBase.xls";
            colIdPeticionHelix = 9;

        }


        public List<string> cabeceraFinal = new() {
            "operacion",
            "ticket",
            "perfil",
            "banco",
            "usuario",
            "identificacion",
            "nombres apellidos",
            "correo",
            "area",
            "numero",
            "estandar"
        };



        public void ConfigLog()
        {
            string logPathFinal = Path.Combine(logPath, new string($@"{DateTime.Now:yyyy-M-d}\"));
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File($"{logPathFinal}{System.AppDomain.CurrentDomain.FriendlyName}_{DateTime.Now:yyyyMMdd-HHmm}.log",
                                 outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Log configurado...");
        }

    }
}
