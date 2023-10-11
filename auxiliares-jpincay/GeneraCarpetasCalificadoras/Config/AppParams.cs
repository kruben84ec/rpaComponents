using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneraCarpetasCalificadoras.Config
{
    class AppParams
    {
        public string rutaArchivoConfig;
        public string baseFolderPath;
        public string rutaLog;

        public AppParams() {
            rutaArchivoConfig = string.Empty;
            baseFolderPath = string.Empty;
            rutaLog = Path.Combine(System.AppContext.BaseDirectory.ToString(),new string((@$"{DateTime.Now:yyyy-MM-dd}\")));
        }

        public AppParams(string rutaBaseCarpetas, string archivoConfigRuta, string logRuta) { 
            rutaArchivoConfig = archivoConfigRuta;
            baseFolderPath = rutaBaseCarpetas;
            rutaLog = logRuta;
        }
    }

    class LogConfigurator {

        public static void ConfigLog(string logPath="")
        {
            string logToPath = (string.IsNullOrEmpty(logPath)) ? new AppParams().rutaLog : logPath;

            if(!Directory.Exists(logToPath)) 
            { 
                Directory.CreateDirectory(logToPath);
            }

            string logFileName = $"{System.AppDomain.CurrentDomain.FriendlyName}_{DateTime.Now:yyyyMMdd-HHmm}.log";

            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File($"{Path.Combine(logToPath,logFileName)}",
                                 outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Log configurado...");
        }

    }
}
