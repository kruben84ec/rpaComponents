using Serilog;
using MergeFiles = MergeHistoriales.Models.MergeFiles;
namespace MergeHistoriales
{
    internal class MergeHistoriales
    {
        static void Main(string[] args)
        {
            string rutaIncidencias = @"E:\RECURSOS ROBOT\DATA\BUSQUEDAJUICIOS\ARCHIVOS\CasosIncidenciaHistorial"; // Ruta del directorio donde se encuentran los archivos Excel

            string archivoConsolidado = Path.Combine(rutaIncidencias, "ConsolidadoHistoriales.xlsx"); // Ruta del archivo de salida

            string logPath = @"E:\RECURSOS ROBOT\LOGS\BUSQUEDAJUICIOS\";

            new AppLog(logPath);

            try
            {
                MergeFiles.MergeHistorialIncidencias(rutaIncidencias, archivoConsolidado);

            }
            catch (Exception e)
            {
                Log.Error($"{e.Message}\n{e.StackTrace}");
            }

        }

        class AppLog
        {
            public AppLog(string logPath)
            {

                ConfigLog(logPath);
            }

            private void ConfigLog(string logPath)
            {
                string logPathFinal = Path.Combine(logPath, new string($@"{DateTime.Now:yyyyMMdd}\"));
                Log.Logger = new LoggerConfiguration()
                    .WriteTo.Console()
                    .WriteTo.File($"{logPathFinal}LogTech_BUSQUEDAJUICIOS_{DateTime.Now:yyyyMMdd}.xml",
                                     outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                    .CreateLogger();

                Log.Information("Log configurado...");
            }


        }

    }
}