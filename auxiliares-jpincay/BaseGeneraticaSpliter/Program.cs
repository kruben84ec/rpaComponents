using Serilog;

namespace BaseGeneraticaSpliter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string rutaBaseGeneratica = @"E:\RECURSOS ROBOT\DATA\BUSQUEDAJUICIOS\ARCHIVOS\BaseGeneratica\Resultado Generatica para historial BJ.xlsx";

                string rutaBasePorCorte = @"E:\RECURSOS ROBOT\DATA\BUSQUEDAJUICIOS\ARCHIVOS\BasesPorCorte";

                string logPath = @"E:\RECURSOS ROBOT\LOGS\BUSQUEDAJUICIOS\";

                ConfigLog(logPath);

                ExcelSplitter.SplitExcelFile(rutaBaseGeneratica, 10, rutaBasePorCorte);

                GC.Collect();

            }catch (Exception ex)
            {
                Log.Error($"Error: {ex.Message}\n{ex.StackTrace}");
                GC.Collect();
            }

        }

        private static void ConfigLog(string logPath)
        {
            string logPathFinal = Path.Combine(logPath, new string($@"{DateTime.Now:yyyyMMdd}\LogTech_BUSQUEDAJUICIOS_{DateTime.Now:yyyyMMdd}.xml"));
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File($"{logPathFinal}",
                                 outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Log configurado...");

        }


    }
}