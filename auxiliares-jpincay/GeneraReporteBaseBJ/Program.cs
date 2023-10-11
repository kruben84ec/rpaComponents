using Serilog;

namespace GeneraReporteBaseBJ
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {

                string sourceFile = "";
                string outputFile = "";

                if (args.Length == 0)
                {
                    sourceFile = @"E:\RECURSOS ROBOT\DATA\BUSQUEDAJUICIOS\ARCHIVOS\BaseGeneratica\Resultado Generatica para historial BJ.xlsx";
                    outputFile = @"E:\RECURSOS ROBOT\DATA\BUSQUEDAJUICIOS\ARCHIVOS\ReporteFinal\REPORTE-RPA-BJ.xlsx";
                }
                else
                {
                    sourceFile = args[0];
                    outputFile = args[1];
                }



                string logPath = @"E:\RECURSOS ROBOT\LOGS\BUSQUEDAJUICIOS\";

                ConfigLog(logPath);

                if (!File.Exists(sourceFile))
                {
                    throw new FileNotFoundException(sourceFile);
                }

                ExcelDuplicateFilter.FilterDuplicates(sourceFile, outputFile);

                GC.Collect();

            }
            catch (Exception ex)
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