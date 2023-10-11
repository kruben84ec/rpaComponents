using RiesgoPichinchaQuoteParser.Config;
using RiesgoPichinchaQuoteParser.Models;
using Log = Serilog.Log;


namespace RiesgoPichinchaQuoteParser
{

    class RiesgoPichinchaFilesParser
    {

        static void Main(string[] args)
        {
            AppConfig appConfig = new AppConfig();

            FileProccessor fileProccessor = new FileProccessor();
            LoggerConfigurator logger = new LoggerConfigurator();

            try
            {

                logger.configureLog(appConfig.logPath);

                Log.Information($"Leyendo directorio: {appConfig.inputPath}");

                string[] filesPath = Directory.GetFiles(appConfig.inputPath, "*.DEL", SearchOption.AllDirectories);

                Log.Information($"{filesPath.LongLength.ToString()} archivos encontrados... ");

                for(int i=0; i < filesPath.Length;i++)
                {
                    fileProccessor.ParseFile(filesPath[i]);
                }
                Log.Information($"Proceso terminado con éxito: {filesPath.LongLength.ToString()} archivos procesados....");
            }
            catch (Exception e)
            {
                Log.Error($"Error en lectura de directorio {appConfig.inputPath}: \n\t{0}", e.ToString());
            }
        }
    }
}