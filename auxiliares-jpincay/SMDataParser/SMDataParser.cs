using AppConfig = SMDataParser.Config.AppConfig;
using Estandar = SMDataParser.Models.Estandar;
using FileManager = SMDataParser.Models.FileManager;
using DataManipulator = SMDataParser.Models.DataManipulator;
using Log = Serilog.Log;
using SMDataParser.Models;

namespace SMDataParser
{
    internal class SMDataParser
    {
        static void Main(string[] args)
        {
            
            AppConfig appConfig = new();
            DataManipulator dataManipulator = new();
            FileManager fileManager = new();
            ProccessHandler proccessHandler = new();

            try
            {                    
               
                appConfig.configureLog();

                List<string> dataList = dataManipulator.GetData(appConfig.inputPath);

                (List<Estandar> dataToWrite, List<string> dataNoRegistrados) = dataManipulator.ParseData(dataList);
                
                fileManager.WriteArchivoBase(dataToWrite, Path.Combine(appConfig.outputPath,"ArchivoBase.xls"));

                fileManager.WriteNoGestionados(dataNoRegistrados, appConfig.odtNoGestionados);

                
                Log.Information($"******* PROCESO TERMINADO CON ÉXITO ******* ");

                
                //BORRA EL ARCHIVO INPUT> REEMPLAZAR POR MOVERLO A CARPETA DE LOG Y RENOMBRARLO CON TIMESTAMP
                //fileManager.DeleteInput(Path.Combine(appConfig.inputPath,appConfig.inputFileName));

                fileManager.BackUpInput(Path.Combine(appConfig.inputPath, appConfig.inputFileName),appConfig.logPath);

                proccessHandler.KillExcelProccess();

            }
            catch(Exception e)
            {
                Log.Error($"SMDataParser Error: {e}");
            }
        }
    }
}