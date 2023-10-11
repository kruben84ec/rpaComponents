using Model = GeneraCarpetasCalificadoras.Models.Model;
using GeneraCarpetasCalificadoras.Config;
using Serilog;
using LogConfigurator = GeneraCarpetasCalificadoras.Config.LogConfigurator;
using ProccessHandler = GeneraCarpetasCalificadoras.Models.ProccessHandler;

namespace GeneraCarpetasCalificadoras
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                if(args.Length == 3)
                {
                    AppParams appParams = new AppParams()
                    {
                        rutaArchivoConfig = args[0],
                        baseFolderPath = args[1],
                        rutaLog = args[2]
                    };

                    Model model = new Model();

                    LogConfigurator.ConfigLog(appParams.rutaLog);

                    List<string> folerNamesList = model.GetFolderNames(appParams.rutaArchivoConfig);

                    if(folerNamesList.Count == 0)
                    {
                        throw new Exception($"No se pudo obtener lista de nombres de carpetas desde {appParams.rutaArchivoConfig}");
                    }

                    model.CrearEstructura(folerNamesList, appParams.baseFolderPath);

                }
                else
                {
                    LogConfigurator.ConfigLog();
                    throw new Exception($"No se han recibido los parámetros esperados...\n{ string.Join(",",args) }");
                }

                ProccessHandler.KillExcelProccess();

            }
            catch (Exception e)
            {
                Log.Error($"{AppDomain.CurrentDomain.FriendlyName} App Error {e}");
            }
        }
    }
}