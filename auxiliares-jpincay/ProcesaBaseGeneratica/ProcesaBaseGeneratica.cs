using Serilog;
using ValidaMonitoreoBusquedaj = ProcesaBaseGeneratica.Models.ValidaMonitoreoBusquedaJ;
using ProccessHandler = ProcesaBaseGeneratica.Models.ProccessHandler;
using ProcesarBaseGeneratica = ProcesaBaseGeneratica.Models.ProcesarBaseGeneratica;
using ProcesaBaseGeneratica.Models;

namespace ProcesaBaseGeneratica
{

    public class ProcesaBaseGeneratica
    {
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

        static void Main(string[] args)
        {
            //2023-07-17
            string rutaIncidencias = @"E:\RECURSOS ROBOT\DATA\BUSQUEDAJUICIOS\ARCHIVOS\CasosIncidenciaHistorial"; // Ruta del directorio donde se encuentran los archivos Excel

            string archivoConsolidado = Path.Combine(rutaIncidencias, "ConsolidadoHistoriales.xlsx"); // Ruta del archivo de salida
            
            string rutaBaseGeneratica = @"E:\RECURSOS ROBOT\DATA\BUSQUEDAJUICIOS\ARCHIVOS\BaseGeneratica\Resultado Generatica para historial BJ.xlsx";


            string rutaArchivoReporte = @"E:\RECURSOS ROBOT\DATA\BUSQUEDAJUICIOS\ARCHIVOS\ReporteFinal\REPORTE-RPA-BJ.xlsx";

            string logPath = @"E:\RECURSOS ROBOT\LOGS\BUSQUEDAJUICIOS\";

            ConfigLog(logPath);

            try
            {
                
                //abrir reporte y obtener lista de clientes a validar                
                if (!File.Exists(rutaArchivoReporte))
                {
                    throw new Exception($"No se pudo obtener lista de clientes para analizar: Reporte base no existe {rutaArchivoReporte}");
                }
                List<Cliente> clientes = ValidaMonitoreoBusquedaJ.ObtenerClientes(rutaArchivoReporte);


                if(!File.Exists(archivoConsolidado) )
                {
                    throw new Exception($"No se pudo analizar clientes: Consolidado Historiales no existe {archivoConsolidado}");
                }
                ValidaMonitoreoBusquedaj.AnalizarClientes(clientes, archivoConsolidado);

                if (!File.Exists(rutaBaseGeneratica))
                {
                    throw new Exception($"No se pudo analizar clientes: Archivo base generatica no existe {rutaBaseGeneratica}");
                }
                new ProcesarBaseGeneratica().AnalizarBaseGeneratica(rutaBaseGeneratica, clientes);

                //actualizar reporte
                GestionaReporte.ActualizarReporte(clientes, rutaArchivoReporte);

                //respaldar archivo en carpeta log

                string[] rutasArchivosResapaldo = { archivoConsolidado, rutaBaseGeneratica, rutaArchivoReporte };

                string nombreArchivoRespaldo = new string($"ARCHIVOS-RPA-BUSQUEDAJUICIOS-{new string($@"{DateTime.Now:yyyyMMdd}_{DateTime.Now:HH}")}.zip");
                string rutaRespaldo = Path.Combine(logPath, new string($@"{DateTime.Now:yyyyMMdd}\{nombreArchivoRespaldo}"));

                Log.Information($"Respaldando archivos en {rutaRespaldo}");

                Utils.ComprimirArchivos(rutaRespaldo, rutasArchivosResapaldo);

                GC.Collect();
            }
            catch (Exception ex)
            {
                Log.Error($"{ex.Message}\n{ex.StackTrace}");
                GC.Collect();
            }

        }
    }

}