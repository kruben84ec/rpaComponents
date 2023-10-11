using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActualizarReqSM.Models
{
    internal class Validator
    {
        public static string ValidateInputFilePath(String path, String fileName)
        {

            string recentFilePath = "";

            try
            {

                //Lee directorio en busqueda de archivo mas reciente
                var directory = new DirectoryInfo(path);

                Log.Information($"ValidateInputFilePath(): Leyendo directorio {directory} en busca de {fileName}...");

                recentFilePath = (from f in directory.GetFiles() where f.Name == fileName orderby f.LastWriteTime descending select f).First().ToString();

                Log.Information($"ValidateInputFilePath(): Archivo encontrado: {recentFilePath}");

                return recentFilePath;

            }
            catch (Exception e)
            {
                Log.Error($"ValidateInputFilePath({path},{fileName}) Error: Error en la lectura de directorio ({path}) \n" +
                    $"\nError: {e}");
                throw;
            }

        }

    }
}
