using Serilog;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace RiesgoPichinchaQuoteParser.Models
{
    internal class FileProccessor
    {
        public void ParseFile(string filePathToRead)
        {
            try
            {
                ArrayList fileArrayList = new ArrayList();
                Log.Information($"Procesando archivo: {filePathToRead}");

                //obteniendo nombre de archivo
                string fileName = System.IO.Path.GetFileNameWithoutExtension(filePathToRead);
                Log.Information($"\tConfigurando nuevo nombre de archivo {fileName}");

                //configurando nuevo nombre de archivo
                if (fileName.Contains("PICHINCHA_L"))
                {
                    fileName = fileName.Replace("PICHINCHA_L", "").Trim();
                }
                else if (fileName.Contains("_L"))
                {
                    fileName = fileName.Replace("_L", "").Trim();

                }


                Log.Information($"\tLeyendo contenido de archivo: {filePathToRead}");

                using(StreamReader fileReader = new StreamReader(filePathToRead,System.Text.Encoding.Latin1,false))
                {
                    int counter = 0;
                    string ln;
                    
                    Log.Information($"\tProcesando contenido de archivo: {filePathToRead}");
                    
                    while ((ln = fileReader.ReadLine()) != null)
                    {
                        ln = ln.Replace(@"""", String.Empty).Trim();
                        fileArrayList.Add(ln);
                        counter++;
                    }

                }
                
                WriteFileProccessed(new RiesgoFile(fileName, fileArrayList));
            }
            catch (Exception e)
            {
                Log.Error($"Error en lectura archivo {filePathToRead}: \n\t{0}", e.ToString());
            }


        }

        private void WriteFileProccessed(RiesgoFile fileToWrite)
        {
            try
            {
                Log.Information($"Creando archivo txt {fileToWrite.fullOutPathName}");

                using (TextWriter tw = new StreamWriter(fileToWrite.fullOutPathName,false, System.Text.Encoding.Latin1))
                {

                    for (int i = 0; i < fileToWrite.textContent.Count; i++)
                    {
                        tw.WriteLine(fileToWrite.textContent[i]);
                    }
                }

                Log.Information($"Archivo creado: {fileToWrite.fullOutPathName}\n---***---");
            }
            catch (Exception e)
            {
                Log.Error($"Error en creación de archivo {fileToWrite.fullOutPathName}: \n\t{0}",e.ToString());
            }

        }

    }
}
