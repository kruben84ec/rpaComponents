using Microsoft.Office.Interop.Excel;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;
using AppConfig = SMDataParser.Config.AppConfig;

namespace SMDataParser.Models
{
    internal class FileManager
    {

        readonly ProccessHandler proccessHandler = new();

        //busca archivo, valida nomeclatura de nombre, devuelve archivo más reciente
        public static string ValidarArchivo(String path)
        {
            AppConfig appConfig = new();
            string recentFileDir = "";

            try
            {

                //Lee directorio en busqueda de archivo mas reciente
                var directory = new DirectoryInfo(path);

                Log.Information($"ValidarArchivo(): Leyendo directorio {directory} en busca de {appConfig.inputPath}...");

                recentFileDir = (from f in directory.GetFiles() where f.Name == appConfig.inputFileName orderby f.LastWriteTime descending select f).First().ToString();

                Log.Information($"ValidarArchivo(): Archivo encontrado: {recentFileDir}");

                return recentFileDir;

            }
            catch (Exception e)
            {
                Log.Error($"ValidarArchivo() Error: Error en la lectura de directorio ({path}) \n" +
                    $"\nError: {e}");
                throw;
            }

        }

        public void WriteNoGestionados(List<string> listaNoGestionados, string outPath)
        {
            try
            {
                //new instance excel app
                Excel.Application xlApp = new()
                {
                    Visible = false,
                    DefaultSaveFormat = XlFileFormat.xlCSV,
                    DisplayAlerts = false
                };


                //new workbook
                Workbook xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                //new worksheet
                Worksheet xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                //escribe cabeceras
                xlWorksheet.Cells[1, 1] = "Número";
                xlWorksheet.Cells[1, 2] = "Descripción";

                //recorrer lista de objetos DataEstandar
                for (int r = 0; r < listaNoGestionados.Count; r++)
                {
                    
                    xlWorksheet.Cells[r + 2 , 1] = listaNoGestionados[r].ToUpper();
                }

                Log.Information($"WriteNoGestionados(): Guardando archivo {outPath}");

                xlWorkbook.SaveAs(outPath, Excel.XlFileFormat.xlCSV,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlWorkbook.Close(true);

                Log.Information($"******* Registros no válidos: {listaNoGestionados.Count}");


            }
            catch (Exception e)
            {
                proccessHandler.KillExcelProccess();
                Log.Error($"WriteNoGestionados(): Error al escribir {outPath}" +
                    $"\nError: {e}");
                throw;
            }


        }

        public void WriteArchivoBase(List<Estandar> dataToWrite, string outPath)
        {

            List<String> cabeceraFinal = new AppConfig().cabeceraFinal;

            try
            {
                //new instance excel app
                Excel.Application xlApp = new()
                {
                    Visible = false,
                    DisplayAlerts = false
                };


                //new workbook
                Workbook xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                //new worksheet
                Worksheet xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                //escribe cabeceras de columnas
                foreach (String cabecera in cabeceraFinal)
                {
                    xlWorksheet.Cells[1, cabeceraFinal.IndexOf(cabecera) + 1] = cabecera.ToUpper();
                    xlWorksheet.Cells[1, cabeceraFinal.Count].EntireRow.Font.Bold = true;
                }

                //recorrer lista de objetos DataEstandar
                for (int r = 0; r < dataToWrite.Count; r++)
                {
                    for (int c = 1; c < cabeceraFinal.Count; c++)
                    {
                        //if (Estandar.ValidateFieldsComplete(dataToWrite[r]))
                        //{
                        var value = dataToWrite[r].GetIndexFieldValue(c - 1);
                        xlWorksheet.Cells[r + 2, c] = value;
                        xlWorksheet.Cells[r + 2, c].NumberFormat = "@";
                    }

                    Log.Information(dataToWrite[r].LogData());

                }

                Log.Information($"WriteArchivoBase(): Guardando archivo {outPath}");

                xlWorkbook.SaveAs(outPath, Excel.XlFileFormat.xlWorkbookNormal,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlWorkbook.Close(true);

                Log.Information($"******* Registros válidos: {dataToWrite.Count}");


            }
            catch (Exception e)
            {
                proccessHandler.KillExcelProccess();
                Log.Error($"WriteArchivoBase(): Error al escribir ArchivoBase.xls" +
                    $"\nError: {e}");
                throw;
            }

        }

        public void BackUpInput(string filePath, string newLocation)
        {
            Log.Information($"BackUpInput(): Respaldando archivo input {filePath}...");
            try
            {
                if (Directory.Exists(newLocation))
                {

                    File.Move(filePath,
                        Path.Combine(
                            newLocation,
                            new string($@"{Path.GetFileNameWithoutExtension(filePath)}_{DateTime.Now:yyyy-M-d_HH}.csv")
                            )
                        );
                }
                else
                {
                    throw new Exception($"BackUpInput(): Ruta de nueva ubicacion de archivo no existe {newLocation}");
                }
            }
            catch (Exception e)
            {
                Log.Error($"BackUpInput(): No se ha podido mover archivo input\n{e}");
            }
        }

        public void DeleteInput(string path)
        {
            Log.Warning($"DeleteInput(): Borrando arhivo input {path}");
            try
            {
                System.IO.File.Delete(path);
            }
            catch (Exception e)
            {
                Log.Error($"DeleteInput() Error: \n{e}");
            }
        }

    }
}
