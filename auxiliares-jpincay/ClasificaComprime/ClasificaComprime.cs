using Serilog;
using System.IO.Compression;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ClasificaComprime
{
    internal class ClasificaComprime
    {
        private static readonly string logPath = @"E:\RECURSOS ROBOT\LOGS\CALIFICADORAS\";
        private static readonly string rutaLocalAreas = @"E:\RECURSOS ROBOT\DATA\CALIFICADORAS\ARCHIVOS\";
        private static readonly string rutaLocalFTP = @"E:\RECURSOS ROBOT\DATA\CALIFICADORAS\FTP\";
        private static readonly string rutaArchivoConfig = @"E:\RECURSOS ROBOT\DATA\CALIFICADORAS\CONFIG\RPAConfig.xlsx";

        private static readonly string zipBWR = Path.Combine(rutaLocalAreas, "bwr.zip");
        private static readonly string zipCIR = Path.Combine(rutaLocalAreas, "cir.zip");

        static void Main(string[] args)
        {
            

            try
            {
                ConfigLog(logPath);

                ClasificaArchivos();

                ComprimirCarpetas(rutaLocalFTP);

                string[] carpetasAreas = Directory.GetDirectories(rutaLocalAreas);

                // Iterar sobre cada carpeta
                foreach (string carpeta in carpetasAreas)
                { 
                    Directory.Delete(carpeta, true);
                }


                }
            catch (Exception e)
            {
                Log.Error($"{e.Message}\n{e.StackTrace}");
            }


            GC.Collect();

        }

        private static void ComprimirCarpetas(string directoryPath)
        {
            // Get the list of folders within the directory
            string[] folders = Directory.GetDirectories(directoryPath);

            foreach (string folder in folders)
            {
                // Create a zip file name based on the folder name
                string zipFileName = folder + $" {DateTime.Now:MMMM-yyyy}" + ".zip";

                if (File.Exists(Path.Combine(Path.GetFullPath(folder), zipFileName)))
                {
                    File.Delete(Path.Combine(Path.GetFullPath(folder), zipFileName));
                }

                try
                {
                    // Create a new zip file
                    ZipFile.CreateFromDirectory(folder, zipFileName);

                    // Delete the source folder
                    Directory.Delete(folder, true);

                    Log.Information($"Archivo zip'{zipFileName}' creado");
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error durante la creacion del archivo zip de la carpeta '{folder}': {ex.Message}\n{ex.StackTrace}");
                }
            }
        }

        private static void ClasificaArchivos()
        {
            try {
                List<string> archivosIdBWR = ObtenerIDArchivosCalificadora("BWR", rutaArchivoConfig);
                List<string> archivosIdCIR = ObtenerIDArchivosCalificadora("CIR", rutaArchivoConfig);

                string rutaFinalCIR = Path.Combine(rutaLocalFTP, "CIR");
                string rutaFinalBWR = Path.Combine(rutaLocalFTP, "BWR");

                Directory.CreateDirectory(rutaFinalBWR);
                Directory.CreateDirectory(rutaFinalCIR);

                // Obtener la lista de carpetas dentro de la carpeta principal
                string[] carpetasAreas = Directory.GetDirectories(rutaLocalAreas);

                // Iterar sobre cada carpeta
                foreach (string carpeta in carpetasAreas)
                {

                    DirectoryInfo dir = new DirectoryInfo(carpeta);
                    //Console.WriteLine(dir.Name);

                    string rutaArchivosAreaCalBWR = Path.Combine(rutaFinalBWR, dir.Name);
                    string rutaArchivosAreaCalCIR = Path.Combine(rutaFinalCIR, dir.Name);

                    if (dir.Name != "BWR" && dir.Name != "CIR")
                    {
                        Directory.CreateDirectory(rutaArchivosAreaCalBWR);
                        Directory.CreateDirectory(rutaArchivosAreaCalCIR);
                    }

                    //ubicar archivos
                    string[] archivos = Directory.GetFiles(carpeta);

                    Log.Information($"Clasificando archivos calificadora carpeta {dir.Name}");

                    foreach (string archivo in archivos)
                    {
                        string nombreArchivo = Path.GetFileName(archivo);

                        foreach (string id in archivosIdCIR)
                        {
                            // Escape the variable value for regex pattern
                            string escapedVariable = Regex.Escape(id);

                            // Create the regex pattern dynamically
                            string pattern = @"-(" + escapedVariable + @")-";

                            bool isMatch = Regex.IsMatch(nombreArchivo, pattern);


                            //if (nombreArchivo.Contains(id))
                            if (isMatch)
                            {
                                if (!File.Exists(Path.Combine(rutaArchivosAreaCalCIR, nombreArchivo)))
                                {
                                    File.Copy(archivo, Path.Combine(rutaArchivosAreaCalCIR, nombreArchivo));
                                }
                            }
                        }

                        foreach (string id in archivosIdBWR)
                        {
                            // Escape the variable value for regex pattern
                            string escapedVariable = Regex.Escape(id);

                            // Create the regex pattern dynamically
                            string pattern = @"\[?" + escapedVariable + @"\]?";

                            bool isMatch = Regex.IsMatch(nombreArchivo, pattern);

                            //if (nombreArchivo.Contains(id))
                            if (isMatch)
                            {
                                if (!File.Exists(Path.Combine(rutaArchivosAreaCalBWR, nombreArchivo)))
                                {
                                    File.Copy(archivo, Path.Combine(rutaArchivosAreaCalBWR, nombreArchivo));
                                }
                            }
                        }

                    }
                    // fin ubicar archivos

                }

            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message}\n{e.StackTrace}");
            }
        }

        private static List<string> ObtenerIDArchivosCalificadora(string nombreCalificadora, string rutaConfigFile) { 
        
            var lista = new List<string>();

            try
            {

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook wb = excelApp.Workbooks.Open(rutaConfigFile);
                Excel.Worksheet ws = wb.Worksheets[3];

                for(int row = 2; row <= ws.UsedRange.Rows.Count; row++)
                {
                    Excel.Range cellCalificadora = ws.Cells[row, 5];
                    string valueCalificadora = cellCalificadora != null ? Convert.ToString(cellCalificadora.Value2) : "";

                    Excel.Range cellIDarchivos = ws.Cells[row, 3];
                    string valueIDarchivo = cellIDarchivos != null ? Convert.ToString(cellIDarchivos.Value2) : "";

                    if (valueCalificadora == nombreCalificadora)
                    {
                        lista.Add(valueIDarchivo);
                    }
                }

                wb.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.Collect();



            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message}\n{e.StackTrace}");
            }

            return lista;
        }

        private static void ConfigLog(string logPath)
        {

            string logPathFinal = Path.Combine(logPath, new string($@"{DateTime.Now:yyyyMMdd}\LogTech_CALIFICADORAS_{DateTime.Now:yyyyMMdd}.xml"));
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File($"{logPathFinal}",
                                 outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Log configurado...");

        }


    }
}