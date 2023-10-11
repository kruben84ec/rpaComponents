using Excel = Microsoft.Office.Interop.Excel;
using Serilog;

namespace ValidarDatosCT

{
    internal class ValidarDatosCT
    {

        private static string logPath = Path.Combine(@"E:\RECURSOS ROBOT\LOGS\MESA_SERVICIO\GESTIONDEUSUARIOS\", new string($@"{DateTime.Now:yyyy-M-d}\"));
        private static string rutaArchivFinal = Path.Combine(@"E:\RECURSOS ROBOT\DATA\MESA_SERVICIO\GESTIONDEUSUARIOS\ARCHIVOFINAL\", new string($@"{DateTime.Now:yyyy-M-d}\ArchivoFinal.xlsx"));
        private static string archivoCT = Path.Combine(@"E:\RECURSOS ROBOT\DATA\MESA_SERVICIO\GESTIONDEUSUARIOS\ARCHIVOFINAL\", new string($@"{DateTime.Now:yyyy-M-d}\CT-{DateTime.Now:dMyyyy}.xlsx"));
        

        static void Main(string[] args)
        {
            try
            {
                ConfigureLog();

                Log.Information($"Generando archivo CT");

                if (File.Exists(rutaArchivFinal))
                {
                    File.Copy(rutaArchivFinal, archivoCT,true);
                    
                    if (File.Exists(archivoCT))
                    {
                        Log.Information($"Archivo CT generado en {archivoCT}");
                    }

                    ValidarDatos();

                }
                else
                {
                    Log.Error($"ArchivoFinal.xlsx no encontrado en {rutaArchivFinal}");
                }


            }
            catch (Exception ex)
            {
                Log.Error($"{ex.Message}\n{ex.StackTrace}");
            }
        }


        private static void ValidarDatos()
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                string fecha = DateTime.Now.ToString("dMyyyy");
                string fecha2 = DateTime.Now.ToString("yyyy-M-d");
                string fechaTotal = @"E:\RECURSOS ROBOT\DATA\MESA_SERVICIO\GESTIONDEUSUARIOS\ARCHIVOFINAL\" + fecha2 + @"\CT-" + fecha + ".xlsx";

                Excel.Workbook workbook = excelApp.Workbooks.Open(fechaTotal);
                Excel.Worksheet tickets = workbook.Sheets["ArchivoFinal"];

                tickets.Copy(After: excelApp.Sheets[excelApp.Sheets.Count]);

                int incidentes = tickets.Cells[tickets.Rows.Count, "J"].End(Excel.XlDirection.xlUp).Row;

                for (int indexIncidente = incidentes; indexIncidente >= 2; indexIncidente--)
                {
                    string valorA = tickets.Cells[indexIncidente, "A"].Value?.ToString();
                    string valorB = tickets.Cells[indexIncidente, "B"].Value?.ToString();
                    string valorH = tickets.Cells[indexIncidente, "H"].Value?.ToString();
                    string valorK = tickets.Cells[indexIncidente, "K"].Value?.ToString();

                    if (valorA == "CREAR")
                        tickets.Cells[indexIncidente, "E"] = "";

                    tickets.Cells[indexIncidente, "B"].Value = valorB?.Trim();
                    tickets.Cells[indexIncidente, "H"].Value = valorH?.Replace("' ", "");

                    if (valorK == "MANUAL")
                    {
                        tickets.Rows[indexIncidente].Delete();
                    }
                }

                for (int indexIncidente = 2; indexIncidente <= 100; indexIncidente++)
                {
                    string valorB = tickets.Cells[indexIncidente, "B"].Value?.ToString();

                    if (valorB == "0")
                        tickets.Cells[indexIncidente, "B"].Value = null;
                }

                tickets.Name = "Usuarios";
                workbook.Sheets["ArchivoFinal (2)"].Name = "BACKARCHIVO";
                tickets.Select();

                workbook.Save();
                workbook.Close();

                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            }
            catch (Exception ex)
            {
                throw new Exception($"{ex.Message}\n{ex.StackTrace}");
            }

        }

        private static void ConfigureLog()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File($"{logPath}{System.AppDomain.CurrentDomain.FriendlyName}_{DateTime.Now:yyyyMMdd-HHmm}.log",
                                 outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Log configurado...");
        }


    }
}