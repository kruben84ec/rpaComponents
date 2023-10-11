using Microsoft.Office.Interop.Excel;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ProcesaBaseGeneratica.Models
{
    internal class GestionaReporte
    {
        public static void ComprimirArchivo(string rutaArchivo)
        {
            string nombreArchivo = Path.GetFileNameWithoutExtension(rutaArchivo);
            string rutaArchivoZip = Path.Combine(Path.GetDirectoryName(rutaArchivo), nombreArchivo + ".zip");

            using (FileStream archivoZip = new FileStream(rutaArchivoZip, FileMode.Create))
            {
                using (ZipArchive zip = new ZipArchive(archivoZip, ZipArchiveMode.Create, false))
                {
                    string nombreEntrada = Path.GetFileName(rutaArchivo);
                    zip.CreateEntryFromFile(rutaArchivo, nombreEntrada);
                }
            }

            Console.WriteLine("Archivo comprimido correctamente: " + rutaArchivoZip);
        }

        public static void ActualizarReporte(List<Cliente> clientes, string rutaArchivoReporte)
        {
            try
            {
                string cedulaValue = "";
                string cellPrincipalValue = "";
                string cellTipoBusqValue = "";

                Excel.Application excelApp = new Excel.Application()
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                Excel.Workbook wbook = excelApp.Workbooks.Open(rutaArchivoReporte);
                Excel.Worksheet sheet = wbook.Worksheets[1];

                Log.Information($"Atualizando {rutaArchivoReporte}");

                foreach(Cliente cliente in clientes)
                {
                    for(int row = 2; row <=  sheet.UsedRange.Rows.Count; row++)
                    {

                        Excel.Range cellCedula = sheet.Cells[row, 3];
                        cedulaValue = cellCedula.Value != null ? Convert.ToString(cellCedula.Value2) : "";

                        if(cedulaValue == cliente.Identificacion)
                        {
                            Excel.Range cellCasoID = sheet.Cells[row, 1];
                            Excel.Range cellMitigacionActual = sheet.Cells[row, 5];
                            Excel.Range cellMitigacionNueva = sheet.Cells[row, 6];
                            Excel.Range cellFechaVencimiento = sheet.Cells[row, 7];
                            Excel.Range cellObservacion = sheet.Cells[row, 8];
                            Excel.Range cellDelitos = sheet.Cells[row, 9];
                            Excel.Range cellComentario = sheet.Cells[row, 10];

                            cellCasoID.NumberFormat = "@";
                            cellMitigacionActual.NumberFormat = "@";
                            cellMitigacionNueva.NumberFormat = "@";
                            cellFechaVencimiento.NumberFormat = "@";
                            cellObservacion.NumberFormat = "@";
                            cellDelitos.NumberFormat = "@";
                            cellComentario.NumberFormat = "@";

                            cellCasoID.Value = (cliente.CasoID == "") ? "FICHA NO EXISTE" : cliente.CasoID;
                            cellMitigacionActual.Value = cliente.MitigacionActual;
                            cellMitigacionNueva.Value = cliente.MitigacionNueva;
                            cellFechaVencimiento.Value = cliente.FechaVencimiento;
                            cellObservacion.Value = cliente.Observacion;

                            if( cliente.DelitosCedulaCJ.Count == 0 && cliente.DelitosNombreCJ.Count == 0 && cliente.DelitosCedulaFG.Count == 0 && cliente.DelitosNombreFG.Count == 0)
                            {
                                cellDelitos.Value = "";
                            }
                            else
                            {
                                cellDelitos.Value = $"Delitos encontrados: " +
                                    $"FISCALIA: {new string(string.Join(", ", cliente.DelitosCedulaFG))} {new string(string.Join(", ", cliente.DelitosNombreFG))}" +
                                    $"C JUDICATURA: {new string(string.Join(", ", cliente.DelitosCedulaCJ))} {new string(string.Join(", ", cliente.DelitosNombreCJ))}";
                            }
                            cellComentario.Value = cliente.Comentario;


                        }

                    }

                }

                wbook.Save();
                wbook.Close(true);
                excelApp.Quit();

                Log.Information($"Reporte atualizado: {rutaArchivoReporte}");

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.Collect();


            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message}\n{e.StackTrace}");
            }
        }
    }
}
