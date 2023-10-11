using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Cliente = ProcesaBaseGeneratica.Models.Cliente;
using System.Text.RegularExpressions;
using Serilog;
using System.Diagnostics;

namespace ProcesaBaseGeneratica.Models
{

    public class ValidaMonitoreoBusquedaJ
    {


        public static void AnalizarClientes(List<Cliente> Clientes, string rutaArchivoConsolidado)
        {

            try
            {

                Excel.Application excelApp = new();

                Excel.Workbook wb = excelApp.Workbooks.Open(rutaArchivoConsolidado);
                Excel.Worksheet ws = wb.Worksheets[1];

                foreach (Cliente cliente in Clientes)
                {

                    for (int row = 2; row <= ws.UsedRange.Rows.Count; row++)
                    {
                        Excel.Range cellId = ws.Cells[row, 2];
                        string cellIdValue = cellId.Value != null ? Convert.ToString(cellId.Value2) : "";

                        if (cellIdValue == cliente.Identificacion)
                        {

                            Excel.Range cellCasoID = ws.Cells[row, 1];
                            string cellCasoIDValue = cellCasoID != null ? Convert.ToString(cellCasoID.Value2) : "";
                            cliente.CasoID = cellCasoIDValue;

                            Excel.Range cellMitigacion = ws.Cells[row, 14];
                            string cellMitigacionValue = cellMitigacion != null ? Utils.NormalizeString(Convert.ToString(cellMitigacion.Value2)).ToLower() : "";

                            if (cellMitigacionValue == "busqueda juicio")
                            {
                                cliente.MitigacionActual = cellMitigacionValue.ToUpper();

                                Excel.Range cellFechaMitigacion = ws.Cells[row, 11];
                                string cellFechaMitigacionValue = cellFechaMitigacion != null ? Convert.ToString(cellFechaMitigacion.Value2) : "";

                                cliente.FechaMitigacionActual = cellFechaMitigacionValue;

                                bool res = Utils.VerificarDiferenciaAniosMayorDos(cellFechaMitigacionValue);

                                if (res)
                                {
                                    cliente.MitigacionNueva = "MONITOREO";
                                }
                                else
                                {
                                    cliente.MitigacionNueva = cliente.MitigacionActual.ToUpper();
                                }

                            }
                            else if (cellMitigacionValue == "justificado")
                            {
                                cliente.MitigacionActual = cellMitigacionValue.ToUpper();
                                cliente.MitigacionNueva = cellMitigacionValue.ToUpper();
                            }
                            else { 
                                if (cliente.MitigacionActual == "")
                                {
                                    continue;
                                }
                                else
                                {
                                    break;
                                }
                            }

                        }

                    }

                    if(cliente.MitigacionActual == "")
                    {
                        cliente.MitigacionActual = "FICHA NO EXISTE";
                        cliente.MitigacionNueva = "FICHA NO EXISTE";
                        cliente.Observacion = "FICHA NO EXISTE";
                        
                    }

                }


                wb.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.Collect();
            }
            catch (Exception e)
            {
                Log.Error($"{e.Message}");
            }
        }


        public static List<Cliente> ObtenerClientes(string rutaArchivoReporte)
        {

            try
            {
                List<Cliente> clientes = new List<Cliente>();


                Excel.Application excelApp = new()
                {
                    Visible = false
                };

                Excel.Workbook wb = excelApp.Workbooks.Open(rutaArchivoReporte);
                Excel.Worksheet ws = wb.Worksheets[1];

                for (int row = 2; row <= ws.UsedRange.Rows.Count; row++)
                {
                    //leer columnas identificacion y nombre

                    Excel.Range CellId = ws.Cells[row, 3];
                    Excel.Range CellNombre = ws.Cells[row, 4];

                    string valueId = CellId.Value != null ? Convert.ToString(CellId.Value2.ToLower()) : "";
                    string valueNombre = CellId.Value != null ? Utils.NormalizeString(Convert.ToString(CellNombre.Value2)).ToUpper() : "";

                    Cliente cliente = new Cliente()
                    {
                        Identificacion = valueId.Trim(),
                        Nombre = valueNombre.Trim()
                    };

                    clientes.Add(cliente);

                }

                wb.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.Collect();

                return clientes;

            }
            catch (Exception e)
            {
                Log.Error($"{e.Message}\n{e.StackTrace}");
                return null;
            }


        }

    }
}
