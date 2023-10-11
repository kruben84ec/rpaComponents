using Excel = Microsoft.Office.Interop.Excel;
using Serilog;


namespace ProcesaBaseGeneratica.Models
{
    class ProcesarBaseGeneratica
    {

        public void AnalizarBaseGeneratica(string rutaBaseGeneratica, List<Cliente> clientes)
        {

            try
            {

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(rutaBaseGeneratica);
                Excel.Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Count];

                foreach (Cliente cliente in clientes)
                {
                    AnalizaConsejoJudicatura(cliente, sheet);
                    AnalizaFiscalia(cliente, sheet);


                    Log.Information($"\n<Caso_{cliente.Identificacion}>\n" +
                                    $"\tCasoID: {cliente.CasoID}\n" +
                                    $"\tNombre: {cliente.Nombre} \n" +
                                    $"\tMitigacionNueva: {cliente.MitigacionNueva} \n" +
                                    $"</Caso_{cliente.Identificacion}>\n\n");

                }

                workbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.Collect();

            }
            catch (Exception e)
            {
                Log.Error($"{e.Message}\n{e.StackTrace}");
            }
        }





        private void AnalizaFiscalia(Cliente cliente, Excel.Worksheet sheet)
        {

            try
            {
                string rutaBasePalabras = @"E:\RECURSOS ROBOT\DATA\BUSQUEDAJUICIOS\ARCHIVOS\Palabras clave Juicios.xlsx";

                List<string> palabrasClaveJuicios = BaseDelitos.ObtenerListaPalabras(rutaBasePalabras);

                string cedulaValue = "";
                string cellPrincipalValue = "";
                string cellTipoBusqValue = "";

                for (int row = 4; row <= sheet.UsedRange.Rows.Count; row++)
                {
                    Excel.Range cellCedula = sheet.Cells[row, 3];
                    cedulaValue = cellCedula.Value != null ? Convert.ToString(cellCedula.Value2) : "";

                    Excel.Range cellPrincipal = sheet.Cells[row, 2];
                    cellPrincipalValue = cellCedula.Value != null ? Convert.ToString(cellPrincipal.Value2) : "";

                    Excel.Range cellTipoBusq = sheet.Cells[row, 10];
                    cellTipoBusqValue = cellTipoBusq.Value != null ? Utils.NormalizeString(Convert.ToString(cellTipoBusq.Value2)).ToLower() : "";

                    if (cedulaValue.Contains(cliente.Identificacion) && cellTipoBusqValue != "")
                    {
                        //identificando si es conyugue
                        if (cellPrincipalValue != cedulaValue)
                        {
                            cliente.Conyugue = true;
                            cliente.IdentificacionConyugue = cellPrincipalValue;

                        }

                        Excel.Range cellNoticiaDelito = sheet.Cells[row, 12];
                        string cellNoticiaDelitoValue = cellNoticiaDelito.Value != null ? Convert.ToString(cellNoticiaDelito.Value2).ToLower() : "";

                        //validando casos sin coincidencias/registros
                        if (cellNoticiaDelitoValue == "sin coincidencias encontradas" || cellNoticiaDelitoValue == "sin registros encontrados")
                        {
                            continue;
                        }
                        else
                        {

                            //validando delitos
                            Excel.Range cellDelito = sheet.Cells[row, 17];
                            string cellDelitoValue = cellDelito.Value != null ? Convert.ToString(cellDelito.Value2).ToLower() : "";

                            Excel.Range cellEstado = sheet.Cells[row, 15];
                            string cellEstadoValue = cellEstado.Value != null ? Utils.NormalizeString(Convert.ToString(cellEstado.Value2)).ToLower() : "";

                            Excel.Range cellFecha = sheet.Cells[row, 14];
                            string cellFechaDelitoValue = cellFecha.Value != null ? Convert.ToString(cellFecha.Value2).ToLower() : "";

                            foreach (string palabra in palabrasClaveJuicios)
                            {
                                if (cellDelitoValue.Contains(palabra))
                                {
                                    bool mayor2anios = Utils.VerificarDiferenciaAniosMayorDos(cellFechaDelitoValue);
                                    if (!mayor2anios || cellEstadoValue.Contains("sospechoso") ||
                                        cellEstadoValue.Contains("sospechoso no reconocido") ||
                                        cellEstadoValue.Contains("procesado")
                                        )
                                    {
                                        if (cellTipoBusqValue == "cedula")
                                        {
                                            cliente.DelitosCedulaFG.Add($"{palabra.ToUpper()} - {cellFechaDelitoValue}");

                                        }
                                        if (cellTipoBusqValue == "nombres")
                                        {
                                            cliente.DelitosNombreFG.Add($"{palabra.ToUpper()} - {cellFechaDelitoValue}");
                                        }

                                    }
                                }
                            }


                        }

                    }

                }
                if (cliente.DelitosCedulaCJ.Count == 0 && cliente.DelitosNombreCJ.Count == 0)
                {
                    if (cliente.DelitosCedulaFG.Count > 0 || cliente.DelitosNombreFG.Count > 0)
                    {
                        cliente.MitigacionNueva = "BUSQUEDA JUICIO";
                    }

                }

                if (cliente.DelitosCedulaFG.Count < 4 || cliente.DelitosNombreFG.Count < 4)
                {
                    cliente.Comentario += $"{new string((cliente.Conyugue == true) ? $"CNYG/ACC: {cliente.Identificacion} (Principal: {cliente.IdentificacionConyugue})- " : "")}" +
                                          $"\"BJ\" FISCALIA GENERAL: " +
                                          $"Resultado POR CEDULA: {new string((cliente.DelitosCedulaFG.Count == 0) ? "SIN COINCIDENCIAS ENCONTRADAS" : string.Join(", ", cliente.DelitosCedulaFG))}. " +
                                          $"Resultado POR NOMBRE: {new string((cliente.DelitosNombreFG.Count == 0) ? "SIN COINCIDENCIAS ENCONTRADAS" : string.Join(", ", cliente.DelitosNombreFG))}. ";

                }
                else
                {
                    cliente.Observacion = "MANUAL (VARIOS DELITOS)";

                }

                cliente.FechaVencimiento = (cliente.MitigacionNueva == "MONITOREO") ? Utils.SumarDias(5) : Utils.SumarDias(90);


            }
            catch (Exception e)
            {
                Log.Error($"{e.Message}\n{e.StackTrace}");
            }

        }





        private void AnalizaConsejoJudicatura(Cliente cliente, Excel.Worksheet sheet)
        {

            try
            {
                string rutaBasePalabras = @"E:\RECURSOS ROBOT\DATA\BUSQUEDAJUICIOS\ARCHIVOS\Palabras clave Juicios.xlsx";

                List<string> palabrasClaveJuicios = BaseDelitos.ObtenerListaPalabras(rutaBasePalabras);

                string cedulaValue = "";
                string cellPrincipalValue = "";
                string cellTipoBusqValue = "";

                for (int row = 4; row <= sheet.UsedRange.Rows.Count; row++)
                {
                    Excel.Range cellCedula = sheet.Cells[row, 3];
                    cedulaValue = cellCedula.Value != null ? Convert.ToString(cellCedula.Value2) : "";

                    Excel.Range cellPrincipal = sheet.Cells[row, 2];
                    cellPrincipalValue = cellCedula.Value != null ? Convert.ToString(cellPrincipal.Value2) : "";

                    Excel.Range cellTipoBusq = sheet.Cells[row, 5];
                    cellTipoBusqValue = cellTipoBusq.Value != null ? Utils.NormalizeString(Convert.ToString(cellTipoBusq.Value2)).ToLower() : "";

                    if (cedulaValue.Contains(cliente.Identificacion) && cellTipoBusqValue != "")
                    {

                        //identificando si es conyugue
                        if (cellPrincipalValue != cedulaValue)
                        {
                            cliente.Conyugue = true;
                            cliente.IdentificacionConyugue = cellPrincipalValue;
                        }

                        Excel.Range cellNumProceso = sheet.Cells[row, 8];
                        string procValue = cellNumProceso.Value != null ? Convert.ToString(cellNumProceso.Value2).ToLower() : "";

                        if (procValue == "sin coincidencias encontradas")
                        {
                            continue;
                        }
                        else
                        {
                            Excel.Range cellAccion = sheet.Cells[row, 9];
                            string accionValue = cellNumProceso.Value != null ? Convert.ToString(cellAccion.Value2).ToLower() : "";

                            Excel.Range cellFechaDelito = sheet.Cells[row, 7];
                            string fechaDelito = cellFechaDelito.Value != null ? Convert.ToString(cellFechaDelito.Value2).ToLower() : "";

                            foreach (string palabra in palabrasClaveJuicios)
                            {
                                if (accionValue.Contains(palabra))
                                {
                                    if (cellTipoBusqValue == "cedula")
                                    {
                                        cliente.DelitosCedulaCJ.Add($"{palabra.ToUpper()} - {fechaDelito}");
                                    }
                                    if (cellTipoBusqValue == "nombres")
                                    {
                                        cliente.DelitosNombreCJ.Add($"{palabra.ToUpper()} - {fechaDelito}");
                                    }
                                }
                            }

                        }

                    }

                }

                if (cliente.DelitosCedulaCJ.Count > 0 || cliente.DelitosNombreCJ.Count > 0)
                {
                    cliente.MitigacionNueva = "MONITOREO";
                }


                if (cliente.DelitosCedulaCJ.Count < 4 || cliente.DelitosNombreCJ.Count < 4)
                {
                    cliente.Comentario += $"{new string((cliente.Conyugue == true) ? $"CNYG/ACC: {cliente.Identificacion} (Principal: {cliente.IdentificacionConyugue})- " : "")}" +
                                          $"\"BJ\" CONSEJO JUDICATURA: " +
                                          $"Resultado POR CEDULA: {new string((cliente.DelitosCedulaCJ.Count == 0) ? "SIN COINCIDENCIAS ENCONTRADAS" : string.Join(", ", cliente.DelitosCedulaCJ))}. " +
                                          $"Resultado POR NOMBRE: {new string((cliente.DelitosNombreCJ.Count == 0) ? "SIN COINCIDENCIAS ENCONTRADAS" : string.Join(", ", cliente.DelitosNombreCJ))}. ";

                }
                else
                {
                    cliente.Observacion = "MANUAL (VARIOS DELITOS)";

                }

                cliente.FechaVencimiento = (cliente.MitigacionNueva == "MONITOREO") ? Utils.SumarDias(5) : Utils.SumarDias(90);


            }
            catch (Exception e)
            {
                Log.Error($"{e.Message}\n{e.StackTrace}");
            }


        }
    }
}
