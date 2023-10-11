using AppConfig = SMDataParser.Config.AppConfig;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Serilog;
using System.Diagnostics;

namespace SMDataParser.Models
{
    internal class DataManipulator
    {
        //quita las tildes de una cadena
        static string NormalizeString(string cadena) => Regex.Replace(cadena.Normalize(NormalizationForm.FormD), @"[^a-zA-z0-9 ]+", "");


        //obtiene email de una cadena
        private static string ExtractEmail(string str)
        {
            string matchedEmail = "";

            try
            {
                string RegexPattern = @"\b[A-Z0-9._-]+@[A-Z0-9][A-Z0-9.-]{0,61}[A-Z0-9]\.[A-Z.]{2,6}\b";

                // Find matches
                System.Text.RegularExpressions.MatchCollection matches
                    = System.Text.RegularExpressions.Regex.Matches(str, RegexPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                if (matches.Count > 0) { 
                    matchedEmail = matches[0].Value;
                }

            }catch(Exception e)
            {
                Log.Error($"ExtractEmail() Error: {e}\n\t   Cadena recibida: {str}");
            }
            
            return matchedEmail;

        }


        //lee input y obtiene las tramas a procesar y las devuelve como lista
        public List<string> GetData(string filePath)
        {
            ProccessHandler proccessHandler = new();
            FileManager fileManager = new();

            List<String> data = new();

            try
            {

                Excel.Application xlApp = new();
                xlApp.Visible = false;

                //recibe como parametro una ruta al archivo csv
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(FileManager.ValidarArchivo(filePath));

                Log.Information($"GetData(): Leyendo datos de {xlWorkbook.Path} {xlWorkbook.Name}");
                //recorre cada hoja
                foreach (Excel._Worksheet sheet in xlWorkbook.Worksheets)
                {
                    //recorre las filas desde la fila 2
                    for (int i = 2; i <= sheet.UsedRange.Rows.Count; i++)
                    {
                        //Log.Information($"Obteniendo cadena de datos...\n");
                        
                        //obtiene cadena completa
                        string cadena = Convert.ToString(sheet.Cells[i, 1].Value2).ToLower();
                        string rf = cadena.Split(';')[0];
                        cadena = cadena.Split(';')[1];
                        //Log.Information("GetData(): Procesando data: quitando comillas...");
                        //cadena.Replace(@"""", string.Empty);

                        //string rf = Convert.ToString(sheet.Cells[i, 1].Value2).ToLower();
                        //string cadena = Convert.ToString(sheet.Cells[i, 2].Value2).ToLower();

                        string correo = ExtractEmail(cadena);

                        //Log.Information("GetData(): Normalizando cadena: quitando tildes y caracteres especiales...");
                        string c1 = NormalizeString(cadena.Substring(0, cadena.IndexOf(correo)));
                        string c2 = cadena.Substring(cadena.IndexOf(correo));

                        cadena = string.Concat($"{rf};", c1, c2);

                        data.Add(cadena);

                    }
                }
                xlWorkbook.Close(false);
                proccessHandler.KillExcelProccess();

                Log.Information($"GetData(): Datos leídos correctamente (Total: {data.Count})...");

                return data;

            }
            catch (Exception e)
            {
                Log.Error($"GetData(): Error en la lectura de datos" +
                    $"\nError: {e}");
                proccessHandler.KillExcelProccess();
                throw;

            }
        }

        private static string GetDataBetween(string data, string estandar1, string estandar2)
        {

            int ind1 = data.IndexOf(estandar1) + estandar1.Length;
            int ind2 = data.IndexOf(estandar2, ind1);

            string value = data.Substring(ind1, ind2 - ind1).Trim();

            return value;
        }


        private static string StringDataToCompare(Estandar dataToString)
        {

            string operacion = "";

            if (dataToString.operacion == "modificar")
                operacion = "a";
            if (dataToString.operacion == "borrar")
                operacion = "b";
            if (dataToString.operacion == "crear")
                operacion = "c";

            List<string> data = new()
            {
                ("accion " + operacion).Trim(),
                ("identificacion " + dataToString.identificacion).Trim(),
                ("perfil a asignar " + dataToString.perfil).Trim(),
                ("usuario " + dataToString.usuario).Trim(),
                ("nombres " + dataToString.nombres).Trim(),
                ("correo " + dataToString.correo).Trim()
            };

            string dataString = string.Join(" ", data);


            return dataString;
        }


        private static bool Compare(Estandar dataEstandar, string data)
        {
            bool compare = false;
            string compareData = StringDataToCompare(dataEstandar);

            Log.Information($"\tData recibida: {data}");
            Log.Information($"\tData resultado: {compareData}");

            if (compareData == data)
            {
                compare = true;
            }

            return compare;
        }

        private string ParsePefil(string perfil)
        {
            string[] res = perfil.Split(' ');
            if (res.Length >= 2)
            {
                return res[0] + " " + res[1];
            }
            else
            {
                return res[0];
            }

        }

        public (List<Estandar>, List<string>) ParseData(List<string> dataList)
        {

            List<string> standart = new AppConfig().estandardInput;

            List<Estandar> dataArchivoBase = new();
            List<string> dataNoRegistrados = new();

            Log.Information($"Tabulando datos leídos (Total: {dataList.Count})");

            try
            {

                foreach (string item in dataList)
                {

                    int c = 0;

                    Estandar dataEstandar = new()
                    {
                        //guarda numero rf
                        idot = item.Substring(0, item.IndexOf(";")).ToUpper()
                    };

                    string dataItem = item.Split(";")[1];

                    foreach (string e in standart)
                    {
                        if (dataItem.Contains(e))
                            c++;
                    }

                    if (c == standart.Count)
                    {
                        if (dataItem.Contains("accion"))
                        {
                            string accion = GetDataBetween(dataItem, "accion", "identificacion");
                            if (accion == "b")
                                dataEstandar.operacion = "borrar".ToUpper();
                            else if (accion == "c")
                                dataEstandar.operacion = "crear".ToUpper();
                            else if (accion == "a")
                                dataEstandar.operacion = "modificar".ToUpper();
                            else {
                                dataNoRegistrados.Add(item);
                                continue;
                            }
                        }

                        if (dataItem.Contains("perfil a asignar"))
                        {
                            string perfil = GetDataBetween(dataItem, "perfil a asignar", "usuario").ToUpper();
                            
                            dataEstandar.perfil = ParsePefil(perfil);

                            if (!string.IsNullOrEmpty(dataEstandar.perfil))
                            {
                                //valida si solo es letras define opcionSistema Sistema Gestos
                                if(Regex.IsMatch(dataEstandar.perfil, @"^[a-zA-Z]+$"))
                                {
                                    dataEstandar.opcionSistema = "SISTEMA GESTOR";
                                }
                                else
                                {
                                    dataEstandar.opcionSistema = "SISTEMA CAO";
                                }

                            }
                        }

                        if (dataItem.Contains("usuario"))
                        {
                            string nombreUsuario = GetDataBetween(dataItem, "usuario", "nombres").ToUpper();

                            if (nombreUsuario == "")
                            {
                              dataEstandar.usuario = "";
                            }
                            else
                            {
                                dataEstandar.usuario = nombreUsuario;
                            }
                        }

                        if (dataItem.Contains("identificacion"))
                        {
                            dataEstandar.identificacion = GetDataBetween(dataItem, "identificacion", "perfil a asignar").ToUpper();
                        }

                        if (dataItem.Contains("nombres"))
                        {
                            dataEstandar.nombres = GetDataBetween(dataItem, "nombres", "correo").ToUpper();
                        }

                        if (dataItem.Contains("correo"))
                        {
                            dataEstandar.correo = ExtractEmail(dataItem).ToLower();
                        }

                    }

                    if (Estandar.ValidateFieldsComplete(dataEstandar))
                    {
                        dataArchivoBase.Add(dataEstandar);
                    }
                    else
                    {
                        string logData = dataEstandar.idot.ToUpper();
                        Log.Information($"\t{logData} no registrado, campos incompletos...");
                        dataNoRegistrados.Add(item);
                    }

                }
            
                Log.Information($"ParseData(): Datos procesados correctamente (Total {dataList.Count})...");

            }
            catch (Exception e)
            {
                Log.Error($"ParseData(): Error en el procesamiento de los datos + " +
                    $"\nError: {e}");
            }

            return (dataArchivoBase, dataNoRegistrados);

        }

    }

}
