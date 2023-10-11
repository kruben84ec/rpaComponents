using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using StandartValidator;
using StandartValidator.Models;
using System.Text.Json.Nodes;
using System.ComponentModel;
using System.Reflection.Metadata.Ecma335;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using static System.Net.Mime.MediaTypeNames;
using Serilog;

/*
Autor: Jasson Pincay
*/

namespace StandartValidator
{
    class ErrorEnArchivo : Exception
    {
        public ErrorEnArchivo(string message) : base(message) { }
    }

    internal class StandartValidator
    {
        //String inputPath = "C:\\Users\\Jay\\Desktop\\Diners\\3 StandartValidator Test Files\\input\\";
        String inputPath = "E:\\RECURSOS ROBOT\\DATA\\MESA_SERVICIO\\GESTIONDEUSUARIOS\\AUXILIAR\\";

        //String outputPath = "C:\\Users\\Jay\\Desktop\\Diners\\3 StandartValidator Test Files\\output";
        String outputPath = "E:\\RECURSOS ROBOT\\DATA\\MESA_SERVICIO\\GESTIONDEUSUARIOS\\ARCHIVOFINAL\\";

        List<String> cabeceraFinal = new List<string>() { "operacion", "ticket", "perfil", "banco", "usuario", "identificacion", "nombres apellidos", "correo", "area", "numero", "estandar" };


        //quita las tildes de una cadena
        static string NormalizeString(string cadena) => Regex.Replace(cadena.Normalize(NormalizationForm.FormD), @"[^a-zA-z0-9 ]+", "");


        //kill excel process
        private void KillExcelProccess()
        {
            foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }
        }


        //busca archivo, valida nomeclatura de nombre, devuelve archivo más reciente
        private string ValidarArchivo(String path)
        {
            string recentFileDir = "";

            try
            {

                //Lee directorio en busqueda de archivo mas reciente
                var directory = new DirectoryInfo(path);
                return recentFileDir = (from f in directory.GetFiles() where f.Name == "ArchivoAux.xls" orderby f.LastWriteTime descending select f).First().ToString();

            }
            catch (Exception)
            {
                throw;
            }

        }

        public string RegexEmail(string str)
        {
            string RegexPattern = @"\b[A-Z0-9._-]+@[A-Z0-9][A-Z0-9.-]{0,61}[A-Z0-9]\.[A-Z.]{2,6}\b";
            string correo = "";
            Regex rgx = new Regex(RegexPattern);
            Match match = rgx.Match(str);

            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            return correo;
        }


        public string[] ExtractEmails(string str)
        {
            string RegexPattern = @"\b[A-Z0-9._-]+@[A-Z0-9][A-Z0-9.-]{0,61}[A-Z0-9]\.[A-Z.]{2,6}\b";

            // Find matches
            System.Text.RegularExpressions.MatchCollection matches
                = System.Text.RegularExpressions.Regex.Matches(str, RegexPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            string[] MatchList = new string[matches.Count];

            // add each match
            int c = 0;
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                MatchList[c] = match.ToString();
                c++;
            }

            return MatchList;
        }


        private List<string> GetData(string filePath)
        {

            List<String> data = new List<String>();

            try
            {

                Excel.Application xlApp = new Excel.Application();
                xlApp.Visible = false;

                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ValidarArchivo(filePath));
                                
                foreach (Excel._Worksheet sheet in xlWorkbook.Worksheets)
                {
                    foreach (Excel.Range row in sheet.UsedRange.Rows)
                    {

                        string cadena = sheet.Cells[row.Row, 3].Value2.ToString().ToLower();

                        string[] emailExtract = ExtractEmails(cadena);
                        
                        if (emailExtract.Length > 0)
                        {
                            string correo = emailExtract[0];
                            string c1 = NormalizeString(cadena.Substring(0, cadena.IndexOf(correo)));
                            string c2 = cadena.Substring(cadena.IndexOf(correo));

                            cadena = string.Concat(c1, c2);
                            data.Add(cadena 
                                        + " rf " + sheet.Cells[row.Row, 1].Value2.ToString()
                                        + " ticket " + sheet.Cells[row.Row, 2].Value2.ToString());
                        }
                        else
                        {
                            data.Add(NormalizeString(cadena)
                                        + " rf " + sheet.Cells[row.Row, 1].Value2.ToString()
                                        + " ticket " + sheet.Cells[row.Row, 2].Value2.ToString());
                        }

                    }
                }

                KillExcelProccess();

                return data;

            }
            catch (Exception)
            {
                KillExcelProccess();
                throw;

            }
        }


        public List<DataEstandar> ParseData(List<string> dataList)
        {
            string[] standart = { "accion", "identificacion", "perfil a asignar", "usuario", "nombres", "correo" };
            List<DataEstandar> datToWrite = new List<DataEstandar>();

            foreach (String item in dataList)
            {

                int c = 0;

                List<string> values = item.Split().ToList();

                DataEstandar dataEstandar = new DataEstandar();

                //guarda ticket
                dataEstandar.ticket = values.SkipWhile(x => x != "ticket").Skip(1).DefaultIfEmpty(values[0]).FirstOrDefault();

                //guarda numero rf
                dataEstandar.numerorf = values.SkipWhile(x => x != "rf").Skip(1).DefaultIfEmpty(values[0]).FirstOrDefault();


                foreach (string e in standart)
                { 
                    if(item.Contains(e))
                      c++;
                }

                if (c == standart.Length)
                {
                    if (item.Contains("accion"))
                    {
                        string accion = dataEstandar.GetDataBetween(item, "accion", "identificacion");
                        if (accion == "b")
                            dataEstandar.operacion = "borrar";
                        if (accion == "c")
                            dataEstandar.operacion = "crear";
                        if (accion == "a")
                            dataEstandar.operacion = "modificar";
                        c++;
                    }

                    if (item.Contains("perfil a asignar"))
                    {
                        dataEstandar.perfil = dataEstandar.GetDataBetween(item, "perfil a asignar", "usuario");
                        c++;
                    }

                    if (item.Contains("usuario"))
                    {
                        dataEstandar.usuario = dataEstandar.GetDataBetween(item, "usuario", "nombres");
                        c++;
                    }

                    if (item.Contains("identificacion"))
                    {
                        dataEstandar.identificacion = dataEstandar.GetDataBetween(item, "identificacion", "perfil a asignar");
                        c++;
                    }

                    if (item.Contains("nombres"))
                    {
                        dataEstandar.nombres = dataEstandar.GetDataBetween(item, "nombres", "correo");
                        c++;
                    }

                    if (item.Contains("correo"))
                    {
                        dataEstandar.correo = dataEstandar.GetDataBetween(item, "correo", "rf");
                        c++;
                    }

                    if (!dataEstandar.Compare(dataEstandar, item.Substring(0, item.IndexOf(" rf "))))
                    {
                        dataEstandar.estandar = "MANUAL";
                    }
                    else
                    {
                        dataEstandar.estandar = "SI";
                    }
                }
                else
                {
                    dataEstandar.estandar = "MANUAL";
                }
                

                
                datToWrite.Add(dataEstandar);

            }

            return datToWrite;

        }


        public int ContarEstandarSi(List<DataEstandar> lista)
        {
            int c = 0;

            lista.ForEach(e =>
            {
                if (e.estandar == "SI")
                    c++;
            });

            return c;
        }


        public int ContarEstandarNo(List<DataEstandar> lista)
        {
            int c = 0;
            lista.ForEach(e =>
            {
                if (e.estandar != "SI")
                    c++;
            });
            return c;
        }



        public void WriteFile(List<DataEstandar> dataToWrite)
        {

            string folderName = DateTime.Now.ToString("yyyy-M-d");

            string outputPath = this.outputPath + folderName + "\\";

            bool folderOutput = System.IO.Directory.Exists(outputPath);
            if (!folderOutput)
                System.IO.Directory.CreateDirectory(outputPath);

            try
            {
                //new instance excel app
                Excel.Application xlApp = new Excel.Application();
                xlApp.Visible = false;

                Log.Information("Instanciando Excel App: " + xlApp.Path.ToString());

                //new workbook
                Workbook xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                Log.Information("Nuevo archivo excel: " + xlWorkbook.Name.ToString());

                //new worksheet
                Worksheet xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                Log.Information("Nueva hoja de excel: " + xlWorksheet.Name.ToString()); ;

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
                        if (dataToWrite[r].estandar == "SI")
                        {
                            var value = dataToWrite[r].GetIndexFieldValue(c - 1).ToUpper();
                            xlWorksheet.Cells[r + 2, c] = value;
                            xlWorksheet.Cells[r + 2, c].NumberFormat = "@";
                        }
                        else
                        {
                            xlWorksheet.Cells[r + 2, 2] = dataToWrite[r].ticket;
                            xlWorksheet.Cells[r + 2, 2].NumberFormat = "@";

                            xlWorksheet.Cells[r + 2, 10] = dataToWrite[r].numerorf;
                            xlWorksheet.Cells[r + 2, 10].NumberFormat = "@";

                            xlWorksheet.Cells[r + 2, 11] = dataToWrite[r].estandar;
                            xlWorksheet.Cells[r + 2, 11].NumberFormat = "@";
                        }
                    }
                    dataToWrite[r].PrintDataEstandar();
                }

                Log.Information("Guardando archivo: " + outputPath + "ArchivoFinal.xls");

                xlWorkbook.SaveAs(outputPath + "ArchivoFinal.xls", Excel.XlFileFormat.xlWorkbookNormal);
                xlWorkbook.Close(true);
                KillExcelProccess();


            }
            catch (Exception e)
            {
                Log.Error(e.ToString());
                throw;
            }


        }


        static void Main(String[] args) {
            
            Console.Clear();
            
            StandartValidator app = new StandartValidator();
            
            List<string> dataList = app.GetData(app.inputPath);
            List<DataEstandar> dataToWrite = app.ParseData(dataList);

            try
            {
                String logFolderName = "StandartValidatorLog-"+DateTime.Now.ToString("yyyy-M-d");
                
                //variable de ruta de ejecución
                String logPath = "E:\\" + logFolderName;

                System.IO.Directory.CreateDirectory(logPath);

                Log.Logger = new LoggerConfiguration()
                    .WriteTo.File(logPath+"\\"+"StandartValidator_"+".log",
                        rollingInterval: RollingInterval.Hour,
                        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                    .CreateLogger();

                app.WriteFile(dataToWrite);

                Console.WriteLine("\n\n**** Registros con estandar {0} *****\n", app.ContarEstandarSi(dataToWrite));
                Console.WriteLine("**** Registros sin estandar {0} *****\n", app.ContarEstandarNo(dataToWrite));
                Console.WriteLine("***** Registros procesados: {0} *****\n", dataToWrite.Count);

            }
            catch(Exception e)
            {
                Log.Error(e.ToString());
                throw;
            }


        }
    }
}
