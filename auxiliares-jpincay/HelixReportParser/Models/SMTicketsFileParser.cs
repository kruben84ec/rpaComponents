using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;
using SMTicket = HelixTicketsReportParser.Models.SMTicket;
using DataEstandar = HelixTicketsReportParser.Models.DataEstandar;
using Microsoft.Office.Interop.Excel;
using System.Net.Sockets;
using HelixTicketsReportParser.Config;
using System.Globalization;
using System.Text.RegularExpressions;

namespace HelixTicketsReportParser.Models
{
    internal class SMTicketsFileParser
    {
        public List<SMTicket> GetSMTickets(string smXlFileFullPath, string smXlFileName)
        {
            List<SMTicket> smTickets = new();

            try
            {
                Excel.Application excel = new()
                {
                    Visible = false
                };
                Excel.Workbook workbook = excel.Workbooks.Open(
                    FileManager.ValidateInputFilePath(smXlFileFullPath, smXlFileName));

                Excel.Worksheet sheet = workbook.Worksheets.Item[1];

                //leer worksheet, crear ticket, obtiene idOdt, agrega a lista a retornar
                for (int row = 2; row <= sheet.UsedRange.Rows.Count; row++)
                {

                    SMTicket sMTicket = new()
                    {
                        idOdt           = (sheet.Cells[row, 1].Value != null) ? Convert.ToString(sheet.Cells[row, 1].Value2) : "",
                        operacion       = (sheet.Cells[row, 2].Value != null) ? Convert.ToString(sheet.Cells[row, 2].Value2) : "",
                        nombres         = (sheet.Cells[row, 3].Value != null) ? Convert.ToString(sheet.Cells[row, 3].Value2) : "",
                        identificacion  = (sheet.Cells[row, 4].Value != null) ? Convert.ToString(sheet.Cells[row, 4].Value2) : "",
                        correo          = (sheet.Cells[row, 5].Value != null) ? Convert.ToString(sheet.Cells[row, 5].Value2) : "",
                        perfil          = (sheet.Cells[row, 6].Value != null) ? Convert.ToString(sheet.Cells[row, 6].Value2) : "",
                        opcionsistema   = (sheet.Cells[row, 7].Value != null) ? Convert.ToString(sheet.Cells[row, 7].Value2) : "",
                        usuario         = (sheet.Cells[row, 8].Value != null) ? Convert.ToString(sheet.Cells[row, 8].Value2) : ""
                    };

                    smTickets.Add(sMTicket);

                }

                workbook.Close(false);

                return smTickets;

            }
            catch (Exception e)
            {
                Log.Error($"GetSMTickets({smXlFileFullPath},{smXlFileName}) Error: No se ha podido obtener la lista de tickets creados\n" +
                    $"Exception: {e}\n{e.Data}");
                return null;
            }
        }

        public void UpdateFile(string smXlFileFullPath, string smXlFileName, List<SMTicket> smTicketsParsed, int colIdPeticionHelix)
        {
            try
            {
                Excel.Application excel = new()
                {
                    Visible = false
                };
                Excel.Workbook workbook = excel.Workbooks.Open(
                    FileManager.ValidateInputFilePath(smXlFileFullPath, smXlFileName));

                Excel.Worksheet sheet = workbook.Worksheets.Item[1];

                Log.Information($"Actualizando {smXlFileName}...");

                //leer worksheet, compara  idOdts y registra id peticion
                for (int row = 2; row <= sheet.UsedRange.Rows.Count; row++)
                {
                    foreach(SMTicket ticket in smTicketsParsed)
                    {
                        if (sheet.Cells[row,1].Value2 == ticket.idOdt)
                        {
                            sheet.Cells[row, colIdPeticionHelix] = ticket.idPeticion;
                        }

                    }
                }
                workbook.Save();
                workbook.Close();
                excel.Quit();

            }
            catch (Exception e)
            {
                Log.Error($"UpdateFile() Error: No se ha podido actualizar {smXlFileFullPath}\\{smXlFileName}\n" +
                    $"Exception: {e}\n");

            }
        }

        public static bool ValidateFieldsComplete(SMTicket ticketData)
        {
            bool val = false;
            int c = 0;

            Type type = ticketData.GetType();

            foreach (var f in type.GetFields().Where(f => f.IsPublic))
            {
                if (f.GetValue(ticketData).ToString() == "" && f.Name != "estandar")
                    c++;
            }

            val = c == 0;

            return val;
        }

        static string GetDigitsWithoutLeadingZeros(string input)
        {
            //string pattern = "(REQ.)(0*)(\\d+)";
            //string replacement = "$3";
            //return Regex.Replace(input, pattern, replacement);
            return Regex.Match(input, "(REQ.)(\\d+)").Groups[2].Value.TrimStart('0');
        }

        public bool GenerateArchivoFinal(List<SMTicket> dataList, string outputPath)
        {
            Log.Information($"Generando ArchivoFinal.xlsx...");
            try
            {

                if (!Directory.Exists(outputPath))
                {
                    Directory.CreateDirectory(outputPath);
                    Log.Information($"Ruta configurada {outputPath}");
                }


                Excel.Application xlApp = new()
                {
                    Visible = false,
                    DisplayAlerts = false
                };
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet sheet = xlWorkbook.Worksheets.Item[1];
                sheet.Name = "ArchivoFinal";

                List<string> cabeceras = new AppConfig().cabeceraFinal;

                //escribe cabeceras
                foreach (string cabecera in cabeceras)
                {
                    sheet.Cells[1, cabeceras.IndexOf(cabecera) + 1] = cabecera.ToUpper();
                    sheet.Cells[1, cabeceras.Count].EntireRow.Font.Bold = true;

                }

                //escribe data
                for (int r = 0; r < dataList.Count; r++)
                {

                    sheet.Cells[r + 2, 1] = dataList[r].operacion.ToUpper();
                    
                    sheet.Cells[r + 2, 2] = $"'{GetDigitsWithoutLeadingZeros(dataList[r].idPeticion).ToUpper()}";
                    sheet.Cells[r + 2, 2].NumberFormat = "@";
                    
                    sheet.Cells[r + 2, 3] = dataList[r].perfil.ToUpper();
                    sheet.Cells[r + 2, 4] = dataList[r].banco.ToUpper();
                    sheet.Cells[r + 2, 5] = dataList[r].usuario.ToUpper();
                    
                    sheet.Cells[r + 2, 6] = $"'{dataList[r].identificacion.ToUpper()}";
                    sheet.Cells[r + 2, 6].NumberFormat = "@";

                    sheet.Cells[r + 2, 7] = dataList[r].nombres.ToUpper();
                    sheet.Cells[r + 2, 8] = dataList[r].correo.ToUpper();
                    sheet.Cells[r + 2, 9] = dataList[r].area.ToUpper();
                    sheet.Cells[r + 2, 10] = dataList[r].idOdt.ToUpper();

                    dataList[r].estandar = ValidateFieldsComplete(dataList[r]) ? "SI" : "MANUAL";
                    sheet.Cells[r + 2, 11] = dataList[r].estandar.ToUpper();

                    Log.Information($"Escribiendo registro {r} IDODT: {dataList[r].idOdt.ToUpper()}");

                }

                Log.Information($"GenerateArchivoFinal(): Guardando archivo ArchivoFinal.xlsx en : {outputPath}");

                xlWorkbook.SaveAs(outputPath + "ArchivoFinal.xlsx",
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlWorkbook.Close(true);
                xlApp.Quit();

                return File.Exists(Path.Combine(outputPath, "ArchivoFinal.xlsx"));



            }
            catch (Exception e)
            {
                Log.Error($@"GenerateArchivoFinal() Error: No se pudo generar {outputPath}\ArchivoFinal.xls" +
                    $"Exception: {e}\n");
                return false;
            }
        }

    }
}
