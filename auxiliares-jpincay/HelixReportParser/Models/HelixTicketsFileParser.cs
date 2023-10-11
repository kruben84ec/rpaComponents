using HelixTicket = HelixTicketsReportParser.Models.HelixTicket;
using HelixTicketsReportParser.Models;
using Excel = Microsoft.Office.Interop.Excel;
using Serilog;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Data.Common;

namespace HelixTicketsReportParser.Models
{
    internal class HelixTicketsFileParser
    {

        public bool ValidateCell(Excel.Worksheet sheet, int cellRow, int cellCol) {

            // Get the value of the cell at the specified row and column.
            //string cellValue = (string)(sheet.Cells[cellRow, cellCol] as Excel.Range).Value2.ToString();

            string cellValue = Convert.ToString(sheet.Cells[cellRow, cellCol].Value2);

            // Check if the cell value is not null or empty.
            if (string.IsNullOrEmpty(cellValue))
            {
                return false;
            }

            // Create a regular expression pattern that matches the specified formats.
            string pattern = @"^(WO|REQ|RF)\d+$";

            // Use Regex.IsMatch to test if the cell value matches the pattern.
            return Regex.IsMatch(cellValue, pattern);

        }

        public List<HelixTicket> GetHelixTickets(Excel.Worksheet sheet)
        {
            int row = 1;
            List<HelixTicket> helixTickets = new();

            try
            {

                //leer worksheet, crear ticket y agregarlo a la lista a retornar
                for (row = 1; row <= sheet.UsedRange.Rows.Count; row++)
                {
                    HelixTicket ticket = new();

                    ticket.idWo  = ValidateCell(sheet, row, 1) ? Convert.ToString(sheet.Cells[row, 1].Value2) : "";
                    ticket.noReq = ValidateCell(sheet, row, 2) ? Convert.ToString(sheet.Cells[row, 2].Value2) : "";
                    ticket.idOdt = ValidateCell(sheet, row, 3) ? Convert.ToString(sheet.Cells[row, 3].Value2) : "";

                    helixTickets.Add(ticket);

                }


                return helixTickets;

            }
            catch
            (Exception e)
            {
                Log.Error($"GetHelixTickets() Error: No se ha podido obtener la lista de tickets creados\n" +
                    $"Exception: {e}\n{e.Data} sheet row: {row}");
                return null;
            }


        }

        public List<HelixTicket> GetHelixTicketsList(string inputFileFullPath, string inputFileName)
        {

            try
            {

                Excel.Application excel = new()
                {
                    Visible = false
                };

                Excel.Workbook workbook = excel.Workbooks.Open(
                    FileManager.ValidateInputFilePath(inputFileFullPath, inputFileName),
                    Missing.Value,
                    Missing.Value,
                    Excel.XlFileFormat.xlCSV,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value,
                    ",",
                    Missing.Value,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value
                    );
                Excel.Worksheet worksheet = workbook.Worksheets.Item[1];

                List<HelixTicket> helixTickets = GetHelixTickets(worksheet);

                workbook.Close(false);
                excel.Quit();

                return helixTickets;

            }
            catch (Exception e)
            {
                Log.Error($"GetHelixTicketsList({inputFileFullPath},{inputFileName}) Error: No se ha podido obtener la lista de tickets creados\n" +
                        $"Exception: {e}\n");
                return null;
            }

        }

    }
}
