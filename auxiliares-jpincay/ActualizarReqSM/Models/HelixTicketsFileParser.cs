using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Serilog;
using ProcessHandler = ActualizarReqSM.Models.ProccessHandler;

namespace ActualizarReqSM.Models
{
    internal class HelixTicket
    {

        public string noReq;
        public string idOdt;

        public HelixTicket()
        {
            noReq = string.Empty;
            idOdt = string.Empty;
        } 
    }

    class HelixFileParser
    { 
        public HelixFileParser() { }

        public List<HelixTicket> GetTicketsList(string pathArchivoBase)
        {
            List<HelixTicket> listaTickets = new();

            try
            {
                Excel.Application excel = new()
                {
                    Visible = false
                };

                if (!File.Exists(pathArchivoBase))
                {
                    throw new Exception($"No existe: {pathArchivoBase}");
                }
                
                Excel.Workbook workbook = excel.Workbooks.Open(pathArchivoBase);


                Excel.Worksheet sheet = workbook.Worksheets.Item[1];

                //leer worksheet, crear ticket, obtiene idOdt, agrega a lista a retornar
                for (int row = 2; row <= sheet.UsedRange.Rows.Count; row++)
                {

                    HelixTicket sMTicket = new()
                    {
                        idOdt = (sheet.Cells[row, 1].Value != null) ? Convert.ToString(sheet.Cells[row, 1].Value2) : "",
                        noReq = (sheet.Cells[row, 9].Value != null) ? Convert.ToString(sheet.Cells[row, 9].Value2) : ""

                    };

                    listaTickets.Add(sMTicket);

                }

                workbook.Close(false);
                excel.Quit();

                new ProccessHandler().KillExcelProccess();


                return listaTickets;

            }
            catch (Exception e)
            {
                Log.Error($"GetSMTickets() Error: No se ha podido obtener la lista de tickets creados {pathArchivoBase}\n{e}");
                return listaTickets;
            }
        }

    }
}
