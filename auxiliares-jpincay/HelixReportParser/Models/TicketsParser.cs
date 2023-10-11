using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMTicket = HelixTicketsReportParser.Models.SMTicket;
using HelixTicket = HelixTicketsReportParser.Models.HelixTicket;
using Serilog;

namespace HelixTicketsReportParser.Models
{
    internal class TicketsParser
    {
        public List<SMTicket> GetIdPeticion(List<SMTicket> sMTickets, List<HelixTicket> helixTickets)
        {
            List<SMTicket> parsedList = new();

            try
            {

                foreach (SMTicket ticketSM in sMTickets)
                {
                    foreach(HelixTicket ticketH in helixTickets)
                    {
                        if(ticketH.idOdt == ticketSM.idOdt)
                            ticketSM.idPeticion = ticketH.noReq;

                        parsedList.Add(ticketSM);
                    }
                }

                return parsedList;

            }catch(Exception e)
            {
                Log.Error($"GetIdPeticion() Error: No ha sido posible obtener IdPeticionHelix\n" +
                    $"Error: {e}\n");
                return null;
            }
        }
    }
}
