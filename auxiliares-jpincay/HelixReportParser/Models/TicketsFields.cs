using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelixTicketsReportParser.Models
{
    internal class HelixTicket
    {

        public string idWo;
        public string noReq;
        public string idOdt;

        public HelixTicket() { 
            idWo = string.Empty;
            noReq = string.Empty;
            idOdt = string.Empty;
        }

    }

    internal class SMTicket
    {
        public string idOdt;
        public string operacion;
        public string nombres;
        public string identificacion;
        public string correo;
        public string perfil;
        public string opcionsistema;
        public string usuario;
        public string idPeticion;
        public string banco;
        public string area;
        public string estandar;

        public SMTicket() {
            idOdt = string.Empty;
            operacion = string.Empty;
            nombres = string.Empty;
            identificacion = string.Empty;
            correo = string.Empty;
            perfil = string.Empty;
            opcionsistema = string.Empty;
            usuario = string.Empty;
            idPeticion = string.Empty;
            banco = "PICHINCHA";
            area = "CT";
            estandar = string.Empty;
    }
    }

    internal class DataEstandar
    {

        public string operacion;
        public string ticket;
        public string perfil;
        public string banco;
        public string usuario;
        public string identificacion;
        public string nombres;
        public string correo;
        public string area;
        public string numerorf;
        public string estandar;

        public DataEstandar()
        {

            operacion = string.Empty;
            ticket = string.Empty;
            perfil = string.Empty;
            banco = "PICHINCHA";
            usuario = string.Empty;
            identificacion = string.Empty;
            nombres = string.Empty;
            correo = string.Empty;
            area = "CT";
            numerorf = string.Empty;
            estandar = string.Empty;

        }
    }

}
