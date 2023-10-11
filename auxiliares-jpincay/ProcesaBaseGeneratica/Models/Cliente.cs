using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcesaBaseGeneratica.Models
{
    public class Cliente
    {

        public string       CasoID = "";
        public string       Identificacion = "";
        public string       IdentificacionConyugue = "";
        public string       Nombre = "";
        public bool         Conyugue = false;
        public string       MitigacionActual = "";
        public string       FechaMitigacionActual = "";
        public string       MitigacionNueva = "";
        public string       FechaVencimiento = "";
        public string       Observacion = "";
        public List<string> DelitosCedulaCJ = new ();
        public List<string> DelitosNombreCJ = new ();
        public List<string> DelitosCedulaFG = new ();
        public List<string> DelitosNombreFG = new ();
        public string       Comentario = "";
        
        public Cliente() { }

        public Cliente(bool Conyugue,
                        string CasoID = "",
                        string Identificacion = "",
                        string IdentificacionConyugue = "",
                        string Nombre = "",
                        string MitigacionActual = "",
                        string FechaMitigacionActual = "",
                        string MitigacionNueva = "",
                        string FechaVencimiento = "",
                        string Observacion = "",
                        List<string> DelitosCedulaCJ = null,
                        List<string> DelitosNombreCJ = null,
                        List<string> DelitosCedulaFG = null,
                        List<string> DelitosNombreFG = null,
                        string Comentario = "")
        {

            this.CasoID = CasoID;
            this.Identificacion = Identificacion;
            this.IdentificacionConyugue = IdentificacionConyugue;
            this.Nombre = Nombre;
            this.Conyugue = Conyugue;
            this.MitigacionActual= MitigacionActual;
            this.FechaMitigacionActual = FechaMitigacionActual;
            this.MitigacionNueva = MitigacionNueva;
            this.FechaVencimiento = FechaVencimiento;
            this.Observacion = Observacion;
            this.DelitosCedulaCJ = DelitosCedulaCJ;
            this.DelitosNombreCJ = DelitosNombreCJ;
            this.DelitosCedulaFG = DelitosCedulaFG;
            this.DelitosNombreCJ = DelitosNombreFG;
            this.Comentario = Comentario;

        }
    }
}
