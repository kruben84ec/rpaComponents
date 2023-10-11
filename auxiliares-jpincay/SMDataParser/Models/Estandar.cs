
using Serilog;
using System;
using System.Reflection;
using System.Text;

namespace SMDataParser.Models
{
    public class Estandar
    {

        public string idot;
        public string operacion;
        public string nombres;
        public string identificacion;
        public string correo;
        public string perfil;
        public string opcionSistema;
        public string usuario;
        

        public Estandar(){

            this.idot = "";
            this.operacion = "";
            this.nombres = "";
            this.identificacion = "";
            this.correo = "";
            this.perfil = "";
            this.opcionSistema = "";
            this.usuario = "";
        }

        public string GetIndexFieldValue(int indice)
        {
            List<string> data = new()
            {
                this.idot,
                this.operacion,
                this.nombres,
                this.identificacion,
                this.correo,
                this.perfil,
                this.opcionSistema,
                this.usuario
            };

            return data[indice];

        }

        public void PrintDataEstandar()
        {
            List<string> data = new()
            {
                this.idot,
                this.operacion,
                this.nombres,
                this.identificacion,
                this.correo,
                this.perfil,
                this.opcionSistema,
                this.usuario
            };

            Log.Information($"Escribiendo: {string.Format($"{string.Join(" ", data)}")}");
        }

        public string LogData()
        {
            string logData = $"Escribiendo {this.idot} {this.operacion} ";

            return logData;
        }

        public static bool ValidateFieldsComplete(Estandar dataList)
        {
            bool val = false;
            int c = 0;
            
            Type type = dataList.GetType();
            
            foreach (var f in type.GetFields().Where(f => f.IsPublic))
            {
                
                if ( f.GetValue(dataList).ToString() == "" )
                {
                    c++;
                }
            }

            if (c > 0 && dataList.usuario == "" && dataList.operacion == "CREAR")
            {
                c=0;
            }

            if(c == 0)
                val = true;

            return val;
        }

    }

}
