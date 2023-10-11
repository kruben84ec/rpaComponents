using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RiesgoPichinchaQuoteParser.Config
{
    internal class AppConfig
    {
        public AppConfig() { }

        public string inputPath = "E:\\NUEVO RIESGO PICHINCHA\\Exportaciones.ILB\\";
        //public string inputPath = "C:\\Users\\Jay\\Desktop\\Diners\\5 NuevoRiesgoPichincha\\Exportaciones.ILB\\";

        public string outputPath = "E:\\NUEVO RIESGO PICHINCHA\\Archivos fuente.ILB\\";
        //public string outputPath = "C:\\Users\\Jay\\Desktop\\Diners\\5 NuevoRiesgoPichincha\\Archivos fuente.ILB\\";
        
        public string logPath = "E:\\RECURSOS ROBOT\\LOGS\\NUEVORIESGO_CTLINT\\";
        //public string logPath = "C:\\Users\\Jay\\Desktop\\Diners\\5 NuevoRiesgoPichincha\\Archivos fuente.ILB\\";

    }
}
