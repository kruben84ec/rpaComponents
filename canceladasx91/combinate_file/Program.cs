using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Globalization;

namespace combinate_file
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                Stopwatch timeMeasure = new Stopwatch();
                //Capturar el tiempo de inicio
                String hourMinute = DateTime.Now.ToString("HH:mm");
                timeMeasure.Start();
                String pathReport = @args[0];
                String pathReportMacro = pathReport+ @"\MacroCX91.xlsm";
                String path = @"E:\Canceladasx91\config\marcas.xlsx";
                String pathReportExecute = pathReport + @"\reporte_";

                //Obtener los códigos de las marcas de los archivos txt
                ServiceExcel wb = new ServiceExcel(path, 1);
                List<string> codeBrandList = wb.getCodeBrand();
                //Obtener la información de los reportes generadas por los iconos de transferencia
                ServicesTxt txtBoletinadas = new ServicesTxt();
                List<string> dataExcelBoletindas = new List<string>();
                List<string> dataExcelYobsidiam = new List<string>();
                txtBoletinadas.getDataToExcel(codeBrandList, pathReport);
                dataExcelBoletindas = txtBoletinadas.accountBoletinadas;
                dataExcelYobsidiam = txtBoletinadas.accountYobsidiam;

                //Guadar la información en el archivo de excel
                ServiceExcel wbMacro = new ServiceExcel(pathReportMacro, 2);


                //Las cuentas boletinadas tinen una cabcera por tanto debe empezar en 2
                int initWriteBoletinadas = (wbMacro.initRowWrite(pathReportMacro,2)+1);

                wbMacro.insertDataExcel(pathReportMacro, dataExcelBoletindas,2, initWriteBoletinadas);
                int initWriteYobsidiam = wbMacro.initRowWrite(pathReportMacro, 1);


                wbMacro.insertDataExcel(pathReportMacro, dataExcelYobsidiam, 1, initWriteYobsidiam);
                
                //Ejecutar la macro
                ServiceExcel wbMacroConsolidado = new ServiceExcel(pathReportMacro, 2);

                wbMacroConsolidado.executeMacro(pathReportMacro, "Macro5");
                


                //Generar el reporte de ejecución
                report.ReportController reporteEjecucion = new report.ReportController();
                reporteEjecucion.genereteReport(pathReportMacro, hourMinute, pathReportExecute);
                


                //Fin del tiempode proceso
                timeMeasure.Stop();



            }

        }
    }
}
