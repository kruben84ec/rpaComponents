using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using static System.Net.Mime.MediaTypeNames;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

public class ErrorEnArchivo : Exception {
    public ErrorEnArchivo(string message) : base(message){
        
    }
}
class Program
{
    private void killExcelProccess()
    {
        foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
        {
            proc.Kill();
        }
    }

    //busca archivo, valida nomeclatura de nombre, devuelve archivo más reciente
    private string GetInputDir(String path)
    {
        string recentFileDir = "";

        try
        {

            //Lee directorio en busqueda de archivo mas reciente
            var directory = new DirectoryInfo(path);
            return recentFileDir = (from f in directory.GetFiles() orderby f.LastWriteTime descending select f).First().ToString();

        }
        catch (Exception)
        {
            throw;
        }

    }



    private void ReadInput(String validFilePath)
    {
        Excel.Application xlApp = new Excel.Application();
        Excel.Application xlApp2 = new Excel.Application();

        //string archivoAuxPath = "C:\\Users\\Jay\\Desktop\\Diners\\macro-noticket\\AUXFILE\\";
        string archivoAuxPath = "E:\\RECURSOS ROBOT\\DATA\\MESA_SERVICIO\\GESTIONDEUSUARIOS\\AUXILIAR\\";

        string fileAuxOut = archivoAuxPath + "ArchivoAux.xls";

        Console.WriteLine("Procesando archivo {0}\n", validFilePath);

        try
        {
            xlApp.Visible = false;
            xlApp2.Visible = false;

            Excel.Workbook xlInputFile = xlApp.Workbooks.Open(validFilePath);
            Excel.Workbook xlAuxFileOut = xlApp2.Workbooks.Open(fileAuxOut);

            Excel._Worksheet xlInWorksheet = xlInputFile.Sheets[1];
            Excel._Worksheet xlAuxOutWorksheet = xlAuxFileOut.Sheets[1];

            Excel.Range xlInputFileRange = xlInWorksheet.UsedRange;
            Excel.Range xlOutputFileRange = xlAuxOutWorksheet.UsedRange;

            int rowInputFileCount = xlInputFileRange.Rows.Count;
            int colInputFileCount = xlInputFileRange.Columns.Count;

            int rowAuxFileCount = xlOutputFileRange.Rows.Count;

            string rfInput, rfOutput, noTicket = "";

            //recorre fila No Rf de archivo output
            for (int i = 1; i <= rowAuxFileCount; i++)
            {
                //recorre columna No Rf de archvio output
                for (int j = 1; j == 1; j++)
                {
                    //valida que celda no sea null
                    if (xlOutputFileRange.Cells[i, j] != null && xlOutputFileRange.Cells[i, j].Value2 != null)
                    {
                        //guarda valor a comparar
                        rfOutput = xlOutputFileRange.Cells[i, j].Value2.ToString();
                        //recorre filas No Rf Archivo Input
                        for (int k = 1; k <= rowInputFileCount; k++)
                        {
                            //recorre columnas No Rf Archivo Input
                            for (int l = 2; l == 2; l++)
                            {
                                if (xlInputFileRange.Cells[k, l] != null && xlInputFileRange.Cells[k, l].Value2 != null)
                                {
                                    rfInput = xlInputFileRange.Cells[k, l].Value2.ToString();
                                    //compara y escribe
                                    if (rfInput == rfOutput)
                                    {
                                        noTicket = xlInputFileRange.Cells[k, 1].Value2.ToString();
                                        xlAuxOutWorksheet.Cells[i, 2] = noTicket;

                                        Console.WriteLine("InputCell [{0},{1}] MATCH OuputCell [{2},{3}] -> No. Ticket: {4}", k, l, i, j, noTicket);
                                    }

                                }
                            }
                        }
                    }
                }
            }

            //guardar archivo output
            xlAuxFileOut.Save();

            //cerrar archivo
            xlInputFile.Close(true);
            xlAuxFileOut.Close(true);

            //cerrar excel
            xlApp.Quit();
            xlApp2.Quit();

            killExcelProccess();

        }
        catch (Exception)
        {
            killExcelProccess();
            throw;
        }
    }

    static void Main(string[] args)
    {
        Console.Clear();
        Program app = new Program();

        //String pathInput = "C:\\Users\\Jay\\Desktop\\Diners\\macro-noticket\\SQL\\";
        String pathInput = "E:\\RECURSOS ROBOT\\DATA\\MESA_SERVICIO\\GESTIONDEUSUARIOS\\CONSULTASQL\\";
        

        string validFilePath = "";

        //Valida si recibe parametro para ejecucion reemplaza la ruta quemada por el parametro
        if (args.Length >= 1)
            pathInput = args[0];

        Console.WriteLine("Leyendo directorio {0}...\n", pathInput);
        validFilePath = app.GetInputDir(pathInput);
        app.ReadInput(validFilePath);

        Console.Write("\nSe ha terminado el proceso...");

    }

}