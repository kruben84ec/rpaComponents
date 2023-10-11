using Microsoft.Office.Interop.Excel;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;

class ErrorEnArchivo : Exception
{
    public ErrorEnArchivo(string message) : base(message){    }
}

class Program
{
    
    //kill excel process
    private void killExcelProccess()
    {
        foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
        {
            proc.Kill();
        }
    }

    //busca archivo, valida nomeclatura de nombre, devuelve archivo más reciente
    private string ValidarArchivo(String path)
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

    //open excel file and return worksheet
    private Excel.Workbook OpenXlsWorkbook(String path) {
        try { 
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = false;
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            if(xlWorkbook == null ) {
                throw (new ErrorEnArchivo("Error al abrir el archivo\n"));
            }
            return xlWorkbook;

        }
        catch (Exception) {
            killExcelProccess();
            throw;
        }
    }


    //recorre excel file worksheets and convert columns B and F to Text
    private void ConvertToText(String path) {

        try {

            Excel.Workbook workbook = this.OpenXlsWorkbook(this.ValidarArchivo(path));

            int[] Cols = { 2, 6 }; //Cols B (TICKET), F (IDENTIFICACION)

            foreach (Excel._Worksheet sheet in workbook.Worksheets)
            {
                foreach (Excel.Range row in sheet.UsedRange.Rows)
                {
                    foreach (int c in Cols)
                    {
                        if (sheet.Cells[row.Row, c].Value2 != null) {
                            sheet.Cells[row.Row, c].NumberFormat = "@";
                            Console.WriteLine("Celda {1} a Texto [{0}]", sheet.Cells[row.Row, c].NumberFormat, sheet.Cells[row.Row, c].Value2);
                        }
                    }
                }

            }
            workbook.Save();
            workbook.Close(true);
            this.killExcelProccess();
        }
        catch (Exception){
            throw;
        }

    }

    static void Main(String[] args) { 

        Program app = new Program();

        Console.Clear();

        //convierte a texto las filas de identificacion y no de ticekt 
        app.ConvertToText("\\\\67fb1s2\\ct\\crudas\\");

    }
}