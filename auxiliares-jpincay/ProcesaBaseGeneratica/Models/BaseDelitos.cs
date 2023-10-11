using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProcesaBaseGeneratica.Models
{
    public class BaseDelitos
    {
        public static List<string> ObtenerListaPalabras(string rutaBasePalabras)
        {
            List<string> palabrasJuicios = new();

            Excel.Application excelApp = new()
            {
                Visible = false
            };

            Excel.Workbook wb;
            Excel.Worksheet ws;

            wb = excelApp.Workbooks.Open(rutaBasePalabras);

            ws = wb.Worksheets[1] as Excel.Worksheet;

            for(int row = 2; row <= ws.UsedRange.Rows.Count; row++)
            {
                for(int col = 2; col <= ws.UsedRange.Columns.Count; col++)
                {
                    Excel.Range Cell = ws.Cells[row, col];
                    string palabra = Cell.Value != null ? Utils.NormalizeString(Convert.ToString(Cell.Value2)).ToLower() : "";

                    palabrasJuicios.Add(palabra);
                }
            }

            wb.Close();
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            //GC.Collect();

            return palabrasJuicios;
        }
    }
}
