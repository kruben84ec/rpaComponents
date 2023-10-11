using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

public class ExcelDuplicateFilter
{
    public static void FilterDuplicates(string sourceFilePath, string outputFilePath)
    {
        // Create Excel application object
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook sourceWorkbook = null;
        Excel.Workbook outputWorkbook = null;

        try
        {
            Log.Information($"Obteniendo data de base generatica");

            // Open the source Excel file
            sourceWorkbook = excelApp.Workbooks.Open(sourceFilePath);
            Excel.Worksheet sourceWorksheet = sourceWorkbook.ActiveSheet;

            // Get the used range of the source worksheet
            Excel.Range usedRange = sourceWorksheet.UsedRange;

            // Get the last row in the used range
            int lastRow = usedRange.Rows.Count + 2;

            // Create a dictionary to store the unique IDs and their corresponding values
            Dictionary<string, List<string>> duplicateData = new Dictionary<string, List<string>>();

            // Loop through the rows in the used range starting from the second row
            for (int row = 4; row <= lastRow; row++)
            {
                // Get the value in column "C" for the current row
                string id = sourceWorksheet.Cells[row, 3].Value?.ToString().Trim();

                // Skip empty values
                if (string.IsNullOrEmpty(id))
                {
                    continue;
                }

                // Check if the ID already exists in the dictionary
                if (duplicateData.ContainsKey(id))
                {
                    continue;
                }
                else
                {
                    // Create a new list for the ID and add the corresponding values from columns "B" and "D"
                    string valueB = sourceWorksheet.Cells[row, 2].Value2?.ToString().Trim();
                    string valueD = sourceWorksheet.Cells[row, 4].Value2?.ToString().Trim();
                    duplicateData.Add(id, new List<string> { valueB, valueD });
                }
            }

            // Create a new workbook for the output file
            outputWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet outputWorksheet = outputWorkbook.ActiveSheet;

            Log.Information($"Generando reporte base");

            // Write the headers to the output worksheet
            string[] headers = { "CasoId", "ID Principal", "Identificacion", "Nombre", "Descripcion Mitigacion (antes)", "Descripcion Mitigacion (actual)", "Fecha Vencimiento", "Observacion", "Delitos encontrados", "Comentario" };
            for (int column = 1; column <= headers.Length; column++)
            {
                outputWorksheet.Cells[1, column] = headers[column - 1];
            }

            Log.Information($"Escribiendo data");

            // Write the filtered data to the output worksheet
            int outputRow = 2;
            foreach (var entry in duplicateData)
            {
                string id = entry.Key;
                List<string> values = entry.Value;

                if (values.Count > 1)
                {
                    // Write the ID in the "C" column
                    outputWorksheet.Cells[outputRow, 3].NumberFormat = "@";
                    outputWorksheet.Cells[outputRow, 2].NumberFormat = "@";
                    outputWorksheet.Cells[outputRow, 4].NumberFormat = "@";

                    outputWorksheet.Cells[outputRow, 3] = id;
                    outputWorksheet.Cells[outputRow, 2] = values[0];
                    outputWorksheet.Cells[outputRow, 4] = values[1];

                    outputRow++;

                }
            }

            // Save the output workbook to the specified output file
            outputWorkbook.SaveAs(outputFilePath);

            Log.Information("Reporte base generado: " + outputFilePath);

        }
        catch (Exception ex)
        {
            Log.Error($"Error: {ex.Message}\n{ex.StackTrace}");
        }
        finally
        {
            // Close and release resources
            sourceWorkbook.Close();
            outputWorkbook.Close();
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            GC.Collect();

        }
    }

}
