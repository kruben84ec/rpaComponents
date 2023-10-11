using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

public class ExcelSplitter
{
    private static readonly string[] UniqueIDs = { "1030", "1130", "1230", "1430", "1530" };

    public static void SplitExcelFile(string sourceFilePath, int clientsPerFile, string outputDirectory)
    {
        // Create Excel application object
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook sourceWorkbook = null;

        try
        {
            // Open the source Excel file
            sourceWorkbook = excelApp.Workbooks.Open(sourceFilePath);
            Excel.Worksheet sourceWorksheet = sourceWorkbook.ActiveSheet;

            int totalRecords = sourceWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            // Determine the number of unique clients in the source worksheet
            HashSet<string> uniqueClients = new HashSet<string>();
            for (int i = 4; i <= totalRecords; i++) // Assuming client ID column is in column C (index 3) and data starts from row 4
            {
                string clientID = sourceWorksheet.Cells[i, 3].Value?.ToString();
                if (!string.IsNullOrEmpty(clientID))
                {
                    uniqueClients.Add(clientID);
                }
            }

            int clientCount = uniqueClients.Count;
            int filesPerDay = 5;
            int filesPerID = 6;
            int daysPerIncrement = 1;
            int totalFiles = filesPerID * UniqueIDs.Length;
            int currentClient = 0;
            DateTime currentDate = DateTime.Now.Date.AddDays(1);

            for (int i = 0; i < totalFiles; i++)
            {
                string uniqueID = UniqueIDs[i % UniqueIDs.Length];
                string fileName = GetUniqueFileName(currentDate, uniqueID);
                string filePath = Path.Combine(outputDirectory, fileName);

                // Create a new workbook for each split file
                Excel.Workbook splitWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet splitWorksheet = splitWorkbook.ActiveSheet;

                int clientsToWrite = Math.Min(clientsPerFile, clientCount - currentClient);

                // Write headers to the split worksheet
                Excel.Range sourceHeadersRange = sourceWorksheet.Range["A3:Z3"]; // Assuming headers are in row 3
                Excel.Range destinationHeadersRange = splitWorksheet.Range["A3:Z3"];
                sourceHeadersRange.Copy(destinationHeadersRange);

                int currentRow = 4; // Start writing data from row 3

                for (int j = currentClient; j < currentClient + clientsToWrite; j++)
                {
                    string clientID = uniqueClients.ElementAt(j);

                    // Copy the records for the current client from the source worksheet to the split worksheet
                    for (int k = 4; k <= totalRecords; k++) // Assuming client ID column is in column C (index 3) and data starts from row 4
                    {
                        string currentClientID = sourceWorksheet.Cells[k, 3].Value?.ToString();
                        if (currentClientID == clientID)
                        {
                            Excel.Range sourceRange = sourceWorksheet.Range["A" + k.ToString(), "Z" + k.ToString()];
                            Excel.Range destinationRange = splitWorksheet.Cells[currentRow, 1];
                            sourceRange.Copy(destinationRange);
                            currentRow++;
                        }
                    }
                }

                // Save the split file with a unique name in the specified output directory
                splitWorkbook.SaveAs(filePath);

                if (File.Exists(filePath))
                {
                    Log.Information($"Archivo base corte generado: {filePath}");
                }

                // Close the split workbook
                splitWorkbook.Close();

                // Increment the current client index
                currentClient += clientsToWrite;

                // Check if the date needs to be incremented
                if ((i + 1) % filesPerDay == 0)
                {
                    currentDate = currentDate.AddDays(daysPerIncrement);
                }
            }
        }
        finally
        {
            // Close and release the Excel objects
            if (sourceWorkbook != null)
            {
                sourceWorkbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceWorkbook);
            }

            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            GC.Collect();
        }
    }

    private static string GetUniqueFileName(DateTime date, string uniqueID)
    {
        string dateFormatted = date.ToString("yyyyMMdd");
        return $"{dateFormatted}_{uniqueID}.xlsx";
    }
}
