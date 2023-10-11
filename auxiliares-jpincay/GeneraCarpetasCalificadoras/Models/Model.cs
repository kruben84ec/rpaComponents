using Excel = Microsoft.Office.Interop.Excel;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace GeneraCarpetasCalificadoras.Models
{
    class Model
    {
        //quita las tildes de una cadena
        static string NormalizeString(string cadena) => Regex.Replace(cadena.Normalize(NormalizationForm.FormD), @"[^a-zA-z0-9 ]+", "");
       
        private static void CrearCarpeta(string folderPath)
        {
            try
            {
                if(File.Exists(folderPath))
                {
                    Log.Warning($"Carpeta ya existe: {folderPath}");
                }
                else
                {
                    Directory.CreateDirectory(folderPath);
                    Log.Information($"Carpeta creada {folderPath}");
                }
            }catch (Exception e)
            {
                Log.Error($"{System.Reflection.MethodBase.GetCurrentMethod().Name}() Error {e}");
            }
        }


        public void CrearEstructura(List<string> folderNamesList, string baseFolderPath)
        {
            try
            {

                if(!Directory.Exists(baseFolderPath))
                {
                    throw new Exception($"Directorio base para creacion de carpetas no existe: {baseFolderPath}");
                }

                foreach(string folderName in folderNamesList)
                {
                    CrearCarpeta( Path.Combine(baseFolderPath, folderName) );
                }


            }catch (Exception e)
            {
                Log.Error($"{System.Reflection.MethodBase.GetCurrentMethod().Name}() Error {e}");
            }
        }

        public List<string> GetFolderNames(string xlConfigFilePath)
        {
            List<string> namesLists = new();
            List<string> folderNamesLists = new();

            try
            {
                if (!File.Exists(xlConfigFilePath))
                {
                    throw new Exception($"Archivo no existe {xlConfigFilePath}");
                }

                Excel.Application xlApp = new() { 
                    Visible = false
                };

                Excel.Workbook wb = xlApp.Workbooks.Open(xlConfigFilePath);

                Excel.Worksheet sheet = wb.Worksheets["checklistRPA"];

                for(int row = 1; row <= sheet.UsedRange.Rows.Count; row++)
                {
                    if(sheet.Cells[row, 1].Value2 != null)
                    {
                        namesLists.Add(NormalizeString(Convert.ToString(sheet.Cells[row, 1].Value2.ToLower())));
                    }
                    else
                    {
                        Log.Warning($"Nombre de carpeta no encontrado. Revisar archivo configuracion {xlConfigFilePath}...");
                        continue;
                    }
                }

                folderNamesLists = namesLists.Distinct().ToList();


                wb.Close();
                xlApp.Quit();

            }catch (Exception e)
            {
                Log.Error($"{System.Reflection.MethodBase.GetCurrentMethod().Name}() Error {e}");
                return folderNamesLists;
            }

            return folderNamesLists;
        }
    }
}
