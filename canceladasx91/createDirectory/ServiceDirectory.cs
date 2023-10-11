using System;
using System.Collections.Generic;
using System.IO;


namespace createDirectory
{
    class ServiceDirectory
    {
        private const string macroCanceladas = "MacroCX91.xlsm";
        private const string reporteCanceladas = "reporte_.xlsx";



        public static string createDirectory(string path)
        {
            string pathDestinyLog = "";
            var dateTimeNow = getDate();
            string yearFolder = (string)dateTimeNow["year"];
            string nameFolder = (string)dateTimeNow["nameMouth"];
            string dateFolder = (string)dateTimeNow["dateNow"];

            pathDestinyLog += path + yearFolder;
            pathDestinyLog += @"\" + nameFolder;
            pathDestinyLog += @"\" + dateFolder;

            if (!Directory.Exists(pathDestinyLog))
            {
                Directory.CreateDirectory(pathDestinyLog);
                Log.LogWrite(path + @"\log_bot\crear_directorios.txt", "Se creo el directorio");
            }
            else
            {
                Log.LogWrite(path + @"\log_bot\crear_directorios.txt", "Ya existe el Directrorio:"+pathDestinyLog);

            }
            return pathDestinyLog;
        }

        public static Dictionary<string, object> getDate()
        {
            var dateTime = DateTime.Now;
            var dateNow = dateTime.ToString("yyyy-MM-dd");
            var year = dateTime.ToString("yyyy");
            var nameMouth = dateTime.ToString("MMMM").ToUpper();
            var dateCompleteNow = new Dictionary<string, object>();
            dateCompleteNow.Add("dateNow", dateNow);
            dateCompleteNow.Add("year", year);
            dateCompleteNow.Add("nameMouth", nameMouth);
            return dateCompleteNow;
        }

        public static void createDirectoryLog(string pathDestiny)
        {

            string pathDestinyLog = createDirectory(@pathDestiny);
            string pathLog = Path.Combine(pathDestiny, @"insumos\" + macroCanceladas);
            string pathReport = Path.Combine(pathDestiny,@"insumos\" + reporteCanceladas);

   
            try
            {

                var userUploadsDir = pathDestinyLog+ @"\" + macroCanceladas;
                var fullDirPath = Path.GetFullPath(userUploadsDir);
                var pathReportExist = pathDestinyLog+ @"\" + reporteCanceladas;
                var fullDirPathReport = Path.GetFullPath(pathReportExist);



                if (fullDirPath.StartsWith(userUploadsDir, StringComparison.Ordinal))
                {

                    File.Delete(fullDirPath);
                    File.Copy(pathLog, fullDirPath);
                    Log.LogWrite(pathDestiny + @"\log_bot\crear_directorios.txt", "Copiando y eliminando el template de la macro");

                }


                if (pathReportExist.StartsWith(fullDirPathReport, StringComparison.Ordinal) && !File.Exists(fullDirPathReport))
                {
                    File.Copy(pathReport, pathReportExist);
                    Log.LogWrite(pathDestiny + @"\log_bot\crear_directorios.txt", "Copiando el reporte de ejecución");
                }


            }
            catch (InvalidOperationException e)
            {
                Log.LogWrite(pathDestiny + @"\log_bot\crear_directorios.txt", e.ToString());
            }
            

        }

        public static void createFolders(string pathDestiny)
        {
            createDirectoryLog(pathDestiny);
        }
    }
}
