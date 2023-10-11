using System;
using System.IO;
namespace createDirectory
{
    class Log
    {
        public static void LogWrite(String pathFile, String messageLog)
        {

            var logWrite = "";

            //Obtener la hora del sistema
            var dateTime = DateTime.Now;
            var dateLogWrite = dateTime.ToString("yyyy-MM-dd hh:mm:ss");

            if (pathFile != "" && messageLog != "")
            {
                using (StreamWriter sw = File.AppendText(Path.GetFullPath(pathFile)))
                {
                    logWrite += dateLogWrite + ";";
                    logWrite += messageLog + ";";
                    sw.WriteLine(logWrite);
                }
            }
        }
    }
}
