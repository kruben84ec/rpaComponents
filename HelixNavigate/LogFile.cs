using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace helixIntegration
{
    internal class LogFile
    {
        public DateTime Date { get; set; }
        public string Message { get; set; }
        public string Level { get; set; }

        public LogFile(DateTime date, string message, string level)
        {
            Date = date;
            Message = message;
            Level = level;
        }

        public string getPathLog()
        {
            string path = @"E:\RECURSOS ROBOT\LOGS\MESA_SERVICIO\GESTIONDEUSUARIOS\\";
            string fecha_log = $@"{DateTime.Now:yyyy-M-d}\\log_navigate_helix.txt";
            string logPathFinal = Path.Combine(path, fecha_log);
            return logPathFinal;
        }

        public static void WriteToLog(LogFile log)
        {
            string path = @"E:\RECURSOS ROBOT\LOGS\MESA_SERVICIO\GESTIONDEUSUARIOS\\";
            string fecha_log = $@"{DateTime.Now:yyyy-M-d}\\log_navigate_helix.txt";

            string logPathFinal = Path.Combine(path, fecha_log);

            using (StreamWriter sw = new StreamWriter(logPathFinal, true))
            {
                sw.WriteLine($"{log.Date} {log.Level}: {log.Message}");
            }
        }
    }
}
