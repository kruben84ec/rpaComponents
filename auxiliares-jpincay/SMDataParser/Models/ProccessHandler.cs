using Serilog;
using SMDataParser.Config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMDataParser.Models
{
    internal class ProccessHandler
    {
        public void KillExcelProccess()
        {
            try {
                Log.Information($"KillExcelProccess() Cerrando procesos excel...");
                foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    Log.Information($"Cerrando proceso {proc.ProcessName} {proc.Id}");
                    proc.Kill();
                }
                
            }catch(Exception e)
            {
                Log.Error($"KillExcelProccess()"  +
                    $"\nError: {e.ToString()}");
            }
        }

    }
}
