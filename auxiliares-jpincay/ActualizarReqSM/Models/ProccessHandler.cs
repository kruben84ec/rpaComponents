using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActualizarReqSM.Models
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
                    $"\nError: {e}");
            }
        }

    }
}
