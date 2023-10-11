using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace createDirectory
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                string pathDesteny = args[0];
                ServiceDirectory.createDirectoryLog(pathDesteny);
            }
        }
    }
}
