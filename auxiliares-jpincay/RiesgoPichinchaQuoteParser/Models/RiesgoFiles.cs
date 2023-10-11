using RiesgoPichinchaQuoteParser.Config;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RiesgoPichinchaQuoteParser.Models
{
    internal class RiesgoFile
    {

        public string name;
        public ArrayList textContent;
        public string outPath;
        public string outFileName;
        public string fullOutPathName;
        
        AppConfig appConfig = new AppConfig();

        public RiesgoFile(string name, ArrayList textContent)
        {
            this.name = name;
            this.textContent = textContent;
            this.outPath = appConfig.outputPath;
            this.outFileName = name + ".txt";
            this.fullOutPathName = this.outPath + this.outFileName;
        }

        public RiesgoFile()
        {
            this.name = "";
            this.textContent = new ArrayList();
            this.outPath = "";
            this.outFileName = "";
            this.fullOutPathName = "";
        }


    }
}
