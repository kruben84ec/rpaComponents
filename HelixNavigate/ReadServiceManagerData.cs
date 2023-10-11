
using Excel = Microsoft.Office.Interop.Excel;

namespace DllReadServiceManagerData
{
    public class SMTicket
    {
        /* 
         * Clase para definir objeto Ticket y propiedades de objeto
         */
        public string idOdt;
        public string operacion;
        public string nombres;
        public string identificacion;
        public string correo;
        public string perfil;
        public string opcionSistema;
        public string usuario;
        public bool fileReaded = false;

        public SMTicket()
        {
            idOdt = string.Empty;
            operacion = string.Empty;
            nombres = string.Empty;
            identificacion = string.Empty;
            correo = string.Empty;
            perfil = string.Empty;
            opcionSistema = string.Empty;
            usuario = string.Empty;
        }
    }

    public class ReadServiceManagerData
    {

        public static string ValidateInputFilePath(String path, String fileName)
        {
            /* 
             * valida ruta, obtiene archivo mas reciente y retorna  ruta completa
             */

            string recentFilePath = "";

            try
            {

                //Lee directorio en busqueda de archivo mas reciente
                var directory = new DirectoryInfo(path);
                recentFilePath = (from f in directory.GetFiles() where f.Name == fileName orderby f.LastWriteTime descending select f).First().ToString();
                return recentFilePath;

            }
            catch (Exception)
            {
                throw;
            }

        }

        public List<SMTicket>? GetSMTickets(string smXlFileFullPath, string smXlFileName)
        {
            List<SMTicket> smTickets = new();

            try
            {
                Excel.Application excel = new()
                {
                    Visible = false
                };
                Excel.Workbook workbook = excel.Workbooks.Open(
                    ValidateInputFilePath(smXlFileFullPath, smXlFileName));

                Excel.Worksheet sheet = workbook.Worksheets.Item[1];

                //leer worksheet, crear ticket, obtiene idOdt, agrega a lista a retornar
                for (int row = 2; row <= sheet.UsedRange.Rows.Count; row++)
                {

                    SMTicket smTicket = new()
                    {
                        idOdt = sheet.Cells[row, 1].Value2,
                        operacion = sheet.Cells[row, 2].Value2,
                        nombres = sheet.Cells[row, 3].Value2,
                        identificacion = sheet.Cells[row, 4].Value2,
                        correo = sheet.Cells[row, 5].Value2,
                        perfil = sheet.Cells[row, 6].Value2,
                        opcionSistema = sheet.Cells[row, 7].Value2,
                        usuario = sheet.Cells[row, 8].Value2,
                        fileReaded = true
                    };

                    smTickets.Add(smTicket);

                }

                workbook.Close(false);

                return smTickets;

            }
            catch (Exception)
            {
                return null;
            }
        }


    }
}