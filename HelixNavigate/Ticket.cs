using OpenQA.Selenium;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace helixIntegration
{
    public class tikets
    {
        /* 
         * Clase para definir objeto Ticket y propiedades de objeto
         */
        public string idOdt;
        public string aprobador;
        public string operacion;
        public string nombres;
        public string identificacion;
        public string correo;
        public string perfil;
        public string opcionSistema;
        public string usuario;
        public bool fileReaded = false;

        public tikets()
        {
            idOdt = string.Empty;
            aprobador = string.Empty;
            operacion = string.Empty;
            nombres = string.Empty;
            identificacion = string.Empty;
            correo = string.Empty;
            perfil = string.Empty;
            opcionSistema = string.Empty;
            usuario = string.Empty;
        }
    }

    class Ticket : HelperRpa
    {

        private IWebDriver driverInterface;
        public Ticket(IWebDriver driver) : base(driver) => driverInterface = driver;




        public List<tikets> GetSMTickets(string smXlFileFullPath)
        {

                List<tikets> smTickets = new List<tikets>();
            try
            {
                Excel.Application excel = new Excel.Application()
                {
                    Visible = false
                };
                var pathFile = ValidateInputFilePath(smXlFileFullPath);
                Excel.Workbook workbook = excel.Workbooks.Open(pathFile);

                Excel.Worksheet sheet = workbook.Worksheets.Item[1];

                //leer worksheet, crear ticket, obtiene idOdt, agrega a lista a retornar
                int numTickets = sheet.UsedRange.Rows.Count;
                for (int row = 2; row <= numTickets; row++)
                {
                    string ioOdtTicket = Convert.ToString(sheet.Cells[row, 1].Value2);
                    if(ioOdtTicket != null)
                    {
                        tikets smTicket = new tikets()
                        {
                            idOdt = ioOdtTicket,
                            aprobador = "EL AREA DE CONTROL DE ACCESOS AUTORIZA DESDE SERVICE MANAGER",
                            operacion =Convert.ToString(sheet.Cells[row, 2].Value2),
                            nombres = Convert.ToString(sheet.Cells[row, 3].Value2),
                            identificacion = Convert.ToString(sheet.Cells[row, 4].Value2),
                            correo = Convert.ToString(sheet.Cells[row, 5].Value2),
                            perfil = Convert.ToString(sheet.Cells[row, 6].Value2),
                            opcionSistema = Convert.ToString(sheet.Cells[row, 7].Value2),
                            usuario = Convert.ToString(sheet.Cells[row, 8].Value2),
                            fileReaded = true
                        };
                        smTickets.Add(smTicket);
                    }
                }

                workbook.Close(true);

                return smTickets;

            }catch(Exception ex)
            {
                String errorMessage = ex.Message;

                Log.Error(errorMessage);

                return smTickets;
            }

        }

        public bool crear(tikets ticket)
        {
            if(ticket.operacion == "CREAR")
            {
                try
                {
                    Thread.Sleep(7000);
                    findFieldClick("//span[contains(text(), 'Creación')]");
                    Thread.Sleep(7000);
                    findFieldClick("//button[@id=\"eeadec58-a38c-4378-b3df-00410e7c4b22_button\"]");
                    Thread.Sleep(7000);
                    findFieldClick("//span[contains(text(), 'BANCO PICHINCHA')]");
                    Thread.Sleep(7000);
                    findFieldSetText("//input[@id=\"43c5c9de-1fca-43eb-8ff1-943f59e7c755\"]", ticket.aprobador);
                    findFieldClick("//span[contains(text(), 'OK')]");
                    findFieldSetText("//input[@id=\"296a358c-a6ad-42ac-96a3-27e082f0275e\"]", ticket.idOdt);
                    findFieldSetText("//input[@id=\"417cb464-47fe-41bb-8cb7-4069b5b1b701\"]", ticket.nombres);
                    findFieldSetText("//input[@id=\"4f405620-9a63-4b1b-95fb-98fa52f0c6b1\"]", ticket.identificacion);
                    findFieldSetText("//input[@id=\"ec37dbf2-33bf-4bc9-bd83-2b81935045df\"]", ticket.correo);
                    findFieldSetText("//input[@id=\"455e7cbe-7303-4b90-85ba-c5fe75d97187\"]", ticket.perfil);
                    //Presionar el boton de enviar el tiket esperar 5 segundos y luego 10 segundos
                    Thread.Sleep(10000);

                    findButtonClick("/html/body/dwp-root[1]/dwp-checkout[1]/dwp-configurable-content-page[1]/div[1]/div[2]/aside[1]/div[1]/div[1]/div[1]/div[1]/button[1]",
                        6000,
                        12000
                    );
                    return true;

                }
                catch (Exception e)
                {
                    

                    Log.Error($"Error al crear {ticket.idOdt}: {e}");

                    Process.GetCurrentProcess().Kill();
                    Environment.Exit(0);

                    return false;
                }
            }
            return false;
        }

        public bool modificar(tikets ticket)
        {
            if (ticket.operacion == "MODIFICAR")
            {
                try
                {
                    Thread.Sleep(7000);
                    findFieldClick("//span[contains(text(), 'Modificación')]");
                    Thread.Sleep(7000);
                    findFieldClick("//button[@id=\"1a6132a2-7d02-4757-b874-f4e7138a7f83_button\"]");
                    Thread.Sleep(7000);
                    findFieldClick("//span[contains(text(), 'BANCO PICHINCHA')]");
                    Thread.Sleep(7000);
                    findFieldSetText("//input[@id=\"00f73414-f6ce-46fc-a41b-ece9e49c0ef5\"]", ticket.aprobador);
                    findFieldClick("//span[contains(text(), 'OK')]");
                    findFieldSetText("//input[@id=\"d3af4777-c4d5-4fca-8435-e97fb16b7a13\"]", ticket.idOdt);
                    findFieldSetText("//input[@id=\"c31d8600-6493-451a-bd06-aaaf68d35e1a\"]", ticket.nombres);
                    findFieldSetText("//input[@id=\"3b7f3012-8f25-4023-81c2-0f82cc416753\"]", ticket.identificacion);
                    findFieldSetText("//input[@id=\"7aef6da5-078e-44ad-9244-10eca3ea5b2e\"]", ticket.correo);

                    if (ticket.opcionSistema == "SISTEMA GESTOR")
                    {
                        findFieldClick("//span[contains(text(), 'Sistema Gestor')]");
                    }
                    else
                    {
                        findFieldClick("//span[contains(text(), 'Sistema Cao')]");

                    }



                    findFieldClick("//span[contains(text(), 'Perfil')]");
                    Thread.Sleep(5000);
                    findFieldSetText("//input[@id=\"f2adab17-1afe-45ea-b982-d24d7b4b8d85\"]", ticket.perfil);


                    Thread.Sleep(10000);
                    //Presionar el boton de enviar el tiket esperar 5 segundos y luego 10 segundos
                    findButtonClick("/html/body/dwp-root[1]/dwp-checkout[1]/dwp-configurable-content-page[1]/div[1]/div[2]/aside[1]/div[1]/div[1]/div[1]/div[1]/button[1]",
                        6000,
                       12000
                    );
                    return true;

                }
                catch (Exception ex)
                {
                    Log.Error($"Error al modificar {ticket.idOdt}:  {ex}");

                    Process.GetCurrentProcess().Kill();
                    Environment.Exit(0);
                    return false;
                }
            }
            return false;
        }

        public bool eliminar(tikets ticket)
        {
            if (ticket.operacion == "BORRAR")
            {
                try
                {
                    Thread.Sleep(7000);
                    findFieldClick("//span[contains(text(), 'Eliminación')]");
                    Thread.Sleep(7000);
                    findFieldClick("//button[@id=\"127d1a27-bc10-4c72-9392-c4cfe594dad1_button\"]");
                    Thread.Sleep(7000);
                    findFieldClick("//span[contains(text(), 'BANCO PICHINCHA')]");
                    Thread.Sleep(7000);
                    findFieldSetText("//input[@id=\"bc43e96e-db32-42eb-ba3c-e582d5299967\"]", ticket.aprobador);
                    findFieldClick("//span[contains(text(), 'OK')]");
                    findFieldSetText("//input[@id=\"013c0ab3-e7ed-4ac6-b380-f3500b6e183f\"]", ticket.idOdt);
                    findFieldSetText("//input[@id=\"f9c40411-c626-4f12-973d-87b20a5c0631\"]", ticket.nombres);
                    findFieldSetText("//input[@id=\"009f5891-d2c0-4472-b4ec-ff2c913c958c\"]", ticket.identificacion);
                    findFieldSetText("//input[@id=\"ef6eedb2-0013-407f-bd5b-529a4384b502\"]", ticket.usuario);

                    if (ticket.opcionSistema == "SISTEMA GESTOR")
                    {
                        findFieldClick("//span[contains(text(), 'Sistema Gestor')]");
                    }
                    else
                    {
                        findFieldClick("//span[contains(text(), 'Sistema Cao')]");

                    }
                    Thread.Sleep(6000);

                    //Presionar el boton de enviar el tiket esperar 5 segundos y luego 10 segundos
                    findButtonClick("/html/body/dwp-root[1]/dwp-checkout[1]/dwp-configurable-content-page[1]/div[1]/div[2]/aside[1]/div[1]/div[1]/div[1]/div[1]/button[1]",
                         6000,
                         12000
                     );
                    return true;

                }
                catch (Exception ex)
                {
                    Log.Error($"Error en gestion eliminacion ticekt {ticket.idOdt}: {ex}");

                    Process.GetCurrentProcess().Kill();
                    Environment.Exit(0);
                    return false;
                }
            }
            return false;
        }


    }
}
