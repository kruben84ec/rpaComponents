using Serilog;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.UI;
using Microsoft.Office.Interop.Excel;
using RPAHelper = ActualizarReqSM.NavigatorSM.HelperRpa;
using OpenQA.Selenium.Chrome;
using ActualizarReqSM.Models;
using DriverWeb = ActualizarReqSM.NavigatorSM.DriverWeb;
using SeleniumExtras.WaitHelpers;

namespace ActualizarReqSM.NavigatorSM
{
    class WebSM {

        public string urlSM;
        public string userSM;
        public string passSM;

        private string loggedInURL;
        private string mainPanel;
        private string inputBuscarPeticion;
        private double defaultWaitTime;

        private ChromeDriver driver;

        private RPAHelper helperRpa;

        public WebSM() { 

            urlSM = string.Empty;
            userSM = string.Empty;
            passSM = string.Empty;

            loggedInURL = "https://smgestion.uio.bpichincha.com/sm960/index.do";
            mainPanel = "//*[@id='ext-gen-top53']/em/span/span";
            //inputBuscarPeticion = "//*[@id=\"X21\"]";
            inputBuscarPeticion = "//*[@id=\"X21\"]";
            defaultWaitTime = 10;

            driver = new DriverWeb().GetChromeDriver();
            helperRpa = new(driver);

        }

        public void AbrirPeticion(HelixTicket ticket)
        {
            try
            {
                
                BuscarPeticion(ticket);

            }catch (Exception e)
            {
                throw new Exception($"AbrirPeticion() Error: No se pudo abrir panel y buscar peticion...");
            }
        }

        private IWebElement FluentWaitElement(string elementXpath)
        {
            DefaultWait<IWebDriver> fluentWait = new DefaultWait<IWebDriver>(driver)
            {
                Timeout = TimeSpan.FromSeconds(defaultWaitTime),
                PollingInterval = TimeSpan.FromMilliseconds(250),
                Message = "Elemento no encontrado"
            };

            fluentWait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            fluentWait.IgnoreExceptionTypes(typeof(TimeoutException));

            //IWebElement iwelement = fluentWait.Until(x => x.FindElement(By.XPath(elementXpath)));
            IWebElement iwelement = fluentWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(elementXpath)));

            return iwelement;

        }

        private void ClicElement(string elementXpath)
        {
            bool displayed;
            IWebElement element = FluentWaitElement(elementXpath);
            displayed = false;
            while (displayed == false)
            {
                element = FluentWaitElement(elementXpath);
                if (element.Displayed)
                {
                    displayed = true;
                    break;
                }
                Log.Information("Esperando elemento...");
            }

            element.Click();

            Log.Information($"Element {elementXpath} clicked...");

        }


        private void SetTextElement(string elementXpath, string text)
        {
            ClicElement(elementXpath);
            IWebElement element = FluentWaitElement(elementXpath);
            element.Clear();
            Log.Information($"Element {elementXpath} cleared...");

            element.SendKeys(text);
            Log.Information($"Text set to element {elementXpath}");


        }

        public void AbrirPanelBusquedaPeticion()
        {
            bool displayed;
            try
            {

                //// clic en panel cumplimiento peticiones xpath: //*[@id=\"ROOT/Cumplimiento de peticiones\"]
                IWebElement btnCumplimientoPeticiones = FluentWaitElement("//*[@id=\"ROOT/Cumplimiento de peticiones\"]");
                displayed = false;
                while (displayed == false)
                {
                    btnCumplimientoPeticiones = FluentWaitElement("//*[@id=\"ROOT/Cumplimiento de peticiones\"]");
                    if (btnCumplimientoPeticiones.Displayed)
                    {
                        displayed = true;
                        break;
                    }
                    Log.Information("esperando boton panel cumplimiento peticiones");
                }

                btnCumplimientoPeticiones.Click();

                //// clic en opcion buscar peticiones xpath: //*[@id=\"ROOT/Cumplimiento de peticiones/Buscar peticiones\"]
                IWebElement btnBuscarPeticiones = FluentWaitElement("//*[@id=\"ROOT/Cumplimiento de peticiones/Buscar peticiones\"]");
                displayed = false;
                while (!btnBuscarPeticiones.Displayed)
                {
                    btnBuscarPeticiones = FluentWaitElement("//*[@id=\"ROOT/Cumplimiento de peticiones/Buscar peticiones\"]");
                    if (btnBuscarPeticiones.Displayed)
                    {
                        displayed = true;
                        break;
                    }
                    Log.Information("esperando boton buscar peticiones");
                }
                btnBuscarPeticiones.Click();


                ////2	Verificar que existe iFrame que contiene panel de busqueda y switch a es iFrame: 
                if (ElementDisplayed("//iframe[@title=\"¿Qué petición desea mostrar?\"]"))
                {
                    Log.Information($"Panel de busqueda de peticiones abierto");
                }
                else
                {
                    AbrirPanelBusquedaPeticion();

                }


            }
            catch (Exception e) {
                Log.Error($"AbrirPanelBusquedaPeticion() Error: No se pudo abrir panel de busqueda de peticiones\n{e}");
            }
        }

        public bool ElementDisplayed(string xPathElement)
        {
            IWebElement panelBusqueda = FluentWaitElement(xPathElement);

            return panelBusqueda.Displayed;
        }

        public void SwitchFrame(string frameXpath)
        {
            bool displayed;

            IWebElement frame = FluentWaitElement(frameXpath);
            
            while (!frame.Displayed)
            {
                frame = FluentWaitElement(frameXpath);
                if (frame.Displayed)
                {
                    displayed = true;
                    break;
                }
                Log.Information($"Esperando frame {frameXpath}");
            }

            driver.SwitchTo().Frame(frame);
            Log.Information($"Switched to Frame {frameXpath}");

        }

        public bool isFase(HelixTicket ticket)
        {
            bool isFase = false;
            try
            {
                string faseInput = "/html/body/div[1]/div[1]/form/div[1]/div/div/div[5]/div/div[1]/div[2]/input";
                IWebElement faseElement = FluentWaitElement(faseInput);
                String fase = faseElement.GetAttribute("value");

                if (fase == "Registro" ) isFase = true;
            }
            catch (Exception e)
            {
                Log.Error($"BuscarPeticion() Error: No se pudo buscar peticion\n{e}");
            }finally {
                Log.Error($"El ticket:{ticket.idOdt} es {isFase.ToString()}");

            }
            return isFase ; 
        }


        public bool BuscarPeticion(HelixTicket ticket)
        {
            bool isUpadte = false;
            try
            {

                if (ElementDisplayed("//iframe[@title=\"¿Qué petición desea mostrar?\"]"))
                {
                    SwitchFrame("//iframe[@title=\"¿Qué petición desea mostrar?\"]");
                    if (ElementDisplayed(inputBuscarPeticion))
                    {
                        Log.Information($"Buscando peticion {ticket.idOdt}...");

                        //ingresa idodt a buscar
                        SetTextElement(inputBuscarPeticion, ticket.idOdt);
                        //Buscar información del ticket
                       var frameAct = driver.FindElement(By.TagName("body"));
                       frameAct.SendKeys(Keys.Control + Keys.Shift + Keys.F6);
                       Thread.Sleep(1000);
                       isUpadte = isFase(ticket);
                    }
                    else
                    {
                        Log.Error($"BuscarPeticion() Error: Elemento inputBuscarPeticion no se encontro..");
                        BuscarPeticion(ticket);
                    }

                }
                else
                {
                    Log.Error($"Panel de busqueda no seleccionado...");
                }

                //espera input para buscar peticion

            }
            catch (Exception e)
            {
                Log.Error($"BuscarPeticion() Error: No se pudo buscar peticion\n{e}");
            }
            return isUpadte;
        }

        public void ActualizarPeticion(HelixTicket ticket)
        {

            try
            {
                Log.Information($"Ingresando a panel actividades..."+ ticket.idOdt);
                string actualizacion = "Actualizacion RPA " + ticket.noReq;

                Thread.Sleep(1000);
                // clic tab actividades
                String seccionActiviades = "/html/body/div[1]/div[1]/form/div[2]/div[1]/div[2]/table/tbody/tr/td[2]/a";
                helperRpa.ClickWaitField(seccionActiviades, defaultWaitTime);
     
                //registrar palabra actualizar
                String nuevoTipoActualizacion = "/html/body/div[1]/div[1]/form/div[2]/div[3]/div/div/div[1]/div/div[1]/input";
      
                IWebElement tipoActualizacion = FluentWaitElement(nuevoTipoActualizacion);
                tipoActualizacion.SendKeys("Actualizar");


                Log.Information($"Actualizando peticion: {actualizacion}");

                Thread.Sleep(1000);
                // registra actualizacion
                String nuevaActulizacion = "/html/body/div[1]/div[1]/form/div[2]/div[3]/div/div/div[3]/div/div";
                IWebElement nuevaActulizacionTextArea = FluentWaitElement(nuevaActulizacion);
                nuevaActulizacionTextArea.SendKeys(actualizacion);
                Thread.Sleep(2000);




                Log.Information($"Guardando actualizacion...");
                //Buscar información del ticket
                var frameAct = driver.FindElement(By.TagName("body"));
                frameAct.SendKeys(Keys.Control + Keys.Shift + Keys.F2);
   
               
                Thread.Sleep(2000);

              

            }
            catch (Exception e)
            {
                Log.Error($"ActualizarPeticion() Error: Error en la actualizacion de la peticion {ticket.idOdt}\n{e}");
            }

        }

        public void LogOutSM()
        {
            try {

                //var frameAct = driver.FindElement(By.TagName("body"));

                //cierra panel actualizacion
                //frameAct.SendKeys(Keys.Alt + Keys.F3);

                Log.Information($"Cerrando sesion...");

                //Thread.Sleep(5000);

                driver.SwitchTo().Frame(0);

                helperRpa.findFieldClick("//*[@id=\"ext-gen-top104\"]");

                // //*[@id=\"ext-gen-top374\"]
                helperRpa.findFieldClick("//button[contains(text(),\"Desconexión\")]");

                IAlert iAlert = driver.SwitchTo().Alert();

                iAlert.Accept();

                Thread.Sleep(5000);

                driver.Dispose();

            } catch(Exception e) 
            {
                Log.Error($"LogOutSM() Error: No se pudo cerra la sesion\n{e}");
            }
        
        }

        public bool LogInSM(string urlSM, string userSM, string passSM)
        {
            bool loggedIn = false;

            try {

                driver.Navigate().GoToUrl(urlSM);


                IWebElement userInput = FluentWaitElement("//*[@id=\"LoginUsername\"]");
                userInput.SendKeys(userSM);

                IWebElement passInput = FluentWaitElement("//*[@id=\"LoginPassword\"]");
                passInput.SendKeys(passSM);

                passInput.Submit();

                IWebElement framePrincipal = FluentWaitElement("//iframe[@title=\"Petición Cola: vPeticionesRPA_SM_H\"]");
                if (framePrincipal.Displayed)
                {
                    loggedIn = true;
                }

                return loggedIn;

            } catch (Exception e)
            {
                Log.Error($"{e}");
                return loggedIn;
            }


        }

    }





}
