using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Net.Mail;
using System.Net;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Runtime.InteropServices;
using System.Globalization;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using HtmlAgilityPack;
using System.IO;
using System.Reflection.PortableExecutable;
using System.Text;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Canvas.Parser;
using WS.Models;

namespace WS.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ResidentesController : ControllerBase
    {
        [HttpPost]
        [Route("Guardar")]
        public async Task<IActionResult> GuardarAsync()
        {
            var responseAPI = new ResponseAPI<MPDF.Formato>();
            try
            {
                String rutaInstalador = Environment.CurrentDirectory;

                //Setup driver
                var options = new ChromeOptions();
                options.AddUserProfilePreference("download.default_directory", @"C:\Users\Kochi\Downloads\Pruebas");
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("download.extensions_to_open", "application/pdf");
                options.AddUserProfilePreference("plugins.always_open_pdf_externally", true);
                //options.AddArgument("--headless"); // Ejecución sin cabeza, es decir, sin ventana del navegador visible


                //var chromeDriverService = ChromeDriverService.CreateDefaultService(@"C:\Users\Kochi\Downloads\chromedriver-win64\chromedriver-win64");
                var chromeDriverService = ChromeDriverService.CreateDefaultService(rutaInstalador);
                var driver = new ChromeDriver(chromeDriverService);
                //var driver = new ChromeDriver(options);

                //Navigation on web
                driver.Navigate().GoToUrl("https://sat.gob.mx");
                // Add any additional navigation or waiting logic needed to reach the specific URL.
                driver.Navigate().GoToUrl("https://ptsc32d.clouda.sat.gob.mx/?/reporteOpinion32DContribuyente");

                string cer = $"{rutaInstalador}\\cer.cer";
                string key = $"{rutaInstalador}\\key.key";
                int bandArchivos = 0;

                // Encontrar el elemento de carga de archivos
                var fileInputElement = driver.FindElement(By.Id("txtCertificate"));
                fileInputElement.Click();
                // Esperar un momento para que se abra el cuadro de diálogo
                Thread.Sleep(1000);
                // Llamar a la función que maneja el cuadro de diálogo
                bandArchivos += SetFileUploadDialog(cer);

                // Encontrar el elemento de carga de archivos
                fileInputElement = driver.FindElement(By.Id("txtPrivateKey"));
                fileInputElement.Click();
                // Esperar un momento para que se abra el cuadro de diálogo
                Thread.Sleep(1000);
                bandArchivos += SetFileUploadDialog(key);
                driver.FindElement(By.Id("privateKeyPassword")).SendKeys("");

                //Si es dos es por que se cargaron los dos archivos correctamente
                if (bandArchivos == 2)
                {
                    //Presionamos el boton loggin
                    fileInputElement = driver.FindElement(By.Id("submit"));
                    fileInputElement.Click();

                    Thread.Sleep(10000);
                    // Localizar el elemento del visor PDF que contiene la URL del PDF
                    var pdfViewer = driver.FindElement(By.CssSelector("embed[type='application/pdf'], iframe"));

                    // Obtener la URL del PDF
                    string pdfUrl = pdfViewer.GetAttribute("src");
                    string[] pdfArray = pdfUrl.Split(',');
                    pdfUrl = pdfArray[1];
                    pdfUrl = pdfUrl.ToString();
                    Console.WriteLine($"PDF URL: {pdfUrl}");

                    // Descargar el PDF
                    string downloadDirectory = @"C:\Users\Kochi\Downloads\Pruebas\ROGV041113EK5.pdf";

                    byte[] pdfBytes = Convert.FromBase64String(pdfUrl);
                    // Guardar bytes decodificados como archivo PDF
                    System.IO.File.WriteAllBytes(downloadDirectory, pdfBytes);

                    // Cerrar el navegador
                    driver.Quit();

                    responseAPI.JsonResponse = LlenarClasePDF(downloadDirectory);
                }
            }
            catch (Exception ex)
            {

                responseAPI.EsCorrecto = false;
                responseAPI.Mensaje = ex.Message;
            }

            return Ok(responseAPI);
        }
        private MPDF.Formato LlenarClasePDF(string rutaPDF)
        {
            MPDF.Formato pDF = new MPDF.Formato();
            pDF.obligaciones = new List<MPDF.Obligaciones>();

            try
            {
                //DatosFicticios
                /*
                pDF.anio = 2016;
                pDF.anio = 2017;
                pDF.anio = 2018;
                pDF.anio = 2019;
                pDF.anio = 2020;
                pDF.anio = 2021;
                pDF.anio = 2022;
                pDF.anio = 2023;
                pDF.anio = 2024;
                */

                string valorBool = "";
                string datosPDF = getDatosPDF(rutaPDF);
                string[] boolArray = datosPDF.ToLower().Split("en sentido ");
                string[] boolArrayP = boolArray[1].ToLower().Split(".");

                valorBool = boolArrayP[0];
                if (valorBool.Contains("NEGATIVO"))
                    pDF.contribuyente = false;
                else
                    pDF.contribuyente = true;

                try
                {
                    string[] boolArrayC = datosPDF.ToLower().Split("declaración de proveedores de ");
                    string[] boolArrayPC = boolArrayC[1].ToLower().Split("notas");
                    bool iva = true, isr = true, reten = true, diot = true, declar = true;

                    valorBool = boolArrayPC[0];
                    //pDF.obligaciones = new List<MPDF.Obligaciones>();
                    if (valorBool.Contains("iva"))
                        iva = false;
                    if (valorBool.Contains("isr"))
                        isr = false;
                    if (valorBool.Contains("retecion"))
                        reten = false;
                    if (valorBool.Contains("diot"))
                        diot = false;
                    if (valorBool.Contains("declaracion"))
                        declar = false;

                    pDF.obligaciones.Add(new MPDF.Obligaciones
                    {
                        IVA = iva,
                        ISR = isr,
                        Retenciones = reten,
                        DiOT = diot,
                        declaracionAnual = declar,
                    });
                }
                catch
                {
                    pDF.obligaciones.Add(new MPDF.Obligaciones
                    {
                        IVA = true,
                        ISR = true,
                        Retenciones = true,
                        DiOT = true,
                        declaracionAnual = true,
                    });
                }
                
            }
            catch (Exception)
            {

            }
            return pDF;
        }
        private string getDatosPDF(string path)
        {
            try
            {
                using (PdfReader reader = new PdfReader(path))
                using (PdfDocument pdfDoc = new PdfDocument(reader))
                {
                    StringBuilder text = new StringBuilder();
                    for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                    {
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string currentText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                        text.Append(currentText);
                    }
                    return text.ToString();
                }

            }
            catch { }
            return "";
        }
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);

        const uint WM_SETTEXT = 0x000C;
        const uint BM_CLICK = 0x00F5;
        static int SetFileUploadDialog(string filePath)
        {
            int band = 0;
            string key = "Abrir";
            CultureInfo currentUICulture = Thread.CurrentThread.CurrentUICulture;

            if (currentUICulture.TwoLetterISOLanguageName.Contains("en"))
                key = "Open";

            const int maxRetries = 10;
            int retryCount = 0;

            IntPtr hwndDialog = IntPtr.Zero;
            while (hwndDialog == IntPtr.Zero && retryCount < maxRetries)
            {
                hwndDialog = FindWindow("#32770", key); // Clase de ventana para el diálogo de archivo
                Thread.Sleep(500);
                retryCount++;
            }

            if (hwndDialog == IntPtr.Zero)
            {
                throw new Exception("No se pudo encontrar el cuadro de diálogo de archivo.");
            }

            IntPtr hwndFileName = FindWindowEx(hwndDialog, IntPtr.Zero, "ComboBoxEx32", null);
            hwndFileName = FindWindowEx(hwndFileName, IntPtr.Zero, "ComboBox", null);
            hwndFileName = FindWindowEx(hwndFileName, IntPtr.Zero, "Edit", null);

            IntPtr hwndOpenButton = FindWindowEx(hwndDialog, IntPtr.Zero, "Button", $"&{key}");

            if (hwndFileName == IntPtr.Zero || hwndOpenButton == IntPtr.Zero)
            {
                throw new Exception("No se pudieron encontrar los controles del cuadro de diálogo de archivo.");
            }

            SendMessage(hwndFileName, WM_SETTEXT, IntPtr.Zero, filePath);
            SendMessage(hwndOpenButton, BM_CLICK, IntPtr.Zero, null);
            band++;
            return band;
        }
        
    }
}
