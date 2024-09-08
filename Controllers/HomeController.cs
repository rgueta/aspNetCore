using aspNetCore.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using aspNetCore.Helpers;

namespace aspNetCore.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        string docxMIMEType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";


        public HomeController(ILogger<HomeController> logger)
        {
            
            _logger = logger;
            
        }

        public IActionResult Index()
        {
            return View();
        }

        
        public IActionResult CreateWordDoc_msg(string msg)
        {
            var fecha = DateTime.Now.ToString("yyyy.MM.dd.hh.mm.ss");
            Console.WriteLine(DateTime.Now.ToString("yyyy.MM.dd.hh.mm.ss"));
            //Debug.WriteLine("method called");
            //TempData["AlertMessage"] = msg;

            string fileName = @"\\WIN2022SRV\docs\AllDocs\eDocs\" + $"eDocs.{fecha}.docx";
            Console.WriteLine("Filename: " + fileName);
            using (var stream = new MemoryStream())
            {

                using (WordprocessingDocument doc = WordprocessingDocument.Create(fileName, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    // Add a main document part. 
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());

                    // String msg contains the text, "Hello, Word!"
                    run.AppendChild(new Text(msg));
                }
            }

            //return View(Index);
            //return new ContentResult { Content = msg };

            return View("Index");
        }

        public IActionResult createFromTemplate()
        {
            try
            {
                var fecha = DateTime.Now.ToString("yyyy.MM.dd.hh.mm.ss");
                var fileName = "eDocs.docx";
                var remotePath = @"\\WIN2022SRV\docs\AllDocs\";
                var filePath = $@"\\WIN2022SRV\docs\template\{fileName}";
                var name = "eDocs";
                var country = "Mex";

                using (var wordDoc = WordprocessingDocument.Open(filePath, true))
                {
                    var body = wordDoc.MainDocumentPart.Document.Body;
                    var fields = body.Descendants<FormFieldData>();
                    foreach (var field in fields)
                    {
                        var formField = (FormFieldName)field.FirstChild;
                        var fieldName = formField.Val.InnerText;

                        switch (fieldName)
                        {
                            case "Name":
                                UpdateFormField(field, name);
                                break;
                            case "Country":
                                UpdateFormField(field, country);
                                break;
                            default:
                                break;
                        }
                    }

                    TempData["AlertMessage"] = "Document created by template";
                    wordDoc.Clone();
                }


                
                Console.WriteLine("filePath: " + filePath);
                Console.WriteLine("remote: " + remotePath);

                return File(System.IO.File.ReadAllBytes(filePath), docxMIMEType, $"{name}.{fecha}.docx");
                //return File(System.IO.File.ReadAllBytes(@"c:/Users/Gueta/Downloads/eDocs.docx"), docxMIMEType, $"{name}.{fecha}.docx");
            }
            catch
            {
                return View("Error", Error().ToString());
            }
        }

        private void UpdateFormField(FormFieldData field, string value)
        {
            var text = field.Descendants<TextInput>().First();
            WordHelpers.SetFormFieldValue(text, value);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
