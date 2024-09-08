using aspNetCore.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
//using DocumentFormat.OpenXml.Office.CustomUI;
//using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using aspNetCore.Helpers;

using aspNetCore.Models;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;
using Microsoft.AspNetCore.Routing.Template;
using System.IO;
using System.Reflection.Metadata;

using System.Threading.Tasks;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2010.ExcelAc;


namespace aspNetCore.Controllers
{
    public class HomeController : Controller
    {
        private IHostingEnvironment Environment;

        private readonly ILogger<HomeController> _logger;

        string docxMIMEType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

        //public HomeController(ILogger<HomeController> logger)
        //{
            
        //    _logger = logger;
            
        //}

        public HomeController(IHostingEnvironment _environment)
        {
            Environment = _environment;
        }




        public List<FileModel> getFiles()
        {
            string[] filepaths = Directory.GetFiles(Path.Combine(this.Environment.WebRootPath, "Files/"));
            List<FileModel> list = new List<FileModel>();
            foreach (string filepath in filepaths)
            {
                list.Add(new FileModel { FileName = Path.GetFileName(filepath) });
            }

            return list;

        }

        public IActionResult Index()
        {

            return View(getFiles());
        }

        
        public ActionResult ReadWordDocument(string filename)
        {

            var filePath = Path.Combine(this.Environment.WebRootPath, "Files/") + filename;

            //var filePath = @"\\WIN2022SRV\docs\AllDocs\eDocs.docx";

            try
            {
                // Open the Wordprocessing document for reading.
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true))
                {
                    wordDocument.ExtendedFilePropertiesPart.Properties.DocumentSecurity = new DocumentSecurity("0");
                    DocumentProtection dp = wordDocument.MainDocumentPart.DocumentSettingsPart
                        .Settings.ChildElements.First<DocumentProtection>();



                    // Get the document body.
                    Body body = wordDocument.MainDocumentPart.Document.Body;

                    // Loop through each paragraph in the body.
                    //foreach (Paragraph paragraph in body.Elements<Paragraph>())
                    //{
                    //    foreach (Run run in paragraph.Elements<Run>())
                    //    {
                    //        foreach (Text text in run.Elements<Text>())
                    //        {
                    //            // Print the text content of each paragraph.
                    //            Console.WriteLine(text.Text);
                    //        }
                    //    }
                    //}


                    //Gets all the headers
                    foreach (var headerPart in wordDocument.MainDocumentPart.HeaderParts)
                    {
                        //Gets the text in headers
                        foreach (var currentText in headerPart.RootElement.Descendants<Text>())
                        {

                            Console.WriteLine("header item: ..." + currentText.Text + "...");

                            //if (currentText.Text.Contains("Bytheg"))
                            if (currentText.Text.Contains("Empresa"))
                            {
                                Console.WriteLine("header tag found: ..." + currentText.Text + "...");
                                currentText.Text = currentText.Text.Replace("Empresa", "Bytheg");
                                //currentText.Text = currentText.Text.Replace("Bytheg", "Empresa");
                            }
                            else if (currentText.Text.Contains("Bytheg"))
                            {
                                Console.WriteLine("header tag found: ..." + currentText.Text + "...");
                                currentText.Text = currentText.Text.Replace("Bytheg", "Empresa");
                            }
                        }
                    }

                    wordDocument.Save();

                    Console.WriteLine("DP: " + dp);

                    ////Manage headers ------------
                    //foreach (HeaderPart hp in wordDocument.MainDocumentPart.HeaderParts)
                    //{
                    //    // add/modify header values
                    //    Header h = hp.Header;
                    //}

                    //foreach (FooterPart fp in wordDocument.MainDocumentPart.FooterParts)
                    //{
                    //    // add/modify footer values
                    //    Footer f = fp.Footer;

                    //}
                }
            }
            catch (Exception ex) { 
                Console.WriteLine("Exception:  " + ex.Message);
                TempData["ErrorMessage"] = ex.Message;
            }

            return Redirect("/");

        }


        [HttpGet("download")]
        public async Task<IActionResult> DownloadWordFile()
        {
            // Path to the Word file on the server
            var filePath = Path.Combine(this.Environment.WebRootPath, "Files/") + "SEN-FDO-03-001-1.docx";
            var fileName = "example.docx";
            var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);

            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }

        //Download Template file ok  --------------------------------  OK
        //[HttpGet("download")]
        public IActionResult OpenFile(string filename)
        {
            Console.WriteLine("DownloadFile: " + filename);
            var filePath = Path.Combine(this.Environment.WebRootPath, "Files/") + filename;
            
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound();
            }

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            var contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"; // MIME type for Word files
            return File(fileBytes, contentType, filename);
        }

        //Download blank file 
        public FileResult DownloadFileSream(string filename)
        {
            string path = Path.Combine(this.Environment.WebRootPath, "Files/") + filename;
            var stream = new MemoryStream();

            Console.WriteLine("DownloadFileStream: " + filename);

            //stream.Position = 0;
            stream.Seek(0, SeekOrigin.Begin);


            return File(stream, "application/msword", filename);
            //return File(stream, "application/msword", filename);
            //return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename);
        }
        
        //Download Template file ok  --------------------------------  OK
        public FileResult DownloadFileBytes(string filename)
        {
            string path = Path.Combine(this.Environment.WebRootPath, "Files/") + filename;
            byte[] bytes = System.IO.File.ReadAllBytes(path);

            Console.WriteLine("DownloadFileBytes: " + filename);

            return File(bytes, "application/msword", filename);
        }

        public FileResult DownloadFile_(string filename) {
            string path = Path.Combine(this.Environment.WebRootPath, "Files/") + filename;
            byte[] bytes = System.IO.File.ReadAllBytes(path);

            return File(bytes, "application/octet-stream", filename);
        }

        public ActionResult openDoc(string filename) {


            string path = Path.Combine(this.Environment.WebRootPath, "Files/") + filename;
            var stream = new MemoryStream();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
            {
                //MainDocumentPart mainPart = doc.AddMainDocumentPart();

                //new DocumentFormat.OpenXml.Wordprocessing.Document(new Body()).Save(mainPart);

                //Body body = mainPart.Document.Body;
                //body.Append(new Paragraph(
                //            new Run(
                //                new Text("Hello World!"))));

                //mainPart.Document.Save();

                //if you don't use the using you should close the WordprocessingDocument here
                //doc.Close();
            }
            //stream.Seek(0, SeekOrigin.Begin);
            //return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename);

            stream.Position = 0;
            return File(stream, "application/msword", filename);

        }


        public void openDoc_(string filename)
        {
            try
            {

                //FileStream fs = new FileStream(Server.MapPath(@"~\Content\documetfile.doc"), FileMode.OpenOrCreate, FileAccess.ReadWrite);
                //return File(fs, "documentfile.doc");

                string path = Path.Combine(this.Environment.WebRootPath, @"Files\") + filename;
                //string fileName = $@"Files\{file}";

                Console.WriteLine("path: " + path);

                //using (WordprocessingDocument doc
                //    = WordprocessingDocument.Open(path, true))
                //{
                //    //Body body = doc.MainDocumentPart.Document.Body;
                //    //doc.Close();
                //    TempData["AlertMessage"] = "Doc opened";
                //    Console.WriteLine("path: " + path);
                //    //return View();

                //}


                using (Stream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (WordprocessingDocument wpd = WordprocessingDocument.Open(stream, false))
                    {
                        Console.WriteLine("path: " + path);
                    }
                }

                //using (WordprocessingDocument wpd = WordprocessingDocument.Open(path, true))
                //{
                //    //Do stuff here
                //}

                //Stream stream = File.Open(strDoc, FileMode.Open);
                //// Open a WordprocessingDocument for read-only access based on a stream.
                //using (WordprocessingDocument wordDocument =
                //    WordprocessingDocument.Open(stream, false)) { }

            }
            catch(Exception ex) 
            {
                TempData["ErrorMessage"] = "Error opening file: " + ex.Message;
                //return View("Error", Error().ToString());

            }
        }
        //[HttpPost]
        //[Route("Home/CreateWordprocessingDocument")]
        public IActionResult CreateWordprocessingDocument(string filename)
        {
            Console.WriteLine("filename: " + filename);
            return View("Index");
        }


        private const string ToReplace = "to-replace";
        private const string ReplaceWith = "replace-with";

        [HttpPost]
        [Route("Home/CreateWordprocessingDocument_")]
        public MemoryStream CreateWordprocessingDocument_(string filename="default")
        {

            string path = Path.Combine(this.Environment.WebRootPath, @"Files\") + filename;
            Console.WriteLine("doc: " + @path);

            var stream = new MemoryStream();
            const WordprocessingDocumentType type = WordprocessingDocumentType.Document;

            using WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, type);
            MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
            mainDocumentPart.Document =
                new DocumentFormat.OpenXml.Wordprocessing.Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new Text(ToReplace))),
                        new Paragraph(
                            new Run(
                                new Text("to-")),
                            new Run(
                                new Text("replace")))));

            return stream;
        }

        private static void ReplaceText(MemoryStream stream)
        {
            using WordprocessingDocument doc = WordprocessingDocument.Open(stream, true);

            Body body = doc.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();

            foreach (Paragraph para in paras)
            {
                foreach (Run run in para.Elements<Run>())
                {
                    foreach (Text text in run.Elements<Text>())
                    {
                        if (text.Text.Contains(ToReplace))
                        {
                            text.Text = text.Text.Replace(ToReplace, ReplaceWith);
                            run.AppendChild(new Break());
                        }
                    }
                }
            }
        }

        public  void CreateWordDoc_msg___(string msg = "default")
        {
            string path = Path.Combine(this.Environment.WebRootPath, @"Files\") + msg;
            Console.WriteLine("doc: " + @path);

            var stream = new MemoryStream();
            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                fileStream.CopyTo(stream);


            using (var doc = WordprocessingDocument.Open(stream, true))
            {

                doc.ChangeDocumentType(WordprocessingDocumentType.Document); // change from template to document

                //var body = doc.MainDocumentPart.Document.Body;

                ////add some text 
                //Paragraph paraHeader = body.AppendChild(new Paragraph());
                //Run run = paraHeader.AppendChild(new Run());
                //run.AppendChild(new Text("This is body text"));

                //doc.Close();



                //return View("Index");
            }
            //return View("Index");

        }

        [HttpPost]
        public IActionResult CreateWordDoc_msg_(string msg)
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
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
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


        public void CreateWordDoc_msg__(string msg = "default")
        {
            string path = Path.Combine(this.Environment.WebRootPath, @"Files\") + msg;
            Console.WriteLine("doc: " + @path);
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

        public IActionResult addTable()
        {
            using (var stream = new MemoryStream())
            {
                using (var wordDoc = WordprocessingDocument.Create(stream,
                    WordprocessingDocumentType.Document, true))
                {
                    wordDoc.AddMainDocumentPart();
                    var doc = new DocumentFormat.OpenXml.Wordprocessing.Document();
                    var body = new Body();
                    var table = new Table();
                    var tableWidth = new TableWidth()
                    {
                        Width = "5000",
                        Type = TableWidthUnitValues.Pct
                    };
                    var borderColor = "FF8000";
                    var tableProperties = new TableProperties();
                    var tableBorders = new TableBorders();


                    var topBorder = new TopBorder();
                    topBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    topBorder.Color = borderColor;

                    var bottomBorder = new BottomBorder();
                    bottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    bottomBorder.Color = borderColor;

                    var rightBorder = new BottomBorder();
                    rightBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    rightBorder.Color = borderColor;

                    var leftBorder = new BottomBorder();
                    leftBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    leftBorder.Color = borderColor;

                    var insideHorizontalBorder = new BottomBorder();
                    insideHorizontalBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    insideHorizontalBorder.Color = borderColor;

                    var insideVerticalBorder = new BottomBorder();
                    insideVerticalBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    insideVerticalBorder.Color = borderColor;

                    tableBorders.AppendChild(topBorder);
                    tableBorders.AppendChild(bottomBorder);
                    tableBorders.AppendChild(rightBorder);
                    tableBorders.AppendChild(leftBorder);
                    tableBorders.AppendChild(insideHorizontalBorder);
                    tableBorders.AppendChild(insideVerticalBorder);

                    tableProperties.Append(tableWidth);
                    tableProperties.AppendChild(tableBorders);
                    table.AppendChild(tableProperties);

                    //--- Rows

                    var row1 = new TableRow();
                    var cell1 = new TableCell();
                    var paragraph1 = new Paragraph(new Run(new Text("Guetin")));

                    cell1.Append(paragraph1);
                    row1.Append(cell1);

                    //--cell2
                    var cell2 = new TableCell();
                    var paragraph2 = new Paragraph();
                    var run1 = new Run();
                    var runProperties = new RunProperties();
                    runProperties.Bold = new Bold();

                    run1.Append(runProperties);
                    run1.Append(new Text("400"));
                    paragraph2.Append(run1);
                    cell1.Append(paragraph2);
                    row1.Append(cell2);

                    //---
                    table.Append(row1);
                    var random = new Random();
                    for (int i = 1; i < 5; i++)
                    {
                        var row = new TableRow();
                        var cell3 = new TableCell();
                        var paragraph = new Paragraph(new Run
                            (new Text($"Employee {i}")));
                        cell3.Append(paragraph);

                        var cell4 = new TableCell();
                        var paragraph4 = new Paragraph();
                        var paragraphProperties1 = new ParagraphProperties();
                        paragraphProperties1.Justification = new Justification()
                        { Val = JustificationValues.Center };
                        paragraph4.Append(paragraphProperties1);
                        paragraph4.Append(new Run(new Text(random.Next(100, 500).ToString())));
                        cell4.Append(paragraph4);

                        row.Append(cell3);
                        row.Append(cell4);
                        table.Append(row);
                    }

                    body.Append(table);
                    doc.Append(body);
                    wordDoc.MainDocumentPart.Document = doc;
                    wordDoc.Clone();
                }


                return File(stream.ToArray(), docxMIMEType, "eDocsTable.docx");
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
