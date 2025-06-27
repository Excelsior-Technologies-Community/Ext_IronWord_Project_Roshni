using Ext_IronWord_Project.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using System.IO;
using System.Text;
using IronWord;
using IronWord.Models;
using Microsoft.AspNetCore.Http;
using System;
using DocumentFormat.OpenXml.Packaging;
using SixLabors.Fonts;


namespace Ext_IronWord_Project.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _env;

        public HomeController(IWebHostEnvironment env)
        {
            _env = env;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult CreateEmptyWord()
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            var doc = new WordDocument();

            
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
            doc.SaveAs(tempPath);

            
            byte[] bytes = System.IO.File.ReadAllBytes(tempPath);

            
            System.IO.File.Delete(tempPath);

            
            return File(bytes,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        "EmptyDocument.docx");
        }


        [HttpGet]
        public IActionResult CreateDocx()
        {
            return View();
        }

        [HttpPost]
        public IActionResult CreateDocx(string docText)
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (string.IsNullOrWhiteSpace(docText))
            {
                ViewBag.Message = "Please enter some text.";
                return View();
            }

            try
            {
                // Create a paragraph and add text via AddChild()
                var paragraph = new Paragraph();
                paragraph.AddChild(new TextContent(docText));

                // Create a new document and include the paragraph
                var doc = new WordDocument(paragraph);

                // Save to a temp .docx file
                string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
                doc.SaveAs(tempFile);

                // Read and return file bytes
                byte[] docBytes = System.IO.File.ReadAllBytes(tempFile);

                // Clean up
                System.IO.File.Delete(tempFile);

                return File(
                    docBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "GeneratedDocument.docx"
                );
            }
            catch (Exception ex)
            {
                ViewBag.Message = $"Error generating DOCX: {ex.Message}";
                return View();
            }
        }


        [HttpGet]
        public IActionResult EditDocx()
        {
            return View();
        }

        [HttpPost]
        public IActionResult EditDocx(IFormFile uploadedFile, string oldText, string newText)
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (uploadedFile != null && uploadedFile.Length > 0)
            {
                // Save uploaded file temporarily
                string uploadsFolder = Path.Combine(_env.WebRootPath, "temp");
                if (!Directory.Exists(uploadsFolder))
                    Directory.CreateDirectory(uploadsFolder);

                string originalPath = Path.Combine(uploadsFolder, uploadedFile.FileName);
                using (var fileStream = new FileStream(originalPath, FileMode.Create))
                {
                    uploadedFile.CopyTo(fileStream);
                }

                // Load Word document using IronWord
                var doc = new WordDocument(originalPath);

                // Replace text in all paragraphs
                foreach (var para in doc.Paragraphs)
                {
                    para.ReplaceText(oldText, newText);
                }

                // Save the edited document
                string editedFileName = "Edited_" + Path.GetFileName(uploadedFile.FileName);
                string editedPath = Path.Combine(uploadsFolder, editedFileName);
                doc.SaveAs(editedPath);

                ViewBag.DownloadLink = "/temp/" + editedFileName;
                return View();
            }

            ViewBag.Message = "Please upload a valid Word (.docx) file.";
            return View();
        }

        [HttpGet]
        public IActionResult LogTree()
        {
            return View();
        }

        [HttpPost]
        public IActionResult LogTree(IFormFile uploadedFile)
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (uploadedFile == null || uploadedFile.Length == 0)
            {
                ViewBag.Message = "Please upload a valid Word (.docx) file.";
                return View();
            }

            // Save uploaded file
            string uploadsFolder = Path.Combine(_env.WebRootPath, "temp");
            if (!Directory.Exists(uploadsFolder))
                Directory.CreateDirectory(uploadsFolder);

            string filePath = Path.Combine(uploadsFolder, uploadedFile.FileName);
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                uploadedFile.CopyTo(stream);
            }

            var output = new StringBuilder();
            var docx = new WordDocument(filePath);

            // Manual traversal since LogObjectTree() cannot write to output
            output.AppendLine("📄 Document Children:");
            foreach (var child in docx.Children)
            {
                output.AppendLine($"• {child.GetType().Name}");
            }

            var firstTable = docx.Children.OfType<Table>().FirstOrDefault();
            if (firstTable != null)
            {
                output.AppendLine("\n📊 First Table Rows:");
                foreach (var row in firstTable.Children.OfType<TableRow>())
                {
                    output.AppendLine("  Row:");
                    foreach (var cell in row.Children.OfType<TableCell>())
                    {
                        var text = string.Join(" ", cell.Children
                                                         .OfType<Paragraph>()
                                                         .SelectMany(p => p.Children)
                                                         .OfType<Run>()
                                                         .Select(r => r.ToString()));

                        output.AppendLine($"    Cell: {text}");
                    }

                }
            }
            else
            {
                output.AppendLine("\n❌ No tables found.");
            }

            ViewBag.ObjectTree = output.ToString();
            return View();
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
