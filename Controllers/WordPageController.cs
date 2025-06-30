using IronWord.Models;
using IronWord;
using Microsoft.AspNetCore.Mvc;

namespace Ext_IronWord_Project.Controllers
{
    public class WordPageController : Controller
    {
        private readonly IWebHostEnvironment _env;

        public WordPageController(IWebHostEnvironment env)
        {
            _env = env;
        }

        [HttpGet]
        public IActionResult AddParagraphs()
        {
            return View();
        }

        [HttpPost]
        public IActionResult AddParagraphs(IFormFile uploadedFile)
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (uploadedFile == null || uploadedFile.Length == 0)
            {
                ViewBag.Message = "Please upload a valid Word file.";
                return View();
            }

            // Save uploaded file temporarily
            string uploadsPath = Path.Combine(_env.WebRootPath, "temp");
            if (!Directory.Exists(uploadsPath))
                Directory.CreateDirectory(uploadsPath);

            string inputFile = Path.Combine(uploadsPath, uploadedFile.FileName);
            using (var stream = new FileStream(inputFile, FileMode.Create))
            {
                uploadedFile.CopyTo(stream);
            }

            // Load document
            WordDocument doc = new WordDocument(inputFile);

            // Create styled text
            TextContent intro = new TextContent("This is an example paragraph with italic and bold styling.\n");

            TextStyle italicStyle = new TextStyle { IsItalic = true };
            TextContent italic = new TextContent("Italic example sentence. ");
            italic.Style = italicStyle;

            TextStyle boldStyle = new TextStyle { IsBold = true };
            TextContent bold = new TextContent("Bold example sentence.");
            bold.Style = boldStyle;

            // Build paragraph
            Paragraph para = new Paragraph();
            para.AddText(intro);
            para.AddText(italic);
            para.AddText(bold);

            // Add to document
            doc.AddParagraph(para);

            // Save edited document
            string newFileName = "Edited_" + Path.GetFileName(uploadedFile.FileName);
            string outputFile = Path.Combine(uploadsPath, newFileName);
            doc.SaveAs(outputFile);

            // Return downloadable file
            var fileBytes = System.IO.File.ReadAllBytes(outputFile);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", newFileName);
        }

        [HttpGet]
        public IActionResult AddImage()
        {
            return View();
        }

        [HttpPost]
        public IActionResult AddImage(IFormFile uploadedImage)
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (uploadedImage == null || uploadedImage.Length == 0)
            {
                ViewBag.Message = "Please upload a valid image file.";
                return View();
            }

            // Save uploaded image to wwwroot/temp
            string tempDir = Path.Combine(_env.WebRootPath, "temp");
            if (!Directory.Exists(tempDir))
                Directory.CreateDirectory(tempDir);

            string imagePath = Path.Combine(tempDir, uploadedImage.FileName);
            using (var stream = new FileStream(imagePath, FileMode.Create))
            {
                uploadedImage.CopyTo(stream);
            }

            // Create a new Word document
            WordDocument doc = new WordDocument();

            // Create and configure image
            ImageContent image = new ImageContent(imagePath)
            {
                Width = 200,
                Height = 200
            };

            // Add image to a paragraph
            Paragraph para = new Paragraph();
            para.AddImage(image);
            doc.AddParagraph(para);

            // Save Word document
            string fileName = "ImageDoc_" + Path.GetFileNameWithoutExtension(uploadedImage.FileName) + ".docx";
            string outputPath = Path.Combine(tempDir, fileName);
            doc.SaveAs(outputPath);

            byte[] fileBytes = System.IO.File.ReadAllBytes(outputPath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }


        [HttpGet]
        public IActionResult AddList()
        {
            return View();
        }

        [HttpPost]
        public IActionResult AddList(List<string> ListItems)
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (ListItems == null || ListItems.Count == 0)
            {
                ViewBag.Message = "Please provide at least one list item.";
                return View();
            }

            // Create Word document
            WordDocument doc = new WordDocument();

            // Create multi-level text list
            MultiLevelTextList textList = new MultiLevelTextList();

            foreach (var item in ListItems)
            {
                TextContent text = new TextContent(item);
                Paragraph paragraph = new Paragraph();
                paragraph.AddText(text);
                ListItem listItem = new ListItem(paragraph);
                textList.AddItem(listItem);
            }

            // Add list to document
            doc.AddMultiLevelTextList(textList);

            // Save file to temp and return as download
            string tempDir = Path.Combine(_env.WebRootPath, "temp");
            if (!Directory.Exists(tempDir))
                Directory.CreateDirectory(tempDir);

            string fileName = "WordList_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".docx";
            string outputPath = Path.Combine(tempDir, fileName);
            doc.SaveAs(outputPath);

            byte[] fileBytes = System.IO.File.ReadAllBytes(outputPath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }

        public IActionResult AddTable()
        {
            return View();
        }
        [HttpPost]
        public IActionResult GenerateTableDoc()
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            try
            {
                // Create table
                Table table = new Table(5, 3);
                table.Zebra = new ZebraColor("FFFFFF", "DDDDDD");

                table[0, 0] = new TableCell(new TextContent("Number"));
                table[0, 1] = new TableCell(new TextContent("First Name"));
                table[0, 2] = new TableCell(new TextContent("Last Name"));

                for (int i = 1; i < table.Rows.Count; i++)
                {
                    table[i, 0] = new TableCell(new TextContent(i.ToString()));
                    table[i, 1] = new TableCell(new TextContent($"FirstName{i}"));
                    table[i, 2] = new TableCell(new TextContent($"LastName{i}"));
                }

                // Generate document
                WordDocument doc = new WordDocument(table);

                // Save path
                string tempPath = Path.Combine(Path.GetTempPath(), "TableDoc.docx");
                doc.SaveAs(tempPath);

                // Ensure file exists before returning
                if (!System.IO.File.Exists(tempPath))
                {
                    return Content("Error: Word document was not created.");
                }

                byte[] fileBytes = System.IO.File.ReadAllBytes(tempPath);
                System.IO.File.Delete(tempPath);

                return File(fileBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "TableDocument.docx");
            }
            catch (Exception ex)
            {
                return Content($"Error occurred: {ex.Message}<br><br>{ex.StackTrace}", "text/html");
            }
        }

    }
    
}
