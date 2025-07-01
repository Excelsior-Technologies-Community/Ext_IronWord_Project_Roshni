using Microsoft.AspNetCore.Mvc;
using Ext_IronWord_Project.Models;
using System.IO;
using System.Drawing;
using Xceed.Words.NET;
using Xceed.Document.NET;
using Xceed.Drawing;
using IronWord;
using IronWord.Models;
using DrawingColor = System.Drawing.Color;
using DrawingColorTranslator = System.Drawing.ColorTranslator;
using XceedColor = Xceed.Drawing.Color;
using IronColor = IronWord.Models.Color;
using ExtGradient = Ext_IronWord_Project.Models.Gradient;
using SixLabors.Fonts;
using System.Collections.Generic; 
using WordColor = IronWord.Models.Color;
using IronGradient = IronWord.Models.Gradient;
using SkiaSharp;


namespace Ext_IronWord_Project.Controllers
{
    public class DocumentController : Controller
    {
        public IActionResult AddTextStyle()
        {
            return View();
        }


        [HttpPost]
        public IActionResult GenerateTextDoc(string Text, int FontSize, string Bold, string Italic)
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            try
            {
                string filePath = Path.Combine(Path.GetTempPath(), "StyledText.docx");

                using (var doc = DocX.Create(filePath))
                {
                    var p = doc.InsertParagraph(Text)
                               .FontSize(FontSize);

                    if (!string.IsNullOrEmpty(Bold))
                        p.Bold();

                    if (!string.IsNullOrEmpty(Italic))
                        p.Italic();

                    doc.Save();
                }

                byte[] bytes = System.IO.File.ReadAllBytes(filePath);
                System.IO.File.Delete(filePath);

                return File(bytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "StyledText.docx");
            }
            catch (Exception ex)
            {
                return Content($"<b>Error:</b> {ex.Message}<br><br>{ex.StackTrace}", "text/html");
            }
        }

        [HttpGet]
        public IActionResult AddGlowEffect()
        {
            return View();
        }

        [HttpPost]
        public IActionResult GenerateGlowDoc(string inputText)
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            try
            {
                
                // Create Word Document
                WordDocument doc = new WordDocument();

                // Apply glow effect
                TextStyle textStyle = new TextStyle
                {
                    TextEffect = new TextEffect
                    {
                        GlowEffect = new Glow
                        {
                            GlowColor = IronWord.Models.Color.Orange,
                            GlowRadius = 10
                        }
                    }
                };

                // Add styled text
                doc.AddText(inputText).Style = textStyle;

                // Save document to temporary file
                string tempPath = Path.Combine(Path.GetTempPath(), "GlowText.docx");
                doc.SaveAs(tempPath);  // SaveAs(string path) — CORRECT

                // Read file and return as download
                byte[] fileBytes = System.IO.File.ReadAllBytes(tempPath);
                System.IO.File.Delete(tempPath); // Optional: cleanup

                return File(fileBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "GlowText.docx");
            }
            catch (Exception ex)
            {
                return Content($"<b>Error:</b> {ex.Message}<br><br>{ex.StackTrace}", "text/html");
            }
        }

        public IActionResult AddShadowEffect()
        {
            return View();
        }

        [HttpPost]
        public IActionResult GenerateShadowDoc(string inputText)
        {
            try
            {
                IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

                // Create a new Word document
                WordDocument doc = new WordDocument();

                // Configure text style with shadow
                TextStyle textStyle = new TextStyle
                {
                    TextEffect = new TextEffect
                    {
                        ShadowEffect = Shadow.OuterShadow1
                    }
                };

                // Add text with the style
                doc.AddText(inputText).Style = textStyle;

                // Save document to temp file
                string tempPath = Path.Combine(Path.GetTempPath(), "ShadowText.docx");
                doc.SaveAs(tempPath);

                // Read and return file
                byte[] fileBytes = System.IO.File.ReadAllBytes(tempPath);
                System.IO.File.Delete(tempPath);

                return File(fileBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "ShadowText.docx");
            }
            catch (Exception ex)
            {
                return Content($"<b>Error:</b> {ex.Message}<br><br>{ex.StackTrace}", "text/html");
            }
        }

        public IActionResult AddTextOutline()
        {
            return View();
        }

        [HttpPost]
        public IActionResult GenerateTextOutlineDoc(string inputText)
        {
            try
            {
                IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

                // Create a new Word document
                WordDocument doc = new WordDocument();

                // Configure text style with Text Outline effect
                TextStyle textStyle = new TextStyle
                {
                    TextEffect = new TextEffect
                    {
                        TextOutlineEffect = TextOutlineEffect.DefaultEffect
                    }
                };

                // Add text with the style
                doc.AddText(inputText).Style = textStyle;

                // Save document to a temporary path
                string filePath = Path.Combine(Path.GetTempPath(), "TextOutlineEffect.docx");
                doc.SaveAs(filePath);

                // Read file and return as download
                byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
                System.IO.File.Delete(filePath);

                return File(fileBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "TextOutlineEffect.docx");
            }
            catch (Exception ex)
            {
                return Content($"<b>Error:</b> {ex.Message}<br><br>{ex.StackTrace}", "text/html");
            }
        }
        [HttpGet]
        public IActionResult AddReflection()
        {
            return View();
        }
        [HttpPost]
        public IActionResult GenerateReflectionWord(string userText)
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (string.IsNullOrWhiteSpace(userText))
            {
                TempData["Error"] = "Please enter some text.";
                return RedirectToAction("AddReflection");
            }

            var doc = new WordDocument();

            // Create paragraph for original text
            var originalPara = new IronWord.Models.Paragraph();
            originalPara.AddText(userText).Style = new TextStyle
            {
                Color = new IronWord.Models.Color(0, 0, 0) // Black
            };
            doc.AddParagraph(originalPara);

            // Spacer
            var spacer = new IronWord.Models.Paragraph();
            spacer.AddText(" ");
            doc.AddParagraph(spacer);

            // Create reflection-style paragraph
            var reflectionPara = new IronWord.Models.Paragraph();
            reflectionPara.AddText(userText).Style = new TextStyle
            {
                Color = new IronWord.Models.Color(180, 180, 180) // Light gray
            };
            doc.AddParagraph(reflectionPara);

            // Save and return file
            string filePath = Path.Combine(Path.GetTempPath(), "ReflectionEffect.docx");
            doc.SaveAs(filePath);

            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
            System.IO.File.Delete(filePath);

            return File(fileBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "ReflectionEffect.docx");
        }

        public IActionResult AddGradient()
        {
            return View();
        }

        [HttpPost]
        public IActionResult GenerateGradientWord(string inputText, string startColor, string endColor)
        {
            IronWord.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (string.IsNullOrWhiteSpace(inputText))
                inputText = "Gradient Text Example";

            // Convert HTML hex to RGB
            var start = System.Drawing.ColorTranslator.FromHtml(startColor);
            var end = System.Drawing.ColorTranslator.FromHtml(endColor);

            WordDocument doc = new WordDocument();
            IronWord.Models.Paragraph para = new IronWord.Models.Paragraph();

            int len = inputText.Length;
            for (int i = 0; i < len; i++)
            {
                float ratio = (float)i / Math.Max(1, len - 1);

                byte r = (byte)(start.R + (end.R - start.R) * ratio);
                byte g = (byte)(start.G + (end.G - start.G) * ratio);
                byte b = (byte)(start.B + (end.B - start.B) * ratio);

                IronWord.Models.Color interpolatedColor = new IronWord.Models.Color(r, g, b);

                para.AddText(inputText[i].ToString()).Style = new TextStyle
                {
                    Color = interpolatedColor,
                    //FontSize = 20,
                    //Bold = true
                };
            }

            doc.AddParagraph(para);

            string filePath = Path.Combine(Path.GetTempPath(), "GradientText.docx");
            doc.SaveAs(filePath);

            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
            System.IO.File.Delete(filePath);

            return File(fileBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "GradientText.docx");
        }

    }
}
