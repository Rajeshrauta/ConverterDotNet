using Microsoft.AspNetCore.Mvc;
using Converter.Services;
using SautinSoft.Document;
using PdfSharp.Pdf.IO;
using Section = SautinSoft.Document.Section;
using PdfDocument = PdfSharp.Pdf.PdfDocument;
using PdfReader = PdfSharp.Pdf.IO.PdfReader;
using PdfPage = PdfSharp.Pdf.PdfPage;

namespace PdfToWordConverter.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ConvertController : ControllerBase
    {
        private readonly PdfToWordService _pdfToWordService;

        public ConvertController(PdfToWordService pdfToWordService)
        {
            _pdfToWordService = pdfToWordService;
        }

        //SautinSoft
        [HttpPost("pdf-to-word")]
        public IActionResult ConvertPdfToWord(IFormFile pdfFile)
        {
            if (pdfFile == null || pdfFile.Length == 0)
                return BadRequest("No file uploaded.");

            if (Path.GetExtension(pdfFile.FileName).ToLower() != ".pdf")
                return BadRequest("Please upload a PDF file.");

            try
            {

                var originalFileName = Path.GetFileNameWithoutExtension(pdfFile.FileName);
                var newFileName = $"{originalFileName}.docx";

                using (var memoryStream = new MemoryStream())
                {
                    pdfFile.CopyTo(memoryStream);
                    var pdfBytes = memoryStream.ToArray();
                    pdfBytes = _pdfToWordService.AddBlankPageToPdf(pdfBytes);
                    var wordBytes = _pdfToWordService.ConvertPdfToWord(pdfBytes);
                    if (wordBytes == null)
                        return StatusCode(500, "Conversion failed.");

                    return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", newFileName);
                }
            }
            catch (Exception ex)
            {
                // Log the exception for debugging purposes
                Console.WriteLine($"Error converting PDF to Word: {ex.Message}");
                return StatusCode(500, "An error occurred during conversion.");
            }
        }

        [HttpPost("word-to-pdf")]
        public async Task<IActionResult> ConvertDocxToPdf(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            try
            {
                string originalFileName = Path.GetFileNameWithoutExtension(file.FileName);
                if (string.IsNullOrEmpty(originalFileName))
                    originalFileName = "converted";

                // Load the DOCX file into the DocumentCore
                DocumentCore dc;
                using (var ms = new MemoryStream())
                {
                    await file.CopyToAsync(ms);
                    ms.Position = 0;
                    dc = DocumentCore.Load(ms, new DocxLoadOptions());
                }

                Section blankSection = new Section(dc);
                Paragraph blankParagraph = new Paragraph(dc);
                blankSection.Blocks.Add(blankParagraph);
                dc.Sections.Add(blankSection);


                byte[] pdfBytes;

                // Convert the document to PDF
                using (var msPdf = new MemoryStream())
                {
                    dc.Save(msPdf, new PdfSaveOptions());
                    //msPdf.Position = 0;
                    //return File(msPdf.ToArray(), "application/pdf", "converted.pdf");
                    pdfBytes = msPdf.ToArray();
                }

                // Load the PDF into PdfDocument (PdfSharp)
                using (var inputPdfStream = new MemoryStream(pdfBytes))
                using (var outputPdfStream = new MemoryStream())
                {
                    PdfDocument pdfDocument = PdfReader.Open(inputPdfStream, PdfDocumentOpenMode.Modify);

                    // Remove the last page
                    if (pdfDocument.PageCount > 0)
                    {
                        pdfDocument.Pages.RemoveAt(pdfDocument.PageCount - 1);
                    }

                    // Save the modified PDF to a MemoryStream
                    pdfDocument.Save(outputPdfStream);
                    outputPdfStream.Position = 0;

                    return File(outputPdfStream.ToArray(), "application/pdf", $"{originalFileName}.pdf");

                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }


        [HttpPost("rotate-pdf")]
        public IActionResult Rotate(IFormFile file, string rotationDirection)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("Invalid file.");
            }

            if (string.IsNullOrEmpty(rotationDirection) || (rotationDirection != "left" && rotationDirection != "right"))
            {
                return BadRequest("Invalid rotation direction.");
            }

            try
            {
                string originalFileName = Path.GetFileNameWithoutExtension(file.FileName);
                if (string.IsNullOrEmpty(originalFileName))
                    originalFileName = "Rotated";

                using (MemoryStream outputStream = new MemoryStream())
                {
                    using (PdfDocument originalDocument = PdfReader.Open(file.OpenReadStream(), PdfDocumentOpenMode.Import))
                    {
                        PdfDocument rotatedDocument = new PdfDocument();

                        foreach (PdfPage page in originalDocument.Pages)
                        {
                            PdfPage newPage = rotatedDocument.AddPage(page);
                            newPage.Rotate = (page.Rotate + (rotationDirection == "left" ? 270 : 90)) % 360;
                        }

                        rotatedDocument.Save(outputStream);
                    }

                    outputStream.Seek(0, SeekOrigin.Begin);

                    byte[] fileBytes = outputStream.ToArray();

                    var fileContentResult = new FileContentResult(fileBytes, "application/pdf")
                    {
                        FileDownloadName = $"{originalFileName}_{rotationDirection}_Rotated.pdf"
                    };

                    return fileContentResult;
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"An error occurred: {ex.Message}");
            }
        }


        [HttpPost("split")]
        public IActionResult Split(IFormFile sourceFile, string pageRange, string newFileName = "")
        {
            if (sourceFile == null || sourceFile.Length == 0)
            {
                return BadRequest("Invalid file.");
            }

            if (string.IsNullOrEmpty(pageRange))
            {
                return BadRequest("Page range not provided.");
            }

            try
            {
                using (PdfDocument originalDocument = PdfReader.Open(sourceFile.OpenReadStream(), PdfDocumentOpenMode.Import))
                {
                    var pageNumbers = ParsePageNumbers(pageRange);
                    if (pageNumbers == null || pageNumbers.Count == 0)
                    {
                        return BadRequest("Invalid or empty page range format.");
                    }

                    PdfDocument newDocument = new PdfDocument();

                    foreach (var pageNumber in pageNumbers)
                    {
                        if (pageNumber >= 1 && pageNumber <= originalDocument.PageCount)
                        {
                            newDocument.AddPage(originalDocument.Pages[pageNumber - 1]);
                        }
                    }

                    if (newDocument.PageCount > 0)
                    {
                        string newFilePath;
                        if (string.IsNullOrEmpty(newFileName))
                        {
                            string sourceFileName = Path.GetFileNameWithoutExtension(sourceFile.FileName);
                            newFilePath = Path.Combine(Path.GetTempPath(), $"{sourceFileName}_split.pdf");
                        }
                        else
                        {
                            if (Path.GetExtension(newFileName) != ".pdf")
                            {
                                newFileName += ".pdf";
                            }
                            newFilePath = Path.Combine(Path.GetTempPath(), newFileName);
                        }

                        newDocument.Save(newFilePath);

                        byte[] fileBytes = System.IO.File.ReadAllBytes(newFilePath);

                        // Ensure the temporary file is deleted after reading it
                        System.IO.File.Delete(newFilePath);

                        var fileContentResult = new FileContentResult(fileBytes, "application/pdf")
                        {
                            FileDownloadName = Path.GetFileName(newFilePath)
                        };

                        return fileContentResult;
                    }
                    else
                    {
                        return BadRequest("No pages found in the specified range.");
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"An error occurred: {ex.Message}");
            }
        }

        private List<int> ParsePageNumbers(string pageRange)
        {
            List<int> pageNumbers = new List<int>();

            string[] ranges = pageRange.Split(',');
            foreach (var range in ranges)
            {
                string[] parts = range.Trim().Split('-');
                if (parts.Length == 1)
                {
                    // Single page number
                    if (int.TryParse(parts[0], out int pageNumber))
                    {
                        pageNumbers.Add(pageNumber);
                    }
                    else
                    {
                        return null; // Invalid single page number
                    }
                }
                else if (parts.Length == 2)
                {
                    // Range of page numbers
                    if (int.TryParse(parts[0], out int startPage) && int.TryParse(parts[1], out int endPage))
                    {
                        for (int i = startPage; i <= endPage; i++)
                        {
                            pageNumbers.Add(i);
                        }
                    }
                    else
                    {
                        return null; // Invalid range format
                    }
                }
                else
                {
                    return null; // Invalid format
                }
            }

            return pageNumbers;
        }
    }
}
