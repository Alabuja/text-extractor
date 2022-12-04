using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using TextExtractor.Models;
using IronPdf;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Net.Http.Headers;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using TextExtractor.Helpers;

namespace TextExtractor.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]

        public async Task<IActionResult> Index(string searchedWord)
        {
            try
            {
                var formFile = Request.Form.Files[0];

                if (formFile == null) // Check if the file is empty
                {
                    ViewBag.ErrorMessage = "File needs to be available";
                    return View();
                }

                var extension = System.IO.Path.GetExtension(formFile.FileName);

                if (extension.ToLower() != ".pdf") // Verify file extension to allow only pdf files.
                {
                    ViewBag.ErrorMessage = "Invalid file extension. File Extenison not supported.";
                    return View();
                }

                if (string.IsNullOrEmpty(searchedWord)) // check if the word to be search is null or empty.
                {
                    ViewBag.ErrorMessage = "Word to search by cannot be null";
                    return View();
                }

                var searchedList = searchedWord.Trim().Split(';').Select(x => x).ToList();

                string response = string.Empty;


                var result  = await UploadFile(formFile); // Upload the pdf file to a directory

                var filePath = result.path;

                if (result.Item1)
                {
                    if (!string.IsNullOrEmpty(filePath))
                    {
                        var pdfDocument = PdfDocument.FromFile($"UploadedFiles\\{filePath}"); // Get pdf from the directory

                        string text = pdfDocument.ExtractAllText(); // Extracts the text from the pdf uploaded.

                        var sentenceResponse = CheckIfText(text, searchedList.Distinct().ToList()); // Check of words exist in the pdf

                        if (System.IO.File.Exists($"UploadedFiles\\{filePath}"))
                        {
                            // If file found, delete it    
                            System.IO.File.Delete($"UploadedFiles\\{filePath}");
                        }

                        return GenerateExcel(sentenceResponse); // Export file to Excel with words and places the words were found


                    }
                    else
                    {
                        return View();
                    }

                }
                else
                {
                    return View();
                }
                
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = ex.Message;
                return View();
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        private async Task<(bool, string path)> UploadFile(IFormFile file)
        {
            string path = string.Empty;
            string fileName = string.Empty;
            try
            {
                if (file.Length > 0)
                {
                    var currentFileName = file.FileName;
                    path = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "UploadedFiles"));
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path); // Create the directory 'UploadedFiles'
                    }
                    using (var fileStream = new FileStream(Path.Combine(path, file.FileName), FileMode.Create))
                    {
                        await file.CopyToAsync(fileStream); // Copies file to the directory created above
                    }

                    string[] files = Directory.GetFiles(path);
                    foreach (string currentFile in files)
                    {
                        if(currentFileName == Path.GetFileName(currentFile))
                        {
                            fileName = currentFileName;
                        }
                    }

                    return (true, fileName);
                }
                else
                {
                    return (false, path);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private static Dictionary<string, List<string>> CheckIfText(string text, List<string> searchedWords)
        {
            List<string> result = new();
            Dictionary<string, List<string>> dictionary = new();

            if (!string.IsNullOrEmpty(text))
            {
                var sentences = text.Split(new[] { ". " }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var searchedWord in searchedWords)
                {
                    var matches = from sentence in sentences where sentence.ToLower().Contains(searchedWord.ToLower()) select sentence;

                    dictionary.Add(searchedWord, matches.ToList());
                }

            }

            return dictionary;
        }

        private ActionResult GenerateExcel(Dictionary<string, List<string>> keyValuePairs)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFFont myFont = (HSSFFont)workbook.CreateFont();
            myFont.FontHeightInPoints = 11;
            myFont.FontName = "Tahoma"; // Font type to be used is Tahoma


            // Defining a border
            HSSFCellStyle borderedCellStyle = (HSSFCellStyle)workbook.CreateCellStyle();
            borderedCellStyle.SetFont(myFont);
            borderedCellStyle.BorderLeft = BorderStyle.Medium;
            borderedCellStyle.BorderTop = BorderStyle.Medium;
            borderedCellStyle.BorderRight = BorderStyle.Medium;
            borderedCellStyle.BorderBottom = BorderStyle.Medium;
            borderedCellStyle.VerticalAlignment = VerticalAlignment.Center;

            ISheet Sheet = workbook.CreateSheet("Report"); // Name of sheet would be 'Report'
            //Creat The Headers of the excel
            IRow HeaderRow = Sheet.CreateRow(0);

            CreateCell(HeaderRow, 0, "Search Word", borderedCellStyle);
            CreateCell(HeaderRow, 1, "Sentences", borderedCellStyle);

            int RowIndex = 1;

            //Iteration through some collection
            //Creating the CurrentDataRow

            if (keyValuePairs != null)
            {
                foreach(var item in keyValuePairs)
                {
                    IRow CurrentRow = Sheet.CreateRow(RowIndex);
                    CreateCell(CurrentRow, 0, item.Key, borderedCellStyle);
                    // This will be used to calculate the merge area
                    var sentences = item.Value;

                    if (sentences != null && sentences.Count > 0)
                    {
                        int NumberOfSentences = sentences.Count;
                        if (NumberOfSentences > 1)
                        {
                            int MergeIndex = (NumberOfSentences - 1) + RowIndex;

                            //Merging Cells
                            NPOI.SS.Util.CellRangeAddress MergedBatch = new NPOI.SS.Util.CellRangeAddress(RowIndex, MergeIndex, 0, 0);
                            Sheet.AddMergedRegion(MergedBatch);
                        }
                        int i = 0;
                        // Iterate through cub collection
                        foreach (var sentence in sentences)
                        {
                            if (i > 0)
                                CurrentRow = Sheet.CreateRow(RowIndex);
                            CreateCell(CurrentRow, 1, sentence, borderedCellStyle);
                            RowIndex++;
                            i++;
                        }
                        _ = NumberOfSentences >= 1 ? RowIndex : RowIndex + 1;
                    }
                    // Auto sized all the affected columns
                    int lastColumNum = Sheet.GetRow(0).LastCellNum;
                    for (int i = 0; i <= lastColumNum; i++)
                    {
                        Sheet.AutoSizeColumn(i);
                        GC.Collect();
                    }
                    
                }

                // Write Excel to disk 
                var newExcelFileName = $"List_Of_Sentences-{DateTime.UtcNow.ToString("MM-dd-yyyy-HH-mm-ss")}.xlsx";
                var path = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "UploadedFilesExcel"));
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                using (var fileData = new MemoryStream())
                {
                    workbook.Write(fileData);
                    byte[] bytes = fileData.ToArray();
                    return File(bytes, "application/vnd.ms-excel", newExcelFileName);
                }
            }
            else
            {
                return NoContent();
            }
            
        }

        private void CreateCell(IRow CurrentRow, int CellIndex, string Value, HSSFCellStyle Style)
        {
            ICell Cell = CurrentRow.CreateCell(CellIndex);
            Cell.SetCellValue(Value);
            Cell.CellStyle = Style;
        }
    }
}