using ClosedXML.Excel;
using Karpasa.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Diagnostics;

namespace Karpasa.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public IWebHostEnvironment Env { get; }

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment env)
        {
            _logger = logger;
            Env = env;
        }

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult MenClothing()
        {
            return View();
        }
        public IActionResult AboutUs()
        {
            return View();
        }
        public IActionResult Privacy()
        {
            return View();
        }
        public IActionResult ContactUS()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        //public IActionResult ContactUS(ContactViewModel model)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        string folderPath = Path.Combine(Env.WebRootPath, "files");
        //        if (!Directory.Exists(folderPath))
        //            Directory.CreateDirectory(folderPath);

        //        string filePath = Path.Combine(folderPath, "ContactSubmissions.xlsx");

        //        // Use memory stream to avoid lock conflicts
        //        using (var memoryStream = new MemoryStream())
        //        {
        //            XLWorkbook workbook;

        //            if (System.IO.File.Exists(filePath))
        //            {
        //                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        //                {
        //                    workbook = new XLWorkbook(fileStream);
        //                }
        //            }
        //            else
        //            {
        //                workbook = new XLWorkbook();
        //            }

        //            var worksheet = workbook.Worksheets.FirstOrDefault() ?? workbook.Worksheets.Add("Submissions");

        //            int row = worksheet.LastRowUsed()?.RowNumber() + 1 ?? 1;

        //            if (row == 1)
        //            {
        //                worksheet.Cell(1, 1).Value = "Name";
        //                worksheet.Cell(1, 2).Value = "Email";
        //                worksheet.Cell(1, 3).Value = "Message";
        //                worksheet.Cell(1, 4).Value = "Date Submitted";
        //                row = 2;
        //            }

        //            worksheet.Cell(row, 1).Value = model.Name;
        //            worksheet.Cell(row, 2).Value = model.Email;
        //            worksheet.Cell(row, 3).Value = model.Message;
        //            worksheet.Cell(row, 4).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

        //            workbook.SaveAs(memoryStream);
        //            workbook.Dispose(); // release memory explicitly

        //            System.IO.File.WriteAllBytes(filePath, memoryStream.ToArray()); // final write
        //        }

        //        ViewBag.Message = "Thank you! Your submission has been saved.";
        //        ModelState.Clear();
        //    }

        //    return View(model);
        //}
        public IActionResult ContactUS(ContactViewModel model)
        {
            if (ModelState.IsValid)
            {
                string folderPath = Path.Combine(Env.WebRootPath, "files");
                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                string filePath = Path.Combine(folderPath, "ContactSubmissions.xlsx");
                string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

                XLWorkbook workbook;

                // Load workbook safely without locking the file
                if (System.IO.File.Exists(filePath))
                {
                    byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
                    var stream = new MemoryStream(fileBytes); // DO NOT use "using" here
                    workbook = new XLWorkbook(stream);
                }
                else
                {
                    workbook = new XLWorkbook();
                }

                // Use a worksheet named by today's date
                string sheetName = DateTime.Now.ToString("yyyy-MM-dd");
                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName)
                                 ?? workbook.Worksheets.Add(sheetName);

                int row = worksheet.LastRowUsed()?.RowNumber() + 1 ?? 1;

                if (row == 1)
                {
                    worksheet.Cell(1, 1).Value = "Name";
                    worksheet.Cell(1, 2).Value = "Email";
                    worksheet.Cell(1, 3).Value = "Message";
                    worksheet.Cell(1, 4).Value = "DateTime";
                    row = 2;
                }

                worksheet.Cell(row, 1).Value = model.Name;
                worksheet.Cell(row, 2).Value = model.Email;
                worksheet.Cell(row, 3).Value = model.Message;
                worksheet.Cell(row, 4).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Save workbook to a temporary file to avoid locking conflicts
                workbook.SaveAs(tempFilePath);
                workbook.Dispose();

                // Replace original file with the updated temp file
                System.IO.File.Copy(tempFilePath, filePath, true);
                System.IO.File.Delete(tempFilePath);

                ViewBag.Message = "Thank you! Your submission has been saved.";
                ModelState.Clear();
                return View(new ContactViewModel());
            }

            return View(model);
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
