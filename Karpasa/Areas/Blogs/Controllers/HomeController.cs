using Karpasa.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace Karpasa.Areas.Blogs.Controllers
{
    [Area("Blogs")]
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment env;

        public HomeController(IWebHostEnvironment env)
        {
            this.env = env;
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
        public IActionResult ContactUS()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]

        public IActionResult ContactUS(ContactViewModel model)
        
        {
            Console.WriteLine($"Name: {model.Name}, Email: {model.Email}, Message: {model.Message}");

            if (ModelState.IsValid)
            {
                string folderPath = Path.Combine(env.WebRootPath, "files");
                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                string filePath = Path.Combine(folderPath, "ContactSubmissions.xlsx");
                FileInfo fileInfo = new FileInfo(filePath);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage package = fileInfo.Exists ? new ExcelPackage(fileInfo) : new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("Submissions");

                    int row = worksheet.Dimension?.Rows + 1 ?? 1;

                    if (row == 1)
                    {
                        worksheet.Cells[1, 1].Value = "Name";
                        worksheet.Cells[1, 2].Value = "Email";
                        worksheet.Cells[1, 3].Value = "Message";
                        worksheet.Cells[1, 4].Value = "Date Submitted";
                        row = 2;
                    }

                    worksheet.Cells[row, 1].Value = model.Name;
                    worksheet.Cells[row, 2].Value = model.Email;
                    worksheet.Cells[row, 3].Value = model.Message;
                    worksheet.Cells[row, 4].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

                    package.Save();
                }

                ViewBag.Message = "Thank you! Your submission has been saved.";
                ModelState.Clear();
            }

            return View();
        }
    }
}

