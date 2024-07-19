using ExcelManagement.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using ExcelDataReader;
using System.Xml.Linq;
using CsvHelper;
using System.Globalization;

namespace ExcelManagement.Controllers
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

        [HttpPost]
        public async Task<IActionResult> Index(IFormFile file)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            
            if (file != null && file.Length > 0)
            {
                var fileType = Path.GetExtension(file.FileName);
                if (fileType.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
                    fileType.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    // Upload File
                    var uploadDirectory = $"{Directory.GetCurrentDirectory()}\\wwwroot\\Uploads";

                    if (!Directory.Exists(uploadDirectory))
                    {
                        Directory.CreateDirectory(uploadDirectory);
                    }

                    var filePath = Path.Combine(uploadDirectory, file.FileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }

                    //Read File .xlsx/.xls
                    using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        var excelData = new List<List<object>>();
                        int excelRow = 0;
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            while (reader.Read())
                            {
                                excelRow++;
                            }

                            ViewBag.ExcelData = excelData;
                            ViewBag.ExcelRow = excelRow;
                        }
                    }
                }
                else
                {
                    //Read File .csv
                    var excelData = new List<List<object>>();
                    int excelRow = 0;
                    using (var reader = new StreamReader(file.OpenReadStream()))
                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        while (csv.Read())
                        {
                            excelRow++;
                        }
                        ViewBag.ExcelRow = excelRow;
                    }
                }
            } 

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
