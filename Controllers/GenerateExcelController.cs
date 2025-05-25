using ConvertDownloadMsExcel3.Models;
using DinkToPdf;
using DinkToPdf.Contracts;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;

namespace ConvertDownloadMsExcel2.Controllers
{
    public class GenerateExcelController : Controller // ✅ inherit from Controller
    {

        private readonly IConverter _converter;

        [HttpGet]
        public async Task<IActionResult> GenerateExcel()
        {
            var products = GetProductList();

            var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "ExcelTemplate", "ProductReport.html");
            string htmlData = System.IO.File.ReadAllText(templatePath);

            string excelString = string.Join("", products.Select(p =>
                $"<tr><td>{p.Name}</td><td>{p.Description}</td><td>{p.Price}</td></tr>"));

            htmlData = htmlData.Replace("@@ActualData", excelString);

            string filePath = Path.Combine("ExcelFiles", $"{DateTime.Now.Ticks}.xls");
            Directory.CreateDirectory("ExcelFiles");
            System.IO.File.WriteAllText(filePath, htmlData);

            var provider = new FileExtensionContentTypeProvider();
            if (!provider.TryGetContentType(filePath, out var contentType))
            {
                contentType = "application/octet-stream";
            }

            var bytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(bytes, contentType, Path.GetFileName(filePath));
        }

        public static List<ProductModel> GetProductList()
        {
            return new List<ProductModel>
            {
                new ProductModel{Name="Mouse",Description="Laptop Mouse",Price=300 },
                new ProductModel{Name="Desktop",Description="HP Desktop",Price=2300 },
                new ProductModel{Name="Speaker",Description="Boat Speaker",Price=700 },
                new ProductModel{Name="Printer",Description="HP Printer",Price=4300 },
                new ProductModel{Name="Key Board",Description="DELL Key Board",Price=200 },
                new ProductModel{Name="Camera",Description="Nikon Camera",Price=1000 },
                new ProductModel{Name="Scanner",Description="Xerox Scanner",Price=8000 },
            };
        }


        public IActionResult Index()
        {
            return View(); // View at Views/GenerateExcel/Index.cshtml
        }

        public GenerateExcelController(IConverter converter)
        {
            _converter = converter;
        }

        public IActionResult GeneratePdf()
        {
            var html = @"
                <html>
                <head><title>PDF Report</title></head>
                <body>
                    <h1>Product Report</h1>
                    <table border='1'>
                        <tr><th>Name</th><th>Description</th><th>Price</th></tr>
                        <tr><td>Mouse</td><td>Laptop Mouse</td><td>300</td></tr>
                        <tr><td>Keyboard</td><td>Mechanical</td><td>500</td></tr>
                        <tr><td>Monitor</td><td>4K Display</td><td>1200</td></tr>
                    </table>
                </body>
                </html>";

            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = {
                    PaperSize = PaperKind.A4,
                    Orientation = Orientation.Portrait
                },
                Objects = {
                    new ObjectSettings()
                    {
                        HtmlContent = html
                    }
                }
            };

            var pdf = _converter.Convert(doc);
            return File(pdf, "application/pdf", "products.pdf");
        }

    }
}
