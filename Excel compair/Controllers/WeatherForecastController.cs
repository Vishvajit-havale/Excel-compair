using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace Excel_compair.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpPost("compare")]
        public async Task<IActionResult> CompareExcelFiles(IFormFile file1, IFormFile file2)
        {
            if (file1 == null || file2 == null)
                return BadRequest("Both files are required.");

            // ? Ensure .xlsx extension
            string tempPath1 = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            string tempPath2 = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            string outputPath = Path.Combine(Path.GetTempPath(), $"ExcelDiff_{Guid.NewGuid()}.xlsx");

            // Save uploaded files
            using (var stream = new FileStream(tempPath1, FileMode.Create))
                await file1.CopyToAsync(stream);

            using (var stream = new FileStream(tempPath2, FileMode.Create))
                await file2.CopyToAsync(stream);

            // Run comparison
            CompareExcels(tempPath1, tempPath2, outputPath);

            // Return result
            var memory = new MemoryStream();
            using (var stream = new FileStream(outputPath, FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;

            return File(memory,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "ComparisonResult.xlsx");
        }

        public static void CompareExcels(string file1, string file2, string outputFile)
        {
            using (var wb1 = new XLWorkbook(file1))
            using (var wb2 = new XLWorkbook(file2))
            using (var wbOut = new XLWorkbook())
            {
                var ws1 = wb1.Worksheet(1);
                var ws2 = wb2.Worksheet(1);
                var wsOut = wbOut.Worksheets.Add("Comparison");

                int rowCount = Math.Max(ws1.LastRowUsed().RowNumber(), ws2.LastRowUsed().RowNumber());
                int colCount = Math.Max(ws1.LastColumnUsed().ColumnNumber(), ws2.LastColumnUsed().ColumnNumber());

                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        var val1 = ws1.Cell(r, c).GetValue<string>();
                        var val2 = ws2.Cell(r, c).GetValue<string>();
                        var cellOut = wsOut.Cell(r, c);

                        if (val1 == val2)
                        {
                            cellOut.Value = val1;
                        }
                        else
                        {
                            cellOut.Value = $"Old: {val1}, New: {val2}";
                            cellOut.Style.Fill.BackgroundColor = XLColor.Yellow;
                            cellOut.Style.Font.FontColor = XLColor.Red;
                        }
                    }
                }

                wbOut.SaveAs(outputFile);
            }

            Console.WriteLine($"Comparison completed. Output saved at: {outputFile}");
        }
    }
}