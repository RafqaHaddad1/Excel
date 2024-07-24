using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
namespace Excel.Controllers
{
    public class ProductController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Import(IFormFile file)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            if (file == null || file.Length == 0)
            {
                ViewBag.Error = "Please select an excel file";
                return View("Index");
            }

            var fileExtension = Path.GetExtension(file.FileName);
            if (fileExtension != ".xls" && fileExtension != ".xlsx")
            {
                ViewBag.Error = "Incorrect file type";
                return View("Index");
            }

            try
            {
                var excelDataBySheet = await ReadExcelDataAsync(file);

                ViewBag.ExcelDataBySheet = excelDataBySheet;
                ViewBag.ExcelDataBySheetJson = JsonConvert.SerializeObject(excelDataBySheet);
                return View("Success");
            }
            catch (Exception ex)
            {
                ViewBag.Error = $"An error occurred: {ex.Message}";
                return View("Index");
            }
        }

        private async Task<Dictionary<string, List<List<object>>>> ReadExcelDataAsync(IFormFile file)
        {
            var excelDataBySheet = new Dictionary<string, List<List<object>>>();

            using (var stream = file.OpenReadStream())
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        var sheetData = new List<List<object>>();
                        var sheetName = reader.Name;

                        while (reader.Read())
                        {
                            sheetData.Add(Enumerable.Range(0, reader.FieldCount)
                                .Select(reader.GetValue)
                                .ToList());
                        }

                        excelDataBySheet.Add(sheetName, sheetData);

                    } while (reader.NextResult());
                }
            }
            return excelDataBySheet;
        }
    }
}
