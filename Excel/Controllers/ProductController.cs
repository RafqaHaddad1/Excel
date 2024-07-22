using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;

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
            // Register encoding provider for reading Excel files
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            if (file == null || file.Length == 0)
            {
                ViewBag.Error = "Please select an excel file";
                return View("Index");
            }
            else
            {
                if (file.FileName.EndsWith("xls") || file.FileName.EndsWith("xlsx"))
                {
                    var uploadDirectory = $"{Directory.GetCurrentDirectory()}\\wwwroot\\Uploads";

                    if (!Directory.Exists(uploadDirectory))
                    {
                        Directory.CreateDirectory(uploadDirectory);
                    }

                    var filepath = Path.Combine(uploadDirectory, file.FileName);

                    using (var stream = new FileStream(filepath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }

                    var excelDataBySheet = new Dictionary<string, List<List<object>>>();

                    using (var stream = System.IO.File.Open(filepath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            do
                            {
                                var sheetData = new List<List<object>>();
                                var sheetName = reader.Name;

                                while (reader.Read())
                                {
                                    var rowData = new List<object>();
                                    for (int column = 0; column < reader.FieldCount; column++)
                                    {
                                        rowData.Add(reader.GetValue(column));
                                    }
                                    sheetData.Add(rowData);
                                }

                                excelDataBySheet.Add(sheetName, sheetData);

                            } while (reader.NextResult());
                        }
                    }

                    ViewBag.ExcelDataBySheet = excelDataBySheet;
                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "Incorrect file type";
                    return View("Index");
                }
            }
        }

    }
}
/*

 1. Initial Setup:
Registers the CodePagesEncodingProvider to ensure proper handling of various Excel file encodings.

2. File Validation:
Checks if the uploaded file (IFormFile named file) is null or empty.
If empty, sets an error message in ViewBag and redirects to the "Index" view.

3. File Processing:
Validates if the filename ends with ".xls" or ".xlsx" extensions (indicating Excel files).
If not valid, sets an error message and redirects to "Index".
Creates the upload directory ("./wwwroot/Uploads") if it doesn't exist.
Combines the upload directory path with the filename to create the full file path (filepath).
Saves the uploaded file to the specified filepath using a FileStream.

4. Reading Excel Data:
Opens the uploaded file in read mode using File.Open.
Creates an ExcelDataReader object using ExcelReaderFactory. This likely uses a third-party library (not included in the code snippet) for handling Excel file formats.
Loops through each sheet in the Excel file using do...while loop with NextResult():
Initializes an empty list sheetData to store data for the current sheet.
Gets the current sheet name using reader.Name.
Loops through each row in the sheet using reader.Read():
Initializes an empty list rowData to store a single row's data.
Loops through each column in the row using a for loop:
Reads the value for the current cell using reader.GetValue(column).
Adds the cell value to the rowData list.
Adds the entire rowData (representing a single row) to the sheetData list.
After processing all rows in the sheet:
Adds the sheetData (containing all rows for the sheet) to a dictionary named excelDataBySheet with the sheet name as the key.

5. Displaying Results:
Stores the processed data (excelDataBySheet) in ViewBag.
Redirects to the "Success" view, presumably displaying the imported data.
Overall, this code snippet demonstrates how to upload an Excel file, validate its format, and then read the data from each sheet into a dictionary structure for further processing or display.

 */
