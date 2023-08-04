using ClosedXML.Excel;
using ExcelDownloadUsingClosedXML.Models;
using Microsoft.AspNetCore.Mvc;

namespace ExcelDownloadUsingClosedXML.Controllers
{
    public class StudentController : Controller
    {
        // Mocked student list for demo purposes
        private static List<Student> _students = new List<Student>
        {
        new Student { Id = 1, Name = "John Doe", Age = 25 },
        new Student { Id = 2, Name = "Jane Smith", Age = 22 },
        // Add more students as needed
        };

        public IActionResult Index()
        {
            return View(_students);
        }

        public IActionResult DownloadExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Students");

                // Add headers
                worksheet.Cell(1, 1).Value = "ID";
                worksheet.Cell(1, 2).Value = "Name";
                worksheet.Cell(1, 3).Value = "Age";

                // Add data rows
                int rowIndex = 2;
                foreach (var student in _students)
                {
                    worksheet.Cell(rowIndex, 1).Value = student.Id;
                    worksheet.Cell(rowIndex, 2).Value = student.Name;
                    worksheet.Cell(rowIndex, 3).Value = student.Age;
                    rowIndex++;
                }

                // Save the Excel file to a MemoryStream
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);

                    // Create a new MemoryStream with the data from the previous MemoryStream
                    var copiedStream = new MemoryStream(stream.ToArray());

                    // Set the file name and content type for the response
                    string fileName = "StudentList.xlsx";
                    string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                    // Return the file as a downloadable attachment
                    return File(copiedStream, contentType, fileName);
                }
            }
        }
    }
}