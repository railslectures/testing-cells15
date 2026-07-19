// Title: Convert an XLSX workbook to PDF using Aspose.Cells for .NET
// Description: The sample loads an XLSX file into an Aspose.Cells Workbook, checks that the source file exists, and saves it as a PDF using the SaveFormat.Pdf option, with basic exception handling.
// Keywords: Aspose.Cells convert XLSX to PDF | C# Excel to PDF conversion | Workbook.Save PDF Aspose.Cells | SaveFormat.Pdf example | Excel to PDF using Aspose.Cells .NET
// Common Searches: How to convert an .xlsx file to PDF with Aspose.Cells in C# | Aspose.Cells C# code to save workbook as PDF | Example of Excel to PDF conversion using Aspose.Cells .NET | Convert Excel workbook to PDF handling missing file
// Developer Intent: I want to programmatically convert an XLSX workbook to a PDF file using Aspose.Cells for .NET.
// Search Intent: How-to / API implementation
// Use Cases: Automate batch conversion of multiple Excel files to PDFs in a server application. | Generate PDF reports from Excel templates in a web service. | Integrate Excel to PDF conversion in a desktop utility with error handling for missing files.
// AI Prompts: Provide C# code that uses Aspose.Cells to convert an .xlsx workbook to a PDF, including file existence checks and exception handling. | Show how to convert all Excel files in a directory to PDFs using Aspose.Cells with progress logging. | Explain how to customize PDF output options such as page size and orientation when saving a workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    // Generated SEO/GEO/AEO XML summary
    /// <summary>
    /// The sample loads an XLSX file into an Aspose.Cells Workbook, checks that the source file exists, and saves it as a PDF using the SaveFormat.Pdf option, with basic exception handling.
    /// </summary>
    public class ExcelToPdfConverter
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source Excel file (XLSX format)
            string sourcePath = "input.xlsx";

            // Desired path for the output PDF file
            string destPath = "output.pdf";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: Source file \"{sourcePath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook from the source Excel file
                Workbook workbook = new Workbook(sourcePath);

                // Save the workbook as PDF to the destination path
                workbook.Save(destPath, SaveFormat.Pdf);

                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
