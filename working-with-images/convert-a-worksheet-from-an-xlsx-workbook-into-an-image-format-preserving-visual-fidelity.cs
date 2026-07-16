using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class WorksheetToImageDemo
    {
        public static void Run()
        {
            try
            {
                // Path to the source XLSX workbook
                string sourceWorkbookPath = "input.xlsx";

                // Path where the rendered image will be saved
                string outputImagePath = "worksheet_image.png";

                // Verify that the source file exists
                if (!File.Exists(sourceWorkbookPath))
                {
                    Console.WriteLine($"Source file not found: {sourceWorkbookPath}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputImagePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook inside a using block (Workbook implements IDisposable)
                using (Workbook workbook = new Workbook(sourceWorkbookPath))
                {
                    // Access the first worksheet
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Configure image rendering options
                    ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                    {
                        ImageType = Aspose.Cells.Drawing.ImageType.Png,
                        OnePagePerSheet = true
                    };

                    // Render the worksheet to an image (SheetRender does NOT implement IDisposable)
                    SheetRender sheetRender = new SheetRender(worksheet, renderOptions);
                    sheetRender.ToImage(0, outputImagePath);
                }

                Console.WriteLine($"Worksheet rendered successfully to: {outputImagePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetToImageDemo.Run();
        }
    }
}