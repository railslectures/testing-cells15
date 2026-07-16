using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class WorkbookToTiff
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Desired output TIFF file path
        string outputPath = "output.tiff";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(sourcePath);

        // Configure image rendering options for TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,                     // Set output format to TIFF
            TiffCompression = TiffCompression.CompressionLZW, // Use LZW compression for better quality
            HorizontalResolution = 300,                     // 300 DPI horizontal resolution
            VerticalResolution = 300,                       // 300 DPI vertical resolution
            OnePagePerSheet = true                          // Render each worksheet as a single page to preserve layout
        };

        // Create a renderer for the whole workbook using the configured options
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render the entire workbook to a multi‑page TIFF file
        renderer.ToImage(outputPath);

        Console.WriteLine($"Workbook successfully rendered to TIFF: {outputPath}");
    }
}