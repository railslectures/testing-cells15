// Title: Convert an XLSX workbook to XPS using Aspose.Cells ConversionUtility in C#
// Description: The sample uses Aspose.Cells ConversionUtility.Convert to transform an input.xlsx file into an output.xps document. It demonstrates a simple one‑line conversion call.
// Keywords: Aspose.Cells convert XLSX to XPS | C# Excel to XPS conversion | ConversionUtility Convert method | Export Excel workbook as XPS | Aspose.Cells XPS output
// Common Searches: How to convert an Excel .xlsx file to XPS using Aspose.Cells in C# | Aspose.Cells ConversionUtility example for XLSX to XPS | C# code to export workbook to XPS format with Aspose | Generate XPS from Excel workbook programmatically
// Developer Intent: I want to convert an Excel .xlsx workbook to an XPS document using Aspose.Cells.
// Search Intent: How-to / API implementation
// Use Cases: Batch convert multiple .xlsx reports to XPS for printing or archiving. | Automate generation of XPS files from Excel data in a scheduled job. | Expose a web service endpoint that receives an Excel file and returns an XPS document.
// AI Prompts: Show me C# code that converts an Excel workbook to XPS using Aspose.Cells with proper error handling. | Provide a script that iterates over all .xlsx files in a folder and creates corresponding .xps files using ConversionUtility. | Explain how to configure page layout options when exporting to XPS with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Generated SEO/GEO/AEO XML summary
/// <summary>
/// The sample uses Aspose.Cells ConversionUtility.Convert to transform an input.xlsx file into an output.xps document. It demonstrates a simple one‑line conversion call.
/// </summary>
class Program
{
    static void Main()
    {
        // Path to the source Excel file (XLSX)
        string sourcePath = "input.xlsx";

        // Desired output path for the XPS file
        string destPath = "output.xps";

        // Convert the Excel workbook to XPS format using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Conversion completed: " + sourcePath + " -> " + destPath);
    }
}
