// <aspose-example-metadata>
// Title: Remove Threaded Comments with Excel in a workbook loaded in XLSX format using Aspose.Cells for .NET
// Description: Shows how to remove Threaded Comments with Excel in a workbook loaded in XLSX format with runnable C# code and valid Aspose.Cells for .NET APIs.
// Keywords: remove Threaded Comments with Excel in a workbook loaded in XLSX format in C# | Aspose.Cells Comments and Notes example | Excel automation with .NET
// Common Searches: How to remove Threaded Comments with Excel in a workbook loaded in XLSX format in C# | Comments and Notes with Aspose.Cells for .NET
// Developer Intent: Remove Threaded Comments with Excel in a workbook loaded in XLSX format.
// Search Intent: How-to / data retrieval
// Use Cases: Automating comments and notes workflows in .NET applications | Server-side Excel processing without Microsoft Excel automation
// AI Prompts: Remove Threaded Comments with Excel in a workbook loaded in XLSX format. | Generate a complete runnable C# example with valid Aspose.Cells APIs.
// </aspose-example-metadata>

using System;
using Aspose.Cells;

// <aspose-example-xml-summary>
/// <summary>
/// Shows how to remove Threaded Comments with Excel in a workbook loaded in XLSX format with runnable C# code and valid Aspose.Cells for .NET APIs.
/// </summary>
class RemoveThreadedComments
{
    static void Main()
    {
        // Load the workbook (XLSX format)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Remove all comments (including threaded comments) from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments(); // clears both legacy and threaded comments
        }

        // Save the workbook after removal
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
