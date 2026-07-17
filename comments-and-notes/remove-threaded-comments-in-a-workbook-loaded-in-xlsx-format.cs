// <aspose-example-metadata>
// Title: Remove Threaded Comments in a workbook loaded in XLSX format
// Description: Demonstrates how to remove threaded comments in a workbook loaded in xlsx format using Aspose.Cells for .NET.
// Prompt: Remove Threaded Comments in a workbook loaded in XLSX format.
// Product: Aspose.Cells for .NET
// Product Version: Current
// Platform: .NET
// Language: C#
// Category: Comments and Notes
// Search Intent: How-to / API example
// Developer Intent: Remove Threaded Comments in a workbook loaded in XLSX format.
// Primary APIs: Workbook, ClearComments(), Save()
// Namespaces: System, Aspose.Cells
// Input Format: None
// Output Format: XLSX
// Keywords: Aspose.Cells, Aspose.Cells for .NET, C# Excel API, Comments and Notes, Workbook, ClearComments(), Save(), System
// Common Searches: How to remove threaded comments in a workbook loaded in xlsx format in C# | Remove Threaded Comments in a workbook loaded in XLSX format using Aspose.Cells for .NET | Comments and Notes C# code example
// Summary: Demonstrates how to remove threaded comments in a workbook loaded in xlsx format using Aspose.Cells for .NET.
// </aspose-example-metadata>

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and clear all comments (including threaded comments)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clears all comments in the worksheet
                sheet.ClearComments();
            }

            // Save the workbook after removing threaded comments
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Threaded comments removed and workbook saved to: " + outputPath);
        }
    }
}
