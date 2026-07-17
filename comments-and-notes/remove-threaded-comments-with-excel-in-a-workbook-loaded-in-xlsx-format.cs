// <example-metadata>
// Title: Remove Threaded Comments with Excel in a workbook loaded in XLSX format
// Description: Shows how to remove Threaded Comments with Excel in a workbook loaded in XLSX format. using Aspose.Cells for .NET.
// Prompt: Remove Threaded Comments with Excel in a workbook loaded in XLSX format.
// Product: Aspose.Cells for .NET
// Language: C#
// Category: Comments and Notes
// Search Intent: How-to code example
// Developer Intent: Remove Threaded Comments with Excel in a workbook loaded in XLSX format.
// Difficulty: Beginner
// Primary APIs: Cells, Workbook, Worksheet, Save
// Namespaces: Aspose.Cells, System
// Input: Example-specific input or none
// Output: XLSX
// Keywords: Aspose.Cells, C#, .NET, Comments and Notes, Cells, Workbook, Worksheet, Save, Excel automation, spreadsheet API
// Common Searches: Remove Threaded Comments with Excel in a workbook loaded in XLSX format; Comments and Notes example in C#; Aspose.Cells Comments and Notes .NET
// AI Retrieval Summary: Shows how to remove Threaded Comments with Excel in a workbook loaded in XLSX format. using Aspose.Cells for .NET.
// // </example-metadata>

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
                // The Clear method removes both legacy and threaded comments from the worksheet
                sheet.Comments.Clear();
            }

            // Save the workbook after removing threaded comments
            string outputPath = "OutputWorkbook_NoThreadedComments.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}