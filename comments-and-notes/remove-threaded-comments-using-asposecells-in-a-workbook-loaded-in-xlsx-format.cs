// <example-metadata>
// Title: Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format
// Description: Shows how to remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format. using Aspose.Cells for .NET.
// Prompt: Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format.
// Product: Aspose.Cells for .NET
// Language: C#
// Category: Comments and Notes
// Search Intent: How-to code example
// Developer Intent: Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format.
// Difficulty: Beginner
// Primary APIs: Cells, Workbook, Worksheet, ClearComments, Save
// Namespaces: Aspose.Cells, System
// Input: Example-specific input or none
// Output: XLSX
// Keywords: Aspose.Cells, C#, .NET, Comments and Notes, Cells, Workbook, Worksheet, ClearComments, Save, Excel automation, spreadsheet API
// Common Searches: Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format; Comments and Notes example in C#; Aspose.Cells Comments and Notes .NET
// AI Retrieval Summary: Shows how to remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format. using Aspose.Cells for .NET.
// // </example-metadata>

using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets and clear all comments (including threaded comments)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // This method removes every comment attached to the worksheet,
                // which also clears any threaded comments present.
                sheet.ClearComments();
            }

            // Save the workbook after removing the comments
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}