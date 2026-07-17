// <aspose-example-metadata>
// Title: Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format
// Description: Demonstrates how to remove threaded comments using aspose.cells in a workbook loaded in xlsx format using Aspose.Cells for .NET.
// Prompt: Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format.
// Product: Aspose.Cells for .NET
// Product Version: Current
// Platform: .NET
// Language: C#
// Category: Comments and Notes
// Search Intent: How-to / API example
// Developer Intent: Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format.
// Primary APIs: Workbook, ClearComments(), Save()
// Namespaces: System, Aspose.Cells
// Input Format: XLSX
// Output Format: XLSX
// Keywords: Aspose.Cells, Aspose.Cells for .NET, C# Excel API, Comments and Notes, Workbook, ClearComments(), Save(), System
// Common Searches: How to remove threaded comments using aspose.cells in a workbook loaded in xlsx format in C# | Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format using Aspose.Cells for .NET | Comments and Notes C# code example
// Summary: Demonstrates how to remove threaded comments using aspose.cells in a workbook loaded in xlsx format using Aspose.Cells for .NET.
// </aspose-example-metadata>

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
                // Clears both regular and threaded comments from the worksheet
                sheet.ClearComments();
            }

            // Save the workbook after removing the comments
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
