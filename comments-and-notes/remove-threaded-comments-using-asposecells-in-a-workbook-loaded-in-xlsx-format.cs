// <aspose-example-metadata>
// Title: Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format
// Description: Shows how to remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format with runnable C# code and valid Aspose.Cells for .NET APIs.
// Keywords: remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format in C# | Aspose.Cells Comments and Notes example | Excel automation with .NET
// Common Searches: How to remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format in C# | Comments and Notes with Aspose.Cells for .NET
// Developer Intent: Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format.
// Search Intent: How-to / data retrieval
// Use Cases: Automating comments and notes workflows in .NET applications | Server-side Excel processing without Microsoft Excel automation
// AI Prompts: Remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format. | Generate a complete runnable C# example with valid Aspose.Cells APIs.
// </aspose-example-metadata>

using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
{
    // <aspose-example-xml-summary>
    /// <summary>
    /// Shows how to remove Threaded Comments using Aspose.Cells in a workbook loaded in XLSX format with runnable C# code and valid Aspose.Cells for .NET APIs.
    /// </summary>
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Iterate through all worksheets and clear all comments (including threaded comments)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // This method removes both regular and threaded comments from the worksheet
                sheet.ClearComments();
            }

            // Save the workbook after removing the comments
            workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
