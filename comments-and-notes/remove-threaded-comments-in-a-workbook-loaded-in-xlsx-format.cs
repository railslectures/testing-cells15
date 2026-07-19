// Title: Remove Threaded and Standard Comments from an XLSX Workbook with Aspose.Cells for .NET
// Description: Loads an XLSX file, iterates through each worksheet, calls ClearComments to delete both threaded and regular comments, and saves the cleaned workbook.
// Keywords: Aspose.Cells ClearComments | delete threaded comments C# | remove Excel notes programmatically | .NET workbook comment cleanup | Aspose.Cells comment removal example
// Common Searches: how to delete threaded comments using Aspose.Cells .NET | clear all comments from Excel workbook C# | Aspose.Cells remove notes from XLSX | programmatic comment cleanup Aspose.Cells
// Developer Intent: Strip every comment—including threaded ones—from an XLSX workbook via C#.
// Use Cases: Sanitize shared spreadsheets by eliminating all user remarks. | Reset a template after data entry so no comments remain. | Batch‑process multiple Excel files to ensure they contain no notes.
// AI Prompts: Generate C# code that removes only threaded comments while keeping regular ones with Aspose.Cells. | Show how to enumerate comment authors before clearing them in a workbook. | Provide a script to walk through a folder of XLSX files and purge comments from each using Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an XLSX file, iterates through each worksheet, calls ClearComments to delete both threaded and regular comments, and saves the cleaned workbook.
class RemoveThreadedComments
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet and clear all comments (standard and threaded)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments(); // Removes every comment, including threaded ones
        }

        // Save the workbook after comments have been removed
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
