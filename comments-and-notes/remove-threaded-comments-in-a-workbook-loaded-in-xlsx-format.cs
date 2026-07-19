// Title: Remove Threaded Comments from an XLSX Workbook with Aspose.Cells for .NET
// Description: Loads an XLSX file into an Aspose.Cells Workbook, iterates through each worksheet, calls ClearComments() to delete all comments—including threaded ones—and saves the cleaned workbook as a new XLSX file.
// Keywords: Aspose.Cells remove threaded comments | ClearComments C# | delete Excel comments programmatically | strip comments from XLSX .NET | Aspose.Cells comment removal example | C# Excel comment cleanup
// Common Searches: How to delete threaded comments in XLSX using Aspose.Cells C# | Aspose.Cells ClearComments method example | Remove all comments from Excel workbook .NET | C# code to clear Excel comments with Aspose | Batch delete comments from multiple XLSX files Aspose.Cells
// Developer Intent: Eliminate every comment—regular and threaded—from a loaded workbook and save the result.
// Use Cases: Prepare confidential spreadsheets for external sharing by stripping all notes. | Pre‑process incoming Excel files before data import to avoid hidden annotations. | Create a clean archival copy of a workbook after removing user comments.
// AI Prompts: Show a C# snippet that removes only threaded comments while keeping legacy cell comments using Aspose.Cells. | Generate code that logs the count of comments removed per worksheet after ClearComments is called. | Explain how to loop through a folder of XLSX files and purge comments from each with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an XLSX file into an Aspose.Cells Workbook, iterates through each worksheet, calls ClearComments() to delete all comments—including threaded ones—and saves the cleaned workbook as a new XLSX file.
class RemoveThreadedComments
{
    static void Main()
    {
        // Load the workbook from an existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Clear all comments (including threaded comments) from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments();
        }

        // Save the workbook after removing the threaded comments
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
