// Title: C# – Remove Threaded and Regular Comments from an XLSX Workbook with Aspose.Cells
// Description: Loads an existing XLSX file into a Workbook, loops through each worksheet, calls ClearComments() to erase both standard and threaded comments, and saves the cleaned file as a new XLSX document.
// Keywords: Aspose.Cells ClearComments | remove threaded comments C# | delete Excel comments .NET | strip workbook notes | Aspose.Cells comment removal | C# Excel comment cleanup
// Common Searches: Aspose.Cells delete all comments C# | how to clear threaded comments in XLSX using .NET | remove Excel notes programmatically Aspose | C# code to purge comments from workbook | clear comments from every sheet Aspose.Cells
// Developer Intent: Programmatically purge every comment—both standard and threaded—from an XLSX workbook using Aspose.Cells for .NET.
// Use Cases: Sanitize a workbook before distribution by stripping all user feedback. | Prepare a clean Excel template for external partners without any notes. | Automate comment cleanup when processing large batches of Excel files.
// AI Prompts: Generate C# code that removes only threaded comments while keeping regular comments with Aspose.Cells. | Show how to enumerate comment authors before clearing them in a workbook. | Explain how to verify that no comments remain after invoking ClearComments.

using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
{
    // Loads an existing XLSX file into a Workbook, loops through each worksheet, calls ClearComments() to erase both standard and threaded comments, and saves the cleaned file as a new XLSX document.
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
