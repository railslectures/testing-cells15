// Title: C# – Remove Threaded Comments from an XLSX Workbook with Aspose.Cells
// Description: Loads an XLSX file using Aspose.Cells for .NET, iterates every worksheet, clears all ThreadedCommentCollection items while keeping regular comments, and saves the cleaned workbook.
// Keywords: Aspose.Cells remove threaded comments | C# delete Excel threaded comments | Aspose.Cells ThreadedCommentCollection | clear comment threads XLSX | .NET Excel comment cleanup | Aspose.Cells workbook sanitization | remove Excel comment threads programmatically | Aspose.Cells API comment removal
// Common Searches: how to delete threaded comments in Excel using Aspose.Cells C# | remove all comment threads from an XLSX workbook .NET | Aspose.Cells example for clearing threaded comments | C# code to strip Excel threaded comments with Aspose | bulk remove Excel comment threads Aspose.Cells
// Developer Intent: Programmatically delete every threaded comment in a loaded XLSX workbook and save the result.
// Use Cases: Sanitize workbooks before distribution by stripping confidential threaded notes. | Automate cleanup of comment threads in batch‑processed Excel reports. | Archive Excel files where threaded comment data is unnecessary while preserving regular comments.
// AI Prompts: Generate C# code that uses Aspose.Cells to remove all threaded comments from a workbook but leaves regular comments intact. | Create a reusable method that accepts input and output XLSX paths, clears threaded comments on every sheet, and returns a status message. | Explain the most efficient Aspose.Cells approach to delete comment threads without affecting other worksheet data.

using System;
using Aspose.Cells;

// Loads an XLSX file using Aspose.Cells for .NET, iterates every worksheet, clears all ThreadedCommentCollection items while keeping regular comments, and saves the cleaned workbook.
class RemoveThreadedCommentsDemo
{
    static void Main()
    {
        // Load the XLSX workbook (lifecycle rule: load)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the collection of comments on the current worksheet
            CommentCollection comments = sheet.Comments;

            // Loop through all comments (regular and threaded)
            for (int i = 0; i < comments.Count; i++)
            {
                // Get the individual comment
                Comment comment = comments[i];

                // Obtain the threaded comment collection associated with this comment
                ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                // If there are any threaded comments, remove them one by one
                if (threadedComments != null && threadedComments.Count > 0)
                {
                    // Remove from the end to avoid index shifting
                    for (int j = threadedComments.Count - 1; j >= 0; j--)
                    {
                        threadedComments.RemoveAt(j);
                    }
                }
            }
        }

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
