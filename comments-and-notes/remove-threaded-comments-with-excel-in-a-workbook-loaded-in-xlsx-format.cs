// <aspose-example-metadata>
// Title: Remove Threaded Comments with Excel in a workbook loaded in XLSX format
// Description: Demonstrates how to remove threaded comments with excel in a workbook loaded in xlsx format using Aspose.Cells for .NET.
// Prompt: Remove Threaded Comments with Excel in a workbook loaded in XLSX format.
// Product: Aspose.Cells for .NET
// Product Version: Current
// Platform: .NET
// Language: C#
// Category: Comments and Notes
// Search Intent: How-to / API example
// Developer Intent: Remove Threaded Comments with Excel in a workbook loaded in XLSX format.
// Primary APIs: Workbook, ThreadedCommentCollection, GetThreadedComments(), Clear(), Save()
// Namespaces: System, Aspose.Cells
// Input Format: None
// Output Format: XLSX
// Keywords: Aspose.Cells, Aspose.Cells for .NET, C# Excel API, Comments and Notes, Workbook, ThreadedCommentCollection, GetThreadedComments(), Clear(), Save(), System
// Common Searches: How to remove threaded comments with excel in a workbook loaded in xlsx format in C# | Remove Threaded Comments with Excel in a workbook loaded in XLSX format using Aspose.Cells for .NET | Comments and Notes C# code example
// Summary: Demonstrates how to remove threaded comments with excel in a workbook loaded in xlsx format using Aspose.Cells for .NET.
// </aspose-example-metadata>

using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (XLSX format)
            string inputPath = "input.xlsx";

            // Path for the resulting workbook after removing threaded comments
            string outputPath = "output.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Determine the used range limits
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxCol = worksheet.Cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Retrieve threaded comments for the current cell
                        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(row, col);

                        // If there are any threaded comments, clear them
                        if (threadedComments != null && threadedComments.Count > 0)
                        {
                            threadedComments.Clear();
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
