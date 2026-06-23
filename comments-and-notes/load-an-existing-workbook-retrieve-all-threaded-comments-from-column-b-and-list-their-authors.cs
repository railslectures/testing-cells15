using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThreadedComments
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: File \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection of comments for the worksheet
                CommentCollection comments = worksheet.Comments;

                // Determine the used range to know how many rows to scan
                int totalRows = worksheet.Cells.MaxDisplayRange.RowCount;

                // Column B has zero‑based index 1
                const int columnBIndex = 1;

                // Use a set to collect distinct author names
                HashSet<string> authorNames = new HashSet<string>();

                // Iterate through each row in column B
                for (int row = 0; row < totalRows; row++)
                {
                    // Retrieve threaded comments for the current cell (row, column B)
                    ThreadedCommentCollection threadedComments = comments.GetThreadedComments(row, columnBIndex);

                    // If there are no threaded comments, skip to the next cell
                    if (threadedComments == null) continue;

                    // Add each comment's author to the set
                    foreach (ThreadedComment tc in threadedComments)
                    {
                        if (tc?.Author?.Name != null && tc.Author.Name.Length > 0)
                        {
                            authorNames.Add(tc.Author.Name);
                        }
                    }
                }

                // Output the list of authors
                Console.WriteLine("Authors of threaded comments in column B:");
                foreach (string name in authorNames)
                {
                    Console.WriteLine(name);
                }

                // Optional: Save the workbook if any modifications were made
                // workbook.Save("OutputWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}