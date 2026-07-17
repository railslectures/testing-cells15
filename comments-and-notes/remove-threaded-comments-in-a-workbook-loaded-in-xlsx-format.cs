using System;
using Aspose.Cells;

class RemoveThreadedComments
{
    static void Main()
    {
        // Load the workbook from an existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Access the collection of comments in the current worksheet
            CommentCollection comments = worksheet.Comments;

            // Loop over all comments and clear their threaded comments
            for (int i = 0; i < comments.Count; i++)
            {
                Comment comment = comments[i];
                // Remove all threaded comments associated with this comment
                comment.ThreadedComments.Clear();
            }
        }

        // Save the modified workbook to a new XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}