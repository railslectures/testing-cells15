using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveThreadedCommentsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the input XLSX workbook
            string inputPath = "InputWorkbook.xlsx";

            // Prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and clear all comments,
            // which also removes threaded comments.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // ClearComments removes both classic and threaded comments.
                sheet.ClearComments();
            }

            // Path to the output workbook (lifecycle: save)
            string outputPath = "OutputWorkbook_NoThreadedComments.xlsx";

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}