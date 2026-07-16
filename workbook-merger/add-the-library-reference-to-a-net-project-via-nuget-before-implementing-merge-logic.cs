// Ensure the Aspose.Cells library is added to the project via NuGet:
//   Install-Package Aspose.Cells
//   or add the package reference in the .csproj file.

using System;
using System.IO;
using Aspose.Cells;

namespace MergeExcelFilesDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Prepare sample workbooks to demonstrate the merge.
            string[] sourceFiles = new string[2];
            sourceFiles[0] = "Source1.xlsx";
            sourceFiles[1] = "Source2.xlsx";

            // Create first workbook.
            Workbook wb1 = new Workbook();
            wb1.Worksheets[0].Cells["A1"].PutValue("Data from File 1");
            wb1.Save(sourceFiles[0]);

            // Create second workbook.
            Workbook wb2 = new Workbook();
            wb2.Worksheets[0].Cells["A1"].PutValue("Data from File 2");
            wb2.Save(sourceFiles[1]);

            // Temporary file used by CellsHelper.MergeFiles for caching.
            string cacheFile = "MergeCache.tmp";

            // Destination file that will contain the merged result.
            string destFile = "MergedResult.xlsx";

            try
            {
                // Merge the source files into the destination file.
                // This uses the CellsHelper.MergeFiles method as defined in the Aspose.Cells API.
                CellsHelper.MergeFiles(sourceFiles, cacheFile, destFile);

                Console.WriteLine($"Files merged successfully. Output saved to: {destFile}");

                // Verify the merged content.
                Workbook mergedWorkbook = new Workbook(destFile);
                Console.WriteLine("Merged workbook content:");
                Console.WriteLine($"Sheet1!A1: {mergedWorkbook.Worksheets[0].Cells["A1"].StringValue}");
                Console.WriteLine($"Sheet2!A1: {mergedWorkbook.Worksheets[1].Cells["A1"].StringValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merge: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files.
                foreach (var file in sourceFiles)
                {
                    if (File.Exists(file))
                        File.Delete(file);
                }

                if (File.Exists(cacheFile))
                    File.Delete(cacheFile);
            }
        }
    }
}