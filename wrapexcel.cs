using OfficeOpenXml;
using System;
using System.IO;

namespace SwiftBreakerGUI.BLL
{
    internal class WrapColumnAndRow
    {
        public static void WrapTextInExcel(string filePath)
        {
            try
            {
                // Validate file existence
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException("Specified Excel file not found.", filePath);
                }

                // Create FileInfo object for input and output paths
                FileInfo fileInfo = new FileInfo(filePath);
                string directoryName = fileInfo.DirectoryName;
                string outputFileName = Path.GetFileNameWithoutExtension(fileInfo.Name) + "_wrapped" + fileInfo.Extension;
                string outputFilePath = Path.Combine(directoryName, outputFileName);

                // Set EPPlus license context
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                // Load and process the Excel file
                using (var package = new ExcelPackage(fileInfo))
                {
                    // Get the first worksheet
                    var worksheet = package.Workbook.Worksheets[0];

                    // Get the dimensions of the worksheet
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    Console.WriteLine("Wrapping text in all cells...");

                    // Iterate through each cell and apply wrapping
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cell = worksheet.Cells[row, col];
                            if (cell.Value != null)
                            {
                                cell.Style.WrapText = true; // Apply wrapping style
                            }
                        }
                    }

                    // Save the updated file
                    package.SaveAs(new FileInfo(outputFilePath));
                }

                Console.WriteLine($"Text wrapping applied successfully. Saved as: {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Main(string[] args)
        {
            // Hardcoded file path
            string filePath = @"C:\Users\sudip.adhikari4670\Desktop\swift test\output.xlsx";

            Console.WriteLine($"Processing file: {filePath}");
            WrapTextInExcel(filePath);
        }
    }
}
