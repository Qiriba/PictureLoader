using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using ClosedXML.Excel;

class Program
{
    static async Task Main()
    {
        // Gather inputs from the user
        Console.Write("Please enter the full path to the Excel file but without any \" at start or end: ");
        string excelFilePath = Console.ReadLine();

        Console.Write("Enter the folder to save images (e.g., C:\\Users\\User\\Pictures): ");
        string saveFolder = Console.ReadLine();

        Console.Write("Enter the column number with names (e.g., 1 for column A): ");
        int columnNumber = int.Parse(Console.ReadLine());

        Console.Write("Enter the starting row (e.g., 2 for data starting after headers): ");
        int startRow = int.Parse(Console.ReadLine());

        Console.Write("Enter the base URL (e.g., https://example.com/url): ");
        string baseUrl = Console.ReadLine();

        // Ensure the save folder exists
        Directory.CreateDirectory(saveFolder);

        try
        {
            // Step 1: Read names from Excel
            List<string> names = ReadNamesFromExcel(excelFilePath, columnNumber, startRow);

            Console.WriteLine($"Found {names.Count} names. Starting downloads...");

            // Step 2: Download images for each name
            var downloadTasks = new List<Task>();
            foreach (string name in names)
            {
                string imageUrl = $"{baseUrl}/{name}.jpg";
                string savePath = Path.Combine(saveFolder, $"{name}.jpg");

                // Start the download task
                downloadTasks.Add(DownloadImageAsync(imageUrl, savePath));
            }

            // Wait for all downloads to complete
            await Task.WhenAll(downloadTasks);
            Console.WriteLine("All images downloaded successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    static List<string> ReadNamesFromExcel(string filePath, int columnNumber, int startRow)
    {
        var names = new List<string>();

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(1); // Assuming the first worksheet
            var column = worksheet.Column(columnNumber);

            // Iterate through each cell starting from the specified row
            foreach (var cell in column.CellsUsed(c => c.Address.RowNumber >= startRow))
            {
                names.Add(cell.Value.ToString());
            }
        }

        return names;
    }

    static async Task DownloadImageAsync(string url, string filePath)
    {
        using (HttpClient client = new HttpClient())
        {
            Console.WriteLine($"Downloading {url}...");

            try
            {
                // Download image as a byte array
                byte[] imageData = await client.GetByteArrayAsync(url);

                // Save the image to the specified file path
                await File.WriteAllBytesAsync(filePath, imageData);

                Console.WriteLine($"Image saved to {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to download {url}: {ex.Message}");
            }
        }
    }
}
