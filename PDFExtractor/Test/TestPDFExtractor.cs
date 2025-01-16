using System;
using System.IO;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace PDFExtractorApp.Test
{
    public class TestPdfExtractor
    {
        // Main method to run the test
        public static void Main(string[] args)
        {
            RunPdfTest();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        // Method to process PDFs from the "Test" folder and save the extracted text
        public static void RunPdfTest()
        {
            string testFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "Test");
            string outputFolderPath = Path.Combine(testFolderPath, "ExtractedText");

            // Log the input folder and output folder paths
            Console.WriteLine($"Test Folder Path: {testFolderPath}");
            Console.WriteLine($"Output Folder Path: {outputFolderPath}");

            if (!Directory.Exists(outputFolderPath))
            {
                Directory.CreateDirectory(outputFolderPath);
            }

            var pdfFiles = Directory.GetFiles(testFolderPath, "*.pdf");

            // Log the files found
            Console.WriteLine($"Found {pdfFiles.Length} PDF files to process.");

            foreach (var pdfFile in pdfFiles)
            {
                try
                {
                    Console.WriteLine($"Processing PDF: {Path.GetFileName(pdfFile)}");

                    // Extract text from the PDF
                    string extractedText = ExtractTextFromPdf(pdfFile);

                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        Console.WriteLine($"No text extracted from {Path.GetFileName(pdfFile)}.");
                    }
                    else
                    {
                        // Print the first 100 characters of extracted text for verification
                        Console.WriteLine($"Extracted Text: {extractedText.Substring(0, Math.Min(100, extractedText.Length))}...");

                        // Save the extracted text to a text file in the ExtractedText folder
                        string outputTextFile = Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(pdfFile) + "_extracted.txt");

                        // Log the file path where the extracted text is being saved
                        Console.WriteLine($"Saving extracted text to: {outputTextFile}");
                        File.WriteAllText(outputTextFile, extractedText);

                        Console.WriteLine($"Extracted text saved to: {outputTextFile}");
                    }
                }
                catch (Exception ex)
                {
                    // Catch any error during the process
                    Console.WriteLine($"Error processing {Path.GetFileName(pdfFile)}: {ex.Message}");
                }
            }

            Console.WriteLine("TestPDF extraction completed.");
        }

        // Method to extract text from the given PDF
        private static string ExtractTextFromPdf(string pdfPath)
        {
            string text = string.Empty;

            try
            {
                using (PdfReader reader = new PdfReader(pdfPath))
                using (PdfDocument pdfDoc = new PdfDocument(reader))
                {
                    // Extract text from each page
                    for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                    {
                        var strategy = new LocationTextExtractionStrategy();
                        text += PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting text from PDF: {ex.Message}");
            }

            return text;
        }
    }
}
