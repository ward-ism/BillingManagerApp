using System;
using System.IO;
using System.Text.RegularExpressions;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using PDFExtractor.Models;

public class PdfFieldParser
{
    private const string KeyString = "HUNTINGTON PARK";

    public InvoiceModel ProcessPdf(string pdfPath)
    {
        try
        {
            // Extract text directly from the PDF
            string fullText = ExtractTextFromPdf(pdfPath);

            // Count occurrences of "Huntington Park"
            int keyStringCount = CountOccurrences(fullText, KeyString);

            // Only process PDFs where "Huntington Park" appears 4 or more times
            if (keyStringCount >= 4)
            {
                // Extract the details (PRO#, Date, Pieces, LoadID) from the full text
                return ExtractInvoiceDetails(fullText);
            }
            else
            {
                Console.WriteLine($"PDF skipped: 'Huntington Park' appeared {keyStringCount} times, which is less than 4.");
                return null;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing {pdfPath}: {ex.Message}");
            return null;
        }
    }

    // Extract all text from the PDF
    private string ExtractTextFromPdf(string pdfPath)
    {
        StringWriter textWriter = new StringWriter();

        using (PdfReader pdfReader = new PdfReader(pdfPath))
        {
            PdfDocument pdfDocument = new PdfDocument(pdfReader);

            // Iterate through each page and extract text
            for (int page = 1; page <= pdfDocument.GetNumberOfPages(); page++)
            {
                string pageText = PdfTextExtractor.GetTextFromPage(pdfDocument.GetPage(page));
                textWriter.WriteLine(pageText);
            }
        }

        return textWriter.ToString();
    }

    // Count occurrences of the key string in the extracted text
    private int CountOccurrences(string text, string keyString)
    {
        int count = 0;
        int index = 0;

        while ((index = text.IndexOf(keyString, index, StringComparison.OrdinalIgnoreCase)) != -1)
        {
            count++;
            index += keyString.Length;
        }

        return count;
    }

    // Extract details like PRO#, Date, Number of Pieces, and LoadID from the full text
    private InvoiceModel ExtractInvoiceDetails(string fullText)
    {
        string proNumber = ExtractProNumber(fullText);
        DateTime invoiceDate = ExtractInvoiceDate(fullText);
        int numberOfPieces = ExtractNumberOfPieces(fullText);
        string loadID = ExtractLoadID(fullText);

        return new InvoiceModel(proNumber, invoiceDate, numberOfPieces, loadID);
    }

    // Extract PRO# (assuming it's the first 9-digit number)
    private string ExtractProNumber(string text)
    {
        var match = Regex.Match(text, @"\b\d{9}\b");
        return match.Success ? match.Value : "Not found";
    }

    // Extract Invoice Date (find the earliest date in the text)
    private DateTime ExtractInvoiceDate(string text)
    {
        var matches = Regex.Matches(text, @"\b\d{2}/\d{2}/\d{4}\b");
        DateTime earliestDate = DateTime.MaxValue;

        foreach (Match match in matches)
        {
            if (DateTime.TryParse(match.Value, out DateTime parsedDate) && parsedDate < earliestDate)
            {
                earliestDate = parsedDate;
            }
        }

        return earliestDate == DateTime.MaxValue ? DateTime.MinValue : earliestDate;
    }

    // Extract Number of Pieces (assume $10 per piece, find the first monetary value and divide by 10)
    private int ExtractNumberOfPieces(string text)
    {
        var match = Regex.Match(text, @"\$\d+(\.\d{2})?");
        if (match.Success && decimal.TryParse(match.Value.Trim('$'), out decimal amount))
        {
            return (int)(amount / 10);
        }
        return 0;
    }

    // Extract LoadID (find the "GP" number followed by semicolon)
    private string ExtractLoadID(string text)
    {
        var match = Regex.Match(text, @"\bGP\d+;");
        return match.Success ? match.Value.TrimEnd(';') : "Not found";
    }
}
