using System;
using System.IO;
using System.Text.RegularExpressions;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using PDFExtractor.Models;

public class PdfFieldParser
{
    //set keystring
    private const string KeyString = "HUNTINGTON PARK";

    public InvoiceModel ProcessPdf(string pdfPath)
    {
        try
        {
            // extract text
            string fullText = ExtractTextFromPdf(pdfPath);

            // count occurrences of keystring
            int keyStringCount = CountOccurrences(fullText, KeyString);

            // instruct to process PDFs where keystring appears X times
            if (keyStringCount >= 4)
            {
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

    // extract all text from PDF
    private string ExtractTextFromPdf(string pdfPath)
    {
        StringWriter textWriter = new StringWriter();

        using (PdfReader pdfReader = new PdfReader(pdfPath))
        {
            PdfDocument pdfDocument = new PdfDocument(pdfReader);

            // iterate through
            for (int page = 1; page <= pdfDocument.GetNumberOfPages(); page++)
            {
                string pageText = PdfTextExtractor.GetTextFromPage(pdfDocument.GetPage(page));
                textWriter.WriteLine(pageText);
            }
        }

        return textWriter.ToString();
    }

    // count occurrences of the key string
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

    // extract the details (PRO#, Date, Pieces, LoadID) for InvoiceModel
    private InvoiceModel ExtractInvoiceDetails(string fullText)
    {
        string proNumber = ExtractProNumber(fullText);
        DateTime invoiceDate = ExtractInvoiceDate(fullText);
        int numberOfPieces = ExtractNumberOfPieces(fullText);
        string loadID = ExtractLoadID(fullText);

        return new InvoiceModel(proNumber, invoiceDate, numberOfPieces, loadID);
    }

    // extract PRO# (assumes first 9-digit number)
    private string ExtractProNumber(string text)
    {
        var match = Regex.Match(text, @"\b\d{9}\b");
        return match.Success ? match.Value : "Not found";
    }

    // extract Date (earliest date in the text)
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

    // extract pieces (assume $10 per piece, find the first monetary value and divide by 10)
    private int ExtractNumberOfPieces(string text)
    {
        var match = Regex.Match(text, @"\$\d+(\.\d{2})?");
        if (match.Success && decimal.TryParse(match.Value.Trim('$'), out decimal amount))
        {
            // check if the amount is divisible by 10 - hoping to catch bad accessorials
            if (amount % 10 != 0)
            {
                throw new InvalidOperationException("check invoice");
            }

            return (int)(amount / 10);
        }
        return 0;
    }


    // extract LoadID (GP number followed by semicolon)
    private string ExtractLoadID(string text)
    {
        var match = Regex.Match(text, @"\bGP\d+;");
        return match.Success ? match.Value.TrimEnd(';') : "Not found";
    }
}
