using PDFExtractor.Models;
using System;
using System.Collections.Generic;
using System.IO;

namespace PDFExtractorApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //directories
            string tempDirectory = Path.Combine(Path.GetTempPath(), "ExtractedPDFs"); // temp storage for first-page PDFs
            string subfolderName = "WT Invoices"; // subfolder name in outlook

            // check directories exist
            if (!Directory.Exists(tempDirectory))
            {
                Directory.CreateDirectory(tempDirectory);
            }

            OutlookPDFExtractor extractor = new OutlookPDFExtractor(tempDirectory, subfolderName);
            PdfFieldParser parser = new PdfFieldParser();

            try
            {
                Console.WriteLine("Starting PDF extraction and processing...");

                // extract PDFs and save  the first page to temp fodler
                extractor.ExtractPdfAttachments();

                // filter PDFs from the temp folder
                List<InvoiceModel> invoices = new List<InvoiceModel>();
                foreach (string pdfPath in Directory.GetFiles(tempDirectory, "*.pdf"))
                {
                    Console.WriteLine($"Parsing and filtering: {Path.GetFileName(pdfPath)}");
                    var invoice = parser.ProcessPdf(pdfPath);
                    if (invoice != null) 
                    {
                        invoices.Add(invoice);
                        Console.WriteLine($"Invoice Details: PRO# {invoice.ProNumber}, Date: {invoice.InvoiceDate.ToShortDateString()}, Pieces: {invoice.NumberOfPieces}, LoadID: {invoice.LoadID}");
                    }
                }

                //sort by date
                invoices.Sort((a, b) => a.InvoiceDate.CompareTo(b.InvoiceDate));

                // write to Excel
                ExcelWriter excelWriter = new ExcelWriter();
                excelWriter.WriteInvoicesToExcel(invoices);

                Console.WriteLine("Processing completed successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
