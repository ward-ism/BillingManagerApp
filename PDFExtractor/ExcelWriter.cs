using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using PDFExtractor.Models;

public class ExcelWriter
{
    public void WriteInvoicesToExcel(List<InvoiceModel> invoices)
    {
        // Create a new Excel application
        var excelApp = new Application
        {
            Visible = true // Make the application visible
        };

        // Add a new workbook
        var workbook = excelApp.Workbooks.Add();
        var worksheet = (Worksheet)workbook.Worksheets[1];

        // Write headers
        worksheet.Cells[1, 1] = "PRO";
        worksheet.Cells[1, 2] = "Pieces";
        worksheet.Cells[1, 3] = "Date";
        worksheet.Cells[1, 4] = "LoadID";

        // Write invoice data
        int row = 2; // Start from the second row
        foreach (var invoice in invoices)
        {
            worksheet.Cells[row, 1] = invoice.ProNumber;
            worksheet.Cells[row, 2] = invoice.NumberOfPieces;
            worksheet.Cells[row, 3] = invoice.InvoiceDate.ToShortDateString();
            worksheet.Cells[row, 4] = invoice.LoadID;
            row++;
        }
    }
}
