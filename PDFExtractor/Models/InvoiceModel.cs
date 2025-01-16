using System;

namespace PDFExtractor.Models
{
    public class InvoiceModel
    {
        public string ProNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public int NumberOfPieces { get; set; }
        public string LoadID { get; set; } 
        
        public InvoiceModel(string proNumber, DateTime invoiceDate, int numberOfPieces, string loadID)
        {
            ProNumber = proNumber;
            InvoiceDate = invoiceDate;
            NumberOfPieces = numberOfPieces;
            LoadID = loadID;
        }
    }
}
