using Microsoft.Office.Interop.Outlook;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using System;
using System.IO;

public class OutlookPDFExtractor
{
    private string _outputDirectory;
    private string _subfolderName;

    public OutlookPDFExtractor(string outputDirectory, string subfolderName)
    {
        _outputDirectory = outputDirectory;
        _subfolderName = subfolderName;

        // Ensure the directory exists
        if (!Directory.Exists(_outputDirectory))
        {
            Directory.CreateDirectory(_outputDirectory);
        }
    }

    private MAPIFolder GetFolder(NameSpace outlookNamespace, string folderPath)
    {
        string[] folderNames = folderPath.Split('\\'); // for nested folders
        MAPIFolder currentFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

        foreach (string folderName in folderNames)
        {
            bool folderFound = false;

            foreach (MAPIFolder subFolder in currentFolder.Folders)
            {
                if (subFolder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                {
                    currentFolder = subFolder;
                    folderFound = true;
                    break;
                }
            }

            if (!folderFound)
            {
                throw new System.Exception($"Subfolder '{folderName}' not found.");
            }
        }

        return currentFolder;
    }

    public void ExtractPdfAttachments()
    {
        Application outlookApp = new Application();
        NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

        try
        {
            
            MAPIFolder targetFolder = GetFolder(outlookNamespace, _subfolderName);

            // process emails in specified folder
            foreach (object item in targetFolder.Items)
            {
                if (item is MailItem mailItem && mailItem.Attachments.Count > 0)
                {
                    foreach (Attachment attachment in mailItem.Attachments)
                    {
                        if (Path.GetExtension(attachment.FileName).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                        {
                            string tempFilePath = Path.Combine(Path.GetTempPath(), attachment.FileName);

                            // save attachment to temp
                            attachment.SaveAsFile(tempFilePath);

                            // copy first page
                            SaveFirstPage(tempFilePath);

                            // delete temp
                            File.Delete(tempFilePath);
                        }
                    }
                }
            }
        }
        catch (System.Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }


    private void SaveFirstPage(string pdfPath)
    {
        string outputFileName = Path.Combine(_outputDirectory, Path.GetFileNameWithoutExtension(pdfPath) + "_Page1.pdf");

        // open pdf in import mode
        using (PdfDocument originalDocument = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import))
        {
            if (originalDocument.PageCount > 0)
            {
                // create new odf
                using (PdfDocument newDocument = new PdfDocument())
                {
                    // import and add first page to new doc
                    newDocument.AddPage(originalDocument.Pages[0]);

                    // save
                    newDocument.Save(outputFileName);

                    Console.WriteLine($"First page saved: {outputFileName}");
                }
            }
            else
            {
                Console.WriteLine($"The document {Path.GetFileName(pdfPath)} has no pages.");
            }
        }
    }

}
