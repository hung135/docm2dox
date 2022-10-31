﻿using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

class Program {

    static void Main(string[] args) {
        string filename = @"C:\Users\Public\Documents\WithMacros.docm";
        ConvertDOCMtoDOCX(filename);
    }

    public static void ConvertDOCMtoDOCX(string fileName)
    {
        bool fileChanged = false;

        using (WordprocessingDocument document = 
            WordprocessingDocument.Open(fileName, true))
        {
            // Access the main document part.
            var docPart = document.MainDocumentPart;
            
            // Look for the vbaProject part. If it is there, delete it.
            var vbaPart = docPart.VbaProjectPart;
            if (vbaPart != null)
            {
                // Delete the vbaProject part and then save the document.
                docPart.DeletePart(vbaPart);
                docPart.Document.Save();

                // Change the document type to
                // not macro-enabled.
                document.ChangeDocumentType(
                    WordprocessingDocumentType.Document);

                // Track that the document has been changed.
                fileChanged = true;
            }
        }

        // If anything goes wrong in this file handling,
        // the code will raise an exception back to the caller.
        if (fileChanged)
        {
            // Create the new .docx filename.
            var newFileName = Path.ChangeExtension(fileName, ".docx");
            
            // If it already exists, it will be deleted!
            if (File.Exists(newFileName))
            {
                File.Delete(newFileName);
            }

            // Rename the file.
            File.Move(fileName, newFileName);
        }
    }
}