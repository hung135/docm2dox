using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO.Packaging;

class Program {

    static void Main(string[] args) {
        //string filename = @"C:\Users\Public\Documents\WithMacros.docm";
        string filename = @"./example.docm";
        if (args.Length>0)
        {

        Console.Clear();
        Console.WriteLine("Start");
        Console.WriteLine(args[0]);
        filename=args[0];
        Console.WriteLine(filename);
        ConvertDOCMtoDOCX(filename);
        Console.WriteLine("End");
        }
        else{
            Console.WriteLine("Please provide path to DOCM file");
        }

                   
    }

    public static void ConvertDOCMtoDOCX(string fileName)
    {
        bool fileChanged = false;

        using (WordprocessingDocument document = 
            WordprocessingDocument.Open(fileName, true))
        {
            // Access the main document part.
            var docPart = document.MainDocumentPart;
            Console.WriteLine("Inside 1");
            // Look for the vbaProject part. If it is there, delete it.
            var vbaPart = docPart.VbaProjectPart;
            Console.WriteLine(vbaPart);
            if (vbaPart != null)
            {
                Console.WriteLine("Inside 2");
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
            else 
            {
            Console.WriteLine("No VBA found");
            }
        }

        // If anything goes wrong in this file handling,
        // the code will raise an exception back to the caller.
        if (fileChanged)
        {
            // Create the new .docx filename.
            var newFileName = Path.ChangeExtension(fileName, ".docx");
            Console.Clear();
            Console.WriteLine("Standard Numeric Format Specifiers");
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