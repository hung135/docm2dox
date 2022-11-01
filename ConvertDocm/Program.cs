using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO.Packaging;

class Program {

    static void Main(string[] args) {
        //string filename = @"C:\Users\Public\Documents\WithMacros.docm";
        string filename = "";
        string outfile = "";
        if (args.Length>1)
        {

        Console.Clear();
        Console.WriteLine("Start");
        filename=args[0];
        outfile=args[1];
        filename = Path.GetFullPath(filename);
        outfile = Path.GetFullPath(outfile);
        Console.WriteLine(filename);
        Console.WriteLine(outfile);
        ConvertDOCMtoDOCX(filename,outfile);
        Console.WriteLine("End");
        }
        else{
            Console.WriteLine("Please provide path to DOCM file and output file name");
        }

                   
    }

    public static void ConvertDOCMtoDOCX(string fileName,string outFile)
    {
        bool fileChanged = false;
        
        //backup the file
        try
        {
            // Will not overwrite if the destination file already exists.
            Console.WriteLine(outFile);
            File.Copy(fileName, outFile,true);
            
             
        }

        // Catch exception if the file was already copied.
        catch (IOException copyError)
        { 
            Console.WriteLine(copyError.Message);
             
        }

        Console.WriteLine(fileName);
        using (WordprocessingDocument document = 
            WordprocessingDocument.Open( outFile, true))

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
            var newFileName = Path.ChangeExtension(outFile, ".docx");
            Console.Clear();
            Console.WriteLine("File Was Changed");
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