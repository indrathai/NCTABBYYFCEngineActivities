using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

namespace NCTABBYYActivities
{
    [DisplayName("Split Into One Page")]
    public class PDFSplitIntoOnePage : CodeActivity
    {
        [Category("From")]
        [RequiredArgument]
        [DisplayName("Source PDF")]
        public InArgument<String> SourcePDF { get; set; }

        [Category("To")]
        [RequiredArgument]
        [DisplayName("Destination PDF")]
        public InArgument<String> DestinationPDF { get; set; }

        [Category("Output")]
        [DisplayName("Message")]
        public OutArgument<String> Output { get; set; }

        [Category("Output")]
        [DisplayName("Splitted PDF List")]
        public OutArgument<List<String>> SplittedPDFList { get; set; }

        [Category("Output")]
        [DisplayName("Wrong PDF List")]
        public OutArgument<List<String>> WrongPDFList { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            Console.WriteLine("Split PDF to be one page is begin ...");
            string message = "";
            int splitted = 0;
            List<String> newPDFList = new List<String>();
            List<String> wrongPDFList = new List<String>();
            try
            {
                string sourceFolder = SourcePDF.Get(context);

                string destinationPath = DestinationPDF.Get(context);
                Console.WriteLine("destination folder: " + destinationPath);
                if (!Directory.Exists(destinationPath))
                {
                    Directory.CreateDirectory(destinationPath);
                }

                int interval = 1;
                int pageNameSuffix = 0;

                string[] TotalFiles = Directory.GetFiles(sourceFolder, "*.pdf", SearchOption.AllDirectories);
                if (TotalFiles.Length == 0)
                {
                    Console.WriteLine("PDF Files *.pdf is not found");
                    throw new Exception("PDF Files *.pdf is not found");
                }

                foreach (string filename in TotalFiles)
                {
                    // Intialize a new PdfReader instance with the contents of the source Pdf file:  
                    splitted = 0;
                    Console.WriteLine("pdf: " + filename);
                    PdfReader reader = new PdfReader(filename);
                    if (reader.NumberOfPages < 2)
                    {
                        wrongPDFList.Add(filename);
                        Console.WriteLine("file " + filename + " only have one page");
                        continue;
                    }
                    string pdfFileName = Path.GetFileNameWithoutExtension(filename);
                    Console.WriteLine("total pdf in one page" + TotalFiles.Length);
                    for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber += interval)
                    {
                        splitted++;
                        pageNameSuffix++;
                        string newPdfFileName = string.Format(pdfFileName + "_{0}", pageNumber);

                        newPDFList.Add(destinationPath + "\\" + newPdfFileName + ".pdf");
                        this.SplitAndSaveInterval(filename, destinationPath, pageNumber, interval, newPdfFileName);
                        Console.WriteLine("pdf: " + pdfFileName + " is splitted in the " + destinationPath);
                    }

                    File.SetAttributes(filename, FileAttributes.System);
                    Console.WriteLine("delete pdf file in the " + sourceFolder);

                    //File.Copy(filename, filename + ".removed");
                    //File.Delete(filename);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                message = String.Format("PDF spliited is {0}", splitted);
                Console.WriteLine(message);

                Output.Set(context, message);
                SplittedPDFList.Set(context, newPDFList);
                WrongPDFList.Set(context, wrongPDFList);

                Console.WriteLine("Split PDF to be one page is end ...");
            }
        }

        private void SplitAndSaveInterval(string pdfFilePath, string outputPath, int startPage, int interval, string pdfFileName)
        {
            using (PdfReader reader = new PdfReader(pdfFilePath))
            {
                Document document = new Document();
                PdfCopy copy = new PdfCopy(document, new FileStream(outputPath + "\\" + pdfFileName + ".pdf", FileMode.Create));
                document.Open();
                for (int pagenumber = startPage; pagenumber < (startPage + interval); pagenumber++)
                {
                    if (reader.NumberOfPages >= pagenumber)
                    {
                        copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                    }
                    else
                    {
                        break;
                    }

                }
                reader.Close();
                document.Close();
            }
        }
    }
}
