using FCEngine;
using NCTABBYYActivities.config;
using NCTABBYYActivities.model;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace NCTABBYYActivities
{
    [DisplayName("ABBYY Recognition")]
    public class ABBYYRecognition : CodeActivity
    {
        [DllImport(FConfig.DllPath, CharSet = CharSet.Unicode), PreserveSig]
        internal static extern int InitializeEngine(string ProjectId, out IEngine engine);

        [DllImport(FConfig.DllPath, CharSet = CharSet.Unicode), PreserveSig]
        internal static extern int DeinitializeEngine();

        private IEngine engine;
        private IFlexiCaptureProcessor processor;

        [Category("ABBYY")]
        [RequiredArgument]
        [DisplayName("Document Project Id")]
        [DispId(1)]
        public InArgument<string> DocumentProjectId { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("OCR Template Folder")]
        [DispId(2)]
        public InArgument<string> OCRTemplateFolder { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Confidence Level Limitation(%)")]
        [DispId(2)]
        [DefaultValue(0)]
        public InArgument<int> ConfidenceLevelLimitation { get; set; }

        [Category("Source")]
        [RequiredArgument]
        [DisplayName("Source Folder(*.pdf)")]
        [DispId(3)]
        public InArgument<string> SourceFolder { get; set; }

        [Category("Destination")]
        [RequiredArgument]
        [DisplayName("Export Folder")]
        [DispId(4)]
        public InArgument<string> ExportFolder { get; set; }

        [Category("Destination")]
        [RequiredArgument]
        [DisplayName("Recognized Folder")]
        [DispId(4)]
        public InArgument<string> RecognizedFolder { get; set; }

        [Category("Destination")]
        [RequiredArgument]
        [DisplayName("Not Confidence Folder")]
        [DispId(4)]
        public InArgument<string> NotConfidenceFolder { get; set; }

        [Category("Destination")]
        [RequiredArgument]
        [DisplayName("Reject Folder")]
        [DispId(4)]
        public InArgument<string> RejectFolder { get; set; }

        [Category("Output")]
        [DisplayName("Is Recognized?")]
        [DispId(5)]
        public OutArgument<bool> IsRecognized { get; set; }

        [Category("Output")]
        [DisplayName("Log Message")]
        [DispId(6)]
        public OutArgument<List<LogMessage>> Messages { get; set; }

        [Category("Output")]
        [DisplayName("Non Invoice List")]
        [DispId(6)]
        public OutArgument<List<string>> NonInvoiceList { get; set; }

        [Category("Output")]
        [DisplayName("Non Confidence List")]
        [DispId(6)]
        public OutArgument<List<string>> NotConfidenceList { get; set; }

        [Category("Output")]
        [DisplayName("Total Exported")]
        [DispId(6)]
        public OutArgument<Int32> TotalExported { get; set; }

        [Category("Output")]
        [DisplayName("Total Error")]
        [DispId(6)]
        public OutArgument<Int32> TotalError { get; set; }

        [Category("Output")]
        [DisplayName("Total Not Confidence")]
        [DispId(6)]
        public OutArgument<Int32> TotalNotConfidence { get; set; }

        private int confidenceLevel = 0;
        private int confidenceHeaderLevel = 0;
        private int confidenceDetailLevel = 0;
        private int totalConfidenceLevelHD = 0;

        protected override void Execute(CodeActivityContext context)
        {

            string message = "";
            List<LogMessage> logList = new List<LogMessage>();
            List<string> nonInvoices = new List<string>();
            List<string> notConfidenceList = new List<string>();

            message = "Loading FlexiCapture Engine for Recognition ...";
            Console.WriteLine(message);
            logList.Add(new LogMessage(message, LogType.Information));

            string documentProjectId = DocumentProjectId.Get(context);
            message = "Get Document Project ID: " + documentProjectId;
            Console.WriteLine(message);
            logList.Add(new LogMessage(message, LogType.Information));

            int confidenceLevelLimitation = ConfidenceLevelLimitation.Get(context);
            message = string.Format("Confidence Level Limitation: {0}", confidenceLevelLimitation);
            Console.WriteLine(message);
            logList.Add(new LogMessage(message, LogType.Information));

            string sourceFolder = SourceFolder.Get(context);
            message = string.Format("Get PDF folder: {0}", sourceFolder);
            Console.WriteLine(message);
            logList.Add(new LogMessage(message, LogType.Information));

            string ocrTemplate = OCRTemplateFolder.Get(context);
            message = string.Format("OCR Template Folder: {0}", ocrTemplate);
            Console.WriteLine(message);
            logList.Add(new LogMessage(message, LogType.Information));

            string exportFolder = ExportFolder.Get(context);
            message = string.Format("Get Export Folder: {0}", exportFolder);
            Console.WriteLine(message);
            logList.Add(new LogMessage(message, LogType.Information));
            if (!Directory.Exists(exportFolder))
            {
                Directory.CreateDirectory(exportFolder);
                message = string.Format("Folder {0} is created completely", exportFolder);
                Console.WriteLine(message);
                logList.Add(new LogMessage(message, LogType.Information));
            }

            string recognizedFolder = RecognizedFolder.Get(context);
            Console.WriteLine("Get Recognize Folder: " + recognizedFolder);
            logList.Add(new LogMessage("Recognize Folder is:  " + recognizedFolder, LogType.Information));
            if (!Directory.Exists(recognizedFolder))
            {
                Directory.CreateDirectory(recognizedFolder);
                message = string.Format("Folder {0} is created completely", recognizedFolder);
                Console.WriteLine(message);
                logList.Add(new LogMessage(message, LogType.Information));
            }

            string notConfidenceFolder = NotConfidenceFolder.Get(context);
            Console.WriteLine("Get Not Confidence Folder: " + notConfidenceFolder);
            logList.Add(new LogMessage("Not Confidence Folder is:  " + notConfidenceFolder, LogType.Information));
            if (!Directory.Exists(notConfidenceFolder))
            {
                Directory.CreateDirectory(notConfidenceFolder);
                message = string.Format("Folder {0} is created completely", notConfidenceFolder);
                Console.WriteLine(message);
                logList.Add(new LogMessage(message, LogType.Information));
            }

            string rejectFolder = RejectFolder.Get(context);
            Console.WriteLine("Get Reject Folder: " + rejectFolder);
            logList.Add(new LogMessage("Reject Folder is:  " + rejectFolder, LogType.Information));
            if (!Directory.Exists(rejectFolder))
            {
                Directory.CreateDirectory(rejectFolder);
                message = string.Format("Folder {0} is created completely", rejectFolder);
                Console.WriteLine(message);
                logList.Add(new LogMessage(message, LogType.Information));
            }

            engine = LoadEngine(documentProjectId);

            Boolean isRecognized = false;
            bool isValidInvoice = false;
            int count = 0;
            int noError = 0;
            int noSuccess = 0;
            int noNotConfidence = 0;
            try
            {
                message = "Creating and configuring the FlexiCapture Processor...";
                Console.WriteLine(message);
                logList.Add(new LogMessage(message, LogType.Information));

                processor = engine.CreateFlexiCaptureProcessor();

                message = "Adding Document Definition to process...";
                Console.WriteLine(message);
                logList.Add(new LogMessage(message, LogType.Information));
                string[] TotalOCRFiles = Directory.GetFiles(ocrTemplate, "*.fcdot", SearchOption.AllDirectories);
                if (TotalOCRFiles.Length == 0)
                {
                    message = string.Format("OCR Template *.fcdot is not found in the folder {0}", ocrTemplate);
                    Console.WriteLine(message);
                    logList.Add(new LogMessage(message, LogType.Error));
                    throw new Exception(message);
                }
                foreach (string ocr in TotalOCRFiles)
                {
                    processor.AddDocumentDefinitionFile(ocr);
                    message = string.Format("OCR Temhplate {0} is added", ocr);
                    Console.WriteLine(message);
                    logList.Add(new LogMessage(message, LogType.Information));
                }

                message = string.Format("Adding images to process...");
                Console.WriteLine(message);
                logList.Add(new LogMessage(message, LogType.Information));
                string[] TotalFiles = Directory.GetFiles(sourceFolder, "*.pdf", SearchOption.AllDirectories);
                if (TotalFiles.Length == 0)
                {
                    message = string.Format("PDF Files *.pdf is not found in the folder {0}.", sourceFolder);
                    Console.WriteLine(message);
                    logList.Add(new LogMessage(message, LogType.Error));
                    throw new Exception(message);
                }
                foreach (string pdfFile in TotalFiles)
                {
                    processor.AddImageFile(pdfFile);
                    message = string.Format("PDF Files {0} is added", pdfFile);
                    Console.WriteLine(message);
                    logList.Add(new LogMessage(message, LogType.Information));
                }

                message = "Recognizing the images and exporting the results...";
                Console.WriteLine(message);
                logList.Add(new LogMessage(message, LogType.Information));

                while (true)
                {
                    // Recognize next document
                    confidenceLevel = 0;
                    confidenceDetailLevel = 0;
                    confidenceHeaderLevel = 0;
                    totalConfidenceLevelHD = 0;
                    IDocument document = processor.RecognizeNextDocument();
                    if (document == null)
                    {
                        IProcessingError error = processor.GetLastProcessingError();
                        if (error != null)
                        {
                            // Processing error
                            message = string.Format("processing error because of {0}.", error.MessageText());
                            Console.WriteLine(message);
                            logList.Add(new LogMessage(message, LogType.Error));
                            continue;
                        }
                        else
                        {
                            // No more images
                            message = string.Format("all PDF Files has been executed or no PDF file in the folder {0}", sourceFolder);
                            Console.WriteLine(message);
                            logList.Add(new LogMessage(message, LogType.Information));
                            break;
                        }
                    }
                    else if (document.DocumentDefinition == null)
                    {
                        // Couldn't find matching template for the image. In this sample this is an error.
                        // In other scenarios this might be normal
                        message = string.Format("PDF file is not matched with existing OCR Templates.");
                        Console.WriteLine(message);
                        logList.Add(new LogMessage(message, LogType.Error));
                        //string tempPage = document.Pages[0].OriginalImagePath;
                        //if(tempPage != null)
                        //{
                        //    string movefile = Path.GetFileName(tempPage);
                        //    string tempfilename = Path.GetFileNameWithoutExtension(tempPage);

                        //    logList.Add(new LogMessage("Move to Reject Folder", LogType.Information));
                        //    if (File.Exists(tempPage))
                        //    {

                        //        MoveFileToDestinationFolder(sourceFolder, rejectFolder, tempPage, tempfilename);
                        //    }
                        //    else
                        //    {
                        //        MoveFileToDestinationFolder(notConfidenceFolder, rejectFolder, notConfidenceFolder+"\\"+movefile, tempfilename);
                        //    }

                        //}

                        continue;
                    }

                    string originalPath = document.Pages[0].OriginalImagePath;
                    string file = Path.GetFileName(originalPath);
                    string filenamewithoutextention = Path.GetFileNameWithoutExtension(originalPath);

                    message = string.Format("Recognizing pdf {0} is started.", originalPath);
                    Console.WriteLine(message);
                    logList.Add(new LogMessage(message, LogType.Information));

                    //set confident level and status
                    message = string.Format("Extracting data from pdf {0} is started", originalPath);
                    Console.WriteLine(message);
                    logList.Add(new LogMessage(message, LogType.Information));
                    for (int i = 0; i < document.Sections.Count; i++)
                    { // extracing 
                        var section = document.Sections[i];
                        if (object.ReferenceEquals(section, null))
                        {
                            continue;
                        }
                        

               
                            for (int d = 0; d < section.Children.Count; d++)
                            {
                                var child = section.Children[d];
                                if (object.ReferenceEquals(child, null))
                                    continue;

                                var field = ((IField)child);

                                message = string.Format("Extracting column {0} = {1}", field.Name, field.Value.AsText);
                                Console.WriteLine(message);

                                if (field.Name.ToUpper().Trim() == "INV_CONFIDENCE_LEVEL")
                                {
                                    var value = TextFieldHelper.GetConfidenLevel(engine, document);
                                    totalConfidenceLevelHD = value;
                                    var data = engine.CreateText(value.ToString(), null);
                                    field.Value.AsInteger = value;
                                    confidenceLevel = value;

                                    message = string.Format("Confidence level of {0} is {1}", file, value);
                                    Console.WriteLine(message);
                                    logList.Add(new LogMessage(message, LogType.Information));
                                }

                                if (field.Name.ToUpper().Trim() == "INV_STATUS")
                                {
                                    var value = "Recognized";
                                    var data = engine.CreateText(value, null);
                                    field.Value.AsText.Delete(0, field.Value.AsText.Length);
                                    field.Value.AsText.Insert(data, 0);
                                }


                                if (field.Name.ToUpper().Trim() == "FILE_NAME")
                                {
                                    var data = engine.CreateText(file, null);
                                    field.Value.AsText.Delete(0, field.Value.AsText.Length);
                                    field.Value.AsText.Insert(data, 0);
                                }
                            }
                    } //end extracting
                    isValidInvoice = true;
                    if (isValidInvoice)
                    {
                       
                        message = string.Format("Total Confidence is {0} ", totalConfidenceLevelHD);
                        Console.WriteLine(message);
                        logList.Add(new LogMessage(message, LogType.Information));
                        //check confidence level validation
                        if (totalConfidenceLevelHD >= confidenceLevelLimitation)
                        {
                            try
                            {
                                message = string.Format("Exporting process for pdf {0} is started ...", file);
                                Console.WriteLine(message);
                                logList.Add(new LogMessage(message, LogType.Information));

                                //IFileExportParams exportParams = engine.CreateFileExportParams();
                                //Console.WriteLine("XLS");
                                //exportParams.FileFormat = FileExportFormatEnum.FEF_XLS;


                                processor.ExportDocument(document, exportFolder);
                                Console.WriteLine("Exporting process is completed ...");

                                MoveFileToDestinationFolder(sourceFolder, exportFolder, originalPath, filenamewithoutextention);
                                Console.WriteLine(string.Format("Moving {0} to Export folder {1} is completed", file, exportFolder));

                                logList.Add(new LogMessage("Exporting process is ended ...", LogType.Information));
                                noSuccess++;
                            }
                            catch (Exception e)
                            {
                                noError++;
                                Console.WriteLine(string.Format("exporting is failed because of {0}.", e.Message));
                                logList.Add(new LogMessage(message, LogType.Error));

                                MoveFileToDestinationFolder(sourceFolder, rejectFolder, originalPath, filenamewithoutextention);
                                message = string.Format("Moving pdf {0} to Reject folder {1} is completed", file, rejectFolder);
                                Console.WriteLine(message);
                                logList.Add(new LogMessage(message, LogType.Information));
                                continue;
                            }
                        }
                        else
                        {
                            message = string.Format("Confidence of PDF {0} is {1} less than target confidence {2}", file, totalConfidenceLevelHD, confidenceLevelLimitation);
                            logList.Add(new LogMessage(message, LogType.Error));
                            notConfidenceList.Add(file);

                            logList.Add(new LogMessage(string.Format("Total number of not confidence is {0}", notConfidenceList.Count), LogType.Information));
                            MoveFileToDestinationFolder(sourceFolder, notConfidenceFolder, originalPath, filenamewithoutextention);
                            message = string.Format("Moving pdf {0} to Not Confidence folder {1} is completed", file, notConfidenceFolder);
                            Console.WriteLine(message);
                            logList.Add(new LogMessage(message, LogType.Information));
                            noNotConfidence++;
                        }
                    }
                    count++;
                }

                var msg = processor.GetLastProcessingError();
                if (msg != null)
                {
                    var msgError = string.Format("the processing error because of {0}.", msg.MessageText());
                    Console.WriteLine(msgError);
                    logList.Add(new LogMessage(msgError, LogType.Error));
                    noError++;
                }
                message = string.Format("No. of Not Confidence {0} and No. of Error {1} and No. of Exported to DB {2}", noNotConfidence, noError, noSuccess);
                logList.Add(new LogMessage(message, LogType.Information));
                isRecognized = true;
            }
            finally
            {
                UnloadEngine(ref engine);
                message = string.Format("Released FlexiCapture Engine for Recognition...");
                Console.WriteLine(message);
                logList.Add(new LogMessage(message, LogType.Information));

                IsRecognized.Set(context, isRecognized);
                NonInvoiceList.Set(context, nonInvoices);
                Messages.Set(context, logList);
                TotalExported.Set(context, noSuccess);
                TotalError.Set(context, noError);
                TotalNotConfidence.Set(context, noNotConfidence);
                NotConfidenceList.Set(context, notConfidenceList);
            }
        }

        protected IEngine LoadEngine(string projectId)
        {
            Console.WriteLine("Load ABBYY Engine ...");

            int hresult = InitializeEngine(projectId, out engine);
            Marshal.ThrowExceptionForHR(hresult);

            Console.WriteLine("ABBYY Engine is loaded");

            return engine;
        }

        protected void UnloadEngine(ref IEngine engine)
        {
            Console.WriteLine("un-load ABBYY Engine ...");
            int hresult = DeinitializeEngine();
            Marshal.ThrowExceptionForHR(hresult);
            engine = null;

            Console.WriteLine("ABBYY Engine is unloaded");
        }

        protected void MoveFileToDestinationFolder(string source, string destination, string pdfFile, string filename)
        {
            Console.WriteLine(string.Format("copy {0} to destination folder {1} ", pdfFile, destination));
            File.SetAttributes(pdfFile, FileAttributes.System);
            string tempFile = destination + "\\" + filename + ".pdf";
            File.Copy(pdfFile, tempFile, true);
            Console.WriteLine(string.Format("delete pdf {0} in the folder {1}", pdfFile, source));
            File.Delete(pdfFile);
        }
    }
}
