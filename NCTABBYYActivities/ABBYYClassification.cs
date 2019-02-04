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

namespace NCTABBYYActivities
{
    [DisplayName("ABBYY Classification")]
    public class ABBYYClassification : CodeActivity
    {
        [DllImport(FConfig.DllPath, CharSet = CharSet.Unicode), PreserveSig]
        internal static extern int InitializeEngine(string ProjectId, out IEngine engine);

        [DllImport(FConfig.DllPath, CharSet = CharSet.Unicode), PreserveSig]
        internal static extern int DeinitializeEngine();

        private IEngine engine = null;
        private IFlexiCaptureProcessor processor = null;

        [Category("ABBYY")]
        [RequiredArgument]
        [DisplayName("Document Project Id")]
        [DispId(1)]
        public InArgument<string> DocumentProjectId { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Classifier Template")]
        [DispId(2)]
        public InArgument<string> ClassifierTemplate { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Folder Image(*.pdf)")]
        [DispId(3)]
        public InArgument<string> FolderFile { get; set; }

        [Category("To")]
        [RequiredArgument]
        [DisplayName("Reject Folder")]
        [DispId(4)]
        public InArgument<string> RejectFolder { get; set; }

        [Category("To")]
        [RequiredArgument]
        [DisplayName("Classified Folder")]
        [DispId(5)]
        public InArgument<string> ClassifiedFolder { get; set; }

        [Category("Output")]
        [DispId(6)]
        public OutArgument<List<LogMessage>> LogMessages { get; set; }

        [Category("Output")]
        [DispId(6)]
        public OutArgument<bool> IsClassified { get; set; }

        [Category("Output")]
        [DispId(6)]
        public OutArgument<List<string>> AnnexPageList { get; set; }

        [Category("Output")]
        [DispId(6)]
        public OutArgument<List<string>> MissingInvoiceList { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            Console.WriteLine("Loading FlexiCapture Engine for Classification ...");
            bool isClassified = false;
            List<LogMessage> logList = new List<LogMessage>();
            List<string> filenameList = new List<string>();
            List<string> annexPages = new List<string>();
            List<string> missingInvoices = new List<string>();
            int totalClassified = 0;
            int totalReject = 0;
            string name = "";
            bool isInvoice = false;
            bool isPurchaseOrder = false;
            try
            {

                //get Document Project Id
                string documentProjectId = DocumentProjectId.Get(context);
                Console.WriteLine("Document Project ID: " + documentProjectId);
                logList.Add(new LogMessage("Document Project ID: " + documentProjectId, LogType.Information));

                //get classifierTemplate path
                string classifierTemplatePath = ClassifierTemplate.Get(context);
                Console.WriteLine("Classifier path: " + classifierTemplatePath);
                logList.Add(new LogMessage("Classifier path: " + classifierTemplatePath, LogType.Information));

                //get image path
                string sourceFolder = FolderFile.Get(context);
                Console.WriteLine("image path: " + sourceFolder);
                logList.Add(new LogMessage("image path: " + sourceFolder, LogType.Information));

                //get unknown folder
                string rejectFolder = RejectFolder.Get(context);
                Console.WriteLine("unknown folder: " + rejectFolder);
                logList.Add(new LogMessage("image path: " + sourceFolder, LogType.Information));
                if (!Directory.Exists(rejectFolder))
                {
                    Directory.CreateDirectory(rejectFolder);
                }

                //get classified Folder
                string classifiedFolder = ClassifiedFolder.Get(context);
                Console.WriteLine("Classified folder: " + classifiedFolder);
                logList.Add(new LogMessage("Classified folder: " + classifiedFolder, LogType.Information));
                if (!Directory.Exists(classifiedFolder))
                {
                    Directory.CreateDirectory(classifiedFolder);
                }

                Console.WriteLine("Adding images to process...");
                string[] files = Directory.GetFiles(sourceFolder, "*.pdf", SearchOption.AllDirectories);
                List<string> pdfFileList = new List<string>();
                if (files.Length == 0)
                {
                    Console.WriteLine("pdf Files *.pdf is not found");
                    logList.Add(new LogMessage("pdf Files *.pdf is not found", LogType.Error));
                    throw new Exception("pdf Files *.pdf is not found");
                }
                foreach (string pdf in files)
                {
                    Console.WriteLine("pdf: " + pdf);
                    logList.Add(new LogMessage("pdf: " + pdf, LogType.Information));
                    pdfFileList.Add(pdf);

                    string[] splitstr = pdf.Split('_');
                    string compare = splitstr[0];

                    if (pdf.Contains(compare))
                    {
                        filenameList.Add(compare);
                    }
                }

                Console.WriteLine("Added images to process...");

                List<string> invoiceList = new List<string>();
                List<string> poList = new List<string>();

                foreach (string obj in filenameList.Distinct())
                {
                    int total = getTotalPage(files, obj);
                    isInvoice = false;
                    isPurchaseOrder = false;
                    invoiceList.Clear();
                    poList.Clear();
                    for (int i = 1; i <= total; i++)
                    {
                        string pdfFile = obj + "_" + i + ".pdf";

                        name = this.DetermineDocumentType(documentProjectId, pdfFile, classifierTemplatePath);

                        Console.WriteLine("executing pdf: " + pdfFile + " as " + name);
                        logList.Add(new LogMessage("executing pdf: " + pdfFile, LogType.Information));

                        if (Object.Equals(name.ToUpper(), "INVOICE"))
                        {
                            invoiceList.Add(pdfFile);
                            isInvoice = true;
                            continue;
                        }

                        if (Object.Equals(name.ToUpper(), "PURCHASEORDER"))
                        {
                            poList.Add(pdfFile);
                            isPurchaseOrder = true;
                            continue;
                        }

                        if (name.Length == 0)
                        {
                            Console.WriteLine("annex page is found as " + pdfFile);
                            logList.Add(new LogMessage("annex page is found as " + pdfFile, LogType.Error));
                            annexPages.Add(obj);

                            Console.WriteLine("move to reject folder " + rejectFolder);
                            string filenamewithoutextention = Path.GetFileNameWithoutExtension(pdfFile);
                            logList.Add(new LogMessage("move to reject folder", LogType.Information));
                            MoveFileToDestinationFolder(sourceFolder, rejectFolder, pdfFile, pdfFile, filenamewithoutextention);
                            totalReject++;
                        }

                    }

                    if (isInvoice && isPurchaseOrder)
                    {
                        foreach (string pdfFile in invoiceList)
                        {
                            Console.WriteLine("move to classified folder " + classifiedFolder);
                            string filenamewithoutextention = Path.GetFileNameWithoutExtension(pdfFile);
                            logList.Add(new LogMessage("move to classified folder " + classifiedFolder, LogType.Information));
                            MoveFileToDestinationFolder(sourceFolder, classifiedFolder, pdfFile, pdfFile, filenamewithoutextention);
                            totalClassified++;
                        }

                        foreach (string pdfFile in poList)
                        {
                            Console.WriteLine("move to reject folder " + rejectFolder);
                            string filenamewithoutextention = Path.GetFileNameWithoutExtension(pdfFile);
                            logList.Add(new LogMessage("move to reject folder", LogType.Information));
                            MoveFileToDestinationFolder(sourceFolder, rejectFolder, pdfFile, pdfFile, filenamewithoutextention);
                            totalReject++;
                        }
                    }
                    else
                    {
                        missingInvoices.Add(obj);

                        foreach (string pdfFile in invoiceList)
                        {
                            Console.WriteLine("move to reject folder " + rejectFolder);
                            string filenamewithoutextention = Path.GetFileNameWithoutExtension(pdfFile);
                            logList.Add(new LogMessage("move to reject folder", LogType.Information));
                            MoveFileToDestinationFolder(sourceFolder, rejectFolder, pdfFile, pdfFile, filenamewithoutextention);
                            missingInvoices.Add(pdfFile);
                            totalReject++;
                        }

                        foreach (string pdfFile in poList)
                        {
                            Console.WriteLine("move to reject folder " + rejectFolder);
                            string filenamewithoutextention = Path.GetFileNameWithoutExtension(pdfFile);
                            logList.Add(new LogMessage("move to reject folder", LogType.Information));
                            MoveFileToDestinationFolder(sourceFolder, rejectFolder, pdfFile, pdfFile, filenamewithoutextention);
                            totalReject++;
                        }
                    }
                }

                //foreach (string pdfFile in pdfFileList)
                //{
                //    Console.WriteLine("file: " + pdfFile);
                //    name  = this.DetermineDocumentType(documentProjectId, pdfFile, classifierTemplatePath);
                //    string filenamewithoutextention = Path.GetFileNameWithoutExtension(pdfFile);

                //    Console.WriteLine("executing pdf: " + pdfFile);
                //    logList.Add(new LogMessage("executing pdf: " + pdfFile, LogType.Information));


                //    if (name.Length == 0)
                //    {
                //        Console.WriteLine("move to reject folder " + rejectFolder);
                //        logList.Add(new LogMessage("move to reject folder", LogType.Information));
                //        MoveFileToDestinationFolder(sourceFolder, rejectFolder, pdfFile, pdfFile, filenamewithoutextention);
                //        NoRejected++;
                //    }
                //    else
                //    {
                //        Console.WriteLine("move to classified folder "+classifiedFolder);
                //        logList.Add(new LogMessage("move to classified folder "+ classifiedFolder, LogType.Information));
                //        MoveFileToDestinationFolder(sourceFolder, classifiedFolder, pdfFile, pdfFile, filenamewithoutextention);
                //        NoClassified++;
                //    }

                //}

                isClassified = true;
                string message = String.Format("Total Number of classified Invoice is {0} and Rejected is {1}", totalClassified, totalReject);
                logList.Add(new LogMessage(message, LogType.Information));

                var msg = processor.GetLastProcessingError();
                if (msg != null)
                {
                    var msgError = msg.MessageText();
                    logList.Add(new LogMessage(msgError, LogType.Error));
                    Console.WriteLine(msgError);
                }

            }
#pragma warning disable CS0618 // Type or member is obsolete
            catch (ExecutionEngineException e)
#pragma warning restore CS0618 // Type or member is obsolete
            {
                isClassified = false;
                Console.WriteLine("error: " + e.Message);
                logList.Add(new LogMessage("throw exception because of " + e.Message, LogType.Error));
            }
            finally
            {
                Console.WriteLine("Released FlexiCapture Engine ...");
                UnloadEngine(ref engine);

                logList.Add(new LogMessage("Released FlexiCapture Engine ...", LogType.Information));
                IsClassified.Set(context, isClassified);
                LogMessages.Set(context, logList);
                AnnexPageList.Set(context, annexPages);
                MissingInvoiceList.Set(context, missingInvoices);
            }
        }

        private IEngine LoadEngine(string projectId)
        {
            Console.WriteLine("Load ABBYY Engine ...");

            IEngine engine = null;
            int hresult = InitializeEngine(projectId, out engine);
            Marshal.ThrowExceptionForHR(hresult);

            Console.WriteLine("ABBYY Engine is loaded");

            return engine;
        }

        private void UnloadEngine(ref IEngine engine)
        {
            Console.WriteLine("un-load ABBYY Engine ...");
            int hresult = DeinitializeEngine();
            Marshal.ThrowExceptionForHR(hresult);
            engine = null;

            Console.WriteLine("ABBYY Engine is unloaded");
        }

        private string DetermineDocumentType(string projectId, string imageFile, string classifierPath)
        {
            Console.WriteLine("Determine Document Type is processing ...");
            if (engine == null)
            {
                engine = LoadEngine(projectId);
            }
            if (processor == null)
            {
                processor = engine.CreateFlexiCaptureProcessor();
                processor.AddClassificationTreeFile(classifierPath);
                Console.WriteLine("Classifier template is added ...");
            }
            else
            {
                processor.ResetProcessing();
                Console.WriteLine("Processor is reseted");
            }
            processor.AddImageFile(@imageFile);
            var result = processor.ClassifyNextPage();
            if (result != null && result.PageType == PageTypeEnum.PT_MeetsDocumentDefinition)
            {
                var names = result.GetClassNames();
                Console.WriteLine("Determine Document Type is completed");
                return names.Item(0);
            }


            Console.WriteLine("Determine Document Type is not completed");
            return "";
        }

        protected void MoveFileToDestinationFolder(string source, string destination, string pdfFile, string tiffFile, string filename)
        {
            Console.WriteLine("copy to destination folder " + destination);
            File.SetAttributes(pdfFile, FileAttributes.System);
            File.Copy(pdfFile, destination + "\\" + filename + ".pdf", true);
            Console.WriteLine("delete pdf file in the " + source);
            File.Delete(pdfFile);
        }

        protected int getTotalPage(string[] pdfList, string contain)
        {
            int total = 0;
            foreach (string pdf in pdfList)
            {
                if (pdf.Contains(contain))
                {
                    total++;
                }
            }
            return total;
        }
    }
}
