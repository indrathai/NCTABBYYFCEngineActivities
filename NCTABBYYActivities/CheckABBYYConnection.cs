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
    [DisplayName("Check ABBYY Connection")]
    public class CheckAbbyyConnection : CodeActivity
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

        [Category("Output")]
        [DisplayName("Log Messages")]
        public OutArgument<List<LogMessage>> LogMessages { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            Console.WriteLine("Trying to connect to ABBYY Engine ...");
            List<LogMessage> logList = new List<LogMessage>();
            logList.Add(new LogMessage("Trying to connect to ABBYY Engine ...", LogType.Information));
            try
            {


                //get Document Project Id
                string documentProjectId = DocumentProjectId.Get(context);
                Console.WriteLine("Document Project ID: " + documentProjectId);

                engine = LoadEngine(documentProjectId);

            }
            #pragma warning disable CS0618 // Type or member is obsolete
            catch (ExecutionEngineException e)
            #pragma warning restore CS0618 // Type or member is obsolete
            {
                Console.WriteLine("error: " + e.Message);
                logList.Add(new LogMessage("Abbyy Engine is failed to connect because of " + e.Message, LogType.Error));
            }
            finally
            {
                UnloadEngine(ref engine);
                logList.Add(new LogMessage("ABBYY Engine is unloaded...", LogType.Information));
                LogMessages.Set(context, logList);
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
    }
}
