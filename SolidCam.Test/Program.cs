using Interop.SolidCAM2Lib;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidCam.Test {
    internal class Program {
        static void Main(string[] args) {
            Console.WriteLine("Hello, SolidCAM!");

            var calculateOperations = true;
            var generateGCode = true;
            Console.WriteLine($"Calculate Operations set to '{calculateOperations}'");
            Console.WriteLine($"Generate G Code set to '{generateGCode}'");

            Console.Write("Please provide a directory to process: ");
            var directory = Console.ReadLine();
            if (!Directory.Exists(directory)) {
                Console.WriteLine($"Directory '{directory}' does not exist!");
                return;
            }
            var partFiles = Directory.GetFiles(directory, "*.sldprt");
            if (partFiles.Length == 0) {
                Console.WriteLine($"Directory '{directory}' does not contain any SOLIDWORKS part files!");
                return;
            }

            var solidworksApplication = new SldWorks {
                Visible = true,
                CommandInProgress = true
            };
            try {
                foreach (var partFile in partFiles) {
                    Console.WriteLine($"Processing '{Path.GetFileName(partFile)}'");

                    int error = 0, warning = 0;
                    var model = solidworksApplication.OpenDoc6(partFile, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, string.Empty, ref error, ref warning);
                    ProcessPart(solidworksApplication, calculateOperations, generateGCode);
                    solidworksApplication.CloseDoc(model.GetPathName());
                }
            } catch (Exception) {
                throw;
            } finally {
                solidworksApplication.CommandInProgress = false;
            }
        }

        static void ProcessPart(SldWorks solidworksApplication, bool calculateOperations, bool generateGCode) {
            // Get SolidCAM add-in file path
            var solidCAMAddinPath = "C:\\Program Files\\SolidCAM2023\\Solidcam\\HostLibSW.dll";
            if (!File.Exists(solidCAMAddinPath)) {
                Console.WriteLine($" - SolidCAM add-in not found: '{solidCAMAddinPath}'");
                return;
            }

            // Ensure SolidCAM add-in is loaded
            var loadSolidCamAddinResult = solidworksApplication.LoadAddIn(solidCAMAddinPath);
            var loadSolidCamAddinSuccess = loadSolidCamAddinResult == (int)swLoadAddinError_e.swSuccess || loadSolidCamAddinResult == (int)swLoadAddinError_e.swAddinAlreadyLoaded;
            if (!loadSolidCamAddinSuccess) {
                Console.WriteLine($" - Failed to load the SolidCAM Add-in. Error {loadSolidCamAddinSuccess}");
                return;
            }

            // Connect to SolidCAM API
            if (new Launcher().LaunchFromCAD() is not IApplication solidCamApplication) {
                Console.WriteLine($" - Failed to connect to SolidCAM API");
                return;
            }

            if (solidCamApplication.Part is not IPart part) {
                Console.WriteLine($" - This is not a SolidCAM part");
                return;
            }

            try {
                // Sync, calculate, generate G Code
                var checkedSynchronisation = !(bool)part.CheckSynchronization();
                Console.WriteLine($" - Check Synchronize: {checkedSynchronisation}");

                var synchroised = !(bool)part.Synchronize();
                Console.WriteLine($" - Synchronize: {synchroised}");

                // Calculate operations (if selected)
                if (!calculateOperations) return;
                var operations = part.Operations as IOperations;

                // Suppress any non calculated operations
                for (var i = 0; i < (int)operations.Count; i++) {
                    if (operations.Item(i) is not IOperation operation) continue;

                    operation.Calculate(true);

                    var name = (string)operation.Name;
                    if ((bool)operation.Calculated) {
                        Console.WriteLine($"   - Operation '{name}' was calculated successfully");
                    } else {
                        Console.WriteLine($"   - Operation '{name}' failed to calculate and will be suppressed");
                        operation.Suppressed = true;
                    }
                }

                // Generate G Code (if selected)
                if (generateGCode) {
                    var generatedGCode = !(bool)operations.GenerateGCode();
                    Console.WriteLine($" - Generate G Code: {generatedGCode}");
                }
            } catch (Exception e) {
                Console.WriteLine(e.Message);
            } finally {
                part?.Close2(closeModel: false); // Only close SolidCAM (leaving the model open, so that another program can still be in control of closing the model in the usual manner)
            }
        }
    }
}