using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Abstractions;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using NLog;
using Rubberduck.PackageResources;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Ship
{
    /// <summary>What the writer builds: the saved file, what the steps decided, and where the artifact goes.</summary>
    public class ShipArtifactJob
    {
        /// <summary>The user's saved workbook. Only read.</summary>
        public string SourceFile { get; set; }

        public ShipWorkingSet Set { get; set; }
        public ShipFormat Format { get; set; }

        /// <summary>Where the finished artifact is to be made; it does not exist yet and its file name is the one the artifact will keep.</summary>
        public string ArtifactPath { get; set; }

        /// <summary>A folder of its own for the copy and the temporary module files. The caller deletes it.</summary>
        public string WorkFolder { get; set; }
    }

    public interface IShipArtifactWriter
    {
        /// <summary>Builds the artifact. Throws <see cref="ShipRefusedException"/> when Excel cannot be used for it.</summary>
        void Write(ShipArtifactJob job, IProgress<string> progress);
    }

    /// <summary>
    /// Builds the artifact in a second, hidden Excel (late bound): the saved file is copied, the copy is opened with macros force-disabled so that none of the user's code runs,
    /// the steps' results are applied to the copy's project, the copy is saved as the artifact, and Excel is closed. The user's own Excel and project are never involved.
    /// </summary>
    public class ShipArtifactWriter : IShipArtifactWriter
    {
        private const int AutomationSecurityForceDisable = 3;
        private const int XlOpenXmlWorkbookMacroEnabled = 52;
        private const int XlOpenXmlAddIn = 55;
        private const int ComponentDocument = 100;
        private const int ComponentStandard = 1;
        private const int ComponentClass = 2;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IFileSystem _fileSystem;
        private readonly ShipModuleText _text;

        public ShipArtifactWriter(IFileSystem fileSystem, ShipModuleText text)
        {
            _fileSystem = fileSystem;
            _text = text;
        }

        /// <summary>How long the second Excel may take before it is ended.</summary>
        public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(5);

        private void OpenApplyAndSave(dynamic excel, string copy, ShipArtifactJob job, IProgress<string> progress)
        {
            dynamic workbook;
            try
            {
                workbook = excel.Workbooks.Open(copy, 0, false);
            }
            catch (COMException exception)
            {
                throw new ShipRefusedException("excelUnavailable", "Excel could not open the copy of the workbook: " + exception.Message);
            }

            try
            {
                Apply(workbook, job, progress);
                progress?.Report("Saving " + _fileSystem.Path.GetFileName(job.ArtifactPath));
                workbook.SaveAs(job.ArtifactPath, job.Format == ShipFormat.Xlam ? XlOpenXmlAddIn : XlOpenXmlWorkbookMacroEnabled);
            }
            finally
            {
                try
                {
                    workbook.Close(false);
                }
                catch (COMException exception)
                {
                    Logger.Debug(exception, "Could not close the copy.");
                }

                Marshal.FinalReleaseComObject((object)workbook);
            }
        }

        private static void TryEnd(int process, ISet<int> alreadyRunning)
        {
            if (process == 0 || alreadyRunning.Contains(process))
            {
                return;
            }

            try
            {
                Process.GetProcessById(process).Kill();
            }
            catch (Exception exception) when (exception is ArgumentException || exception is InvalidOperationException || exception is System.ComponentModel.Win32Exception)
            {
                Logger.Debug(exception, "The second Excel was already gone.");
            }
        }

        public void Write(ShipArtifactJob job, IProgress<string> progress)
        {
            var type = Type.GetTypeFromProgID("Excel.Application");
            if (type == null)
            {
                throw new ShipRefusedException("excelUnavailable", "Excel is not installed (or not registered), so Ship cannot open a copy of the workbook.");
            }

            var copy = _fileSystem.Path.Combine(job.WorkFolder, "copy", _fileSystem.Path.GetFileName(job.SourceFile));
            _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(copy));
            progress?.Report("Copying the saved workbook");
            Copy(job.SourceFile, copy);

            dynamic excel = null;
            var process = 0;
            var alreadyRunning = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(p => p.Id));
            try
            {
                progress?.Report("Starting a second Excel");
                try
                {
                    excel = Activator.CreateInstance(type);
                }
                catch (COMException exception)
                {
                    throw new ShipRefusedException("excelUnavailable", "Excel could not be started: " + exception.Message);
                }

                process = ProcessOf(excel);
                if (process == 0 || alreadyRunning.Contains(process))
                {
                    //Either Excel handed out the instance the user is working in, or there is no telling whose it is. Neither is ours to hide, change or close.
                    Marshal.FinalReleaseComObject((object)excel);
                    excel = null;
                    process = 0;
                    throw new ShipRefusedException("excelUnavailable", "Windows did not start a separate Excel for Ship that Ship can tell apart from the open one. Close other Excel windows, or try again.");
                }

                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.EnableEvents = false;
                excel.ScreenUpdating = false;
                excel.AskToUpdateLinks = false;
                //Before anything is opened: no macro of the user's runs in this Excel.
                excel.AutomationSecurity = AutomationSecurityForceDisable;

                //All the work with the copy runs under a watchdog: a dialog that nobody can see would otherwise keep the hidden Excel, and Ship, waiting for ever.
                //On a timeout only this Excel (a process that was not running before) is ended, the user's own is never touched.
                var hidden = excel;
                var work = Task.Run((Action)(() => OpenApplyAndSave(hidden, copy, job, progress)));
                if (!work.Wait(Timeout))
                {
                    TryEnd(process, alreadyRunning);
                    throw new ShipRefusedException("excelUnavailable",
                        $"The second Excel did not answer within {Timeout.TotalMinutes:0} minutes (it may be waiting for a dialog) and was ended. Nothing was written.");
                }
            }
            catch (AggregateException exception) when (exception.InnerExceptions.Count == 1)
            {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            }
            finally
            {
                Quit(excel, process);
            }

            RemoveParts(job, progress);
        }

        private void Copy(string source, string target)
        {
            //Excel has the workbook open and shares reads.
            using (var from = _fileSystem.File.Open(source, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite | System.IO.FileShare.Delete))
            using (var to = _fileSystem.File.Create(target))
            {
                from.CopyTo(to);
            }
        }

        private void Apply(dynamic workbook, ShipArtifactJob job, IProgress<string> progress)
        {
            dynamic project;
            try
            {
                project = workbook.VBProject;
                int protection = project.Protection;
                if (protection != 0)
                {
                    throw new ShipRefusedException("excelUnavailable", "The VBA project is locked with a password, so Ship cannot change its copy. Unlock it, save, and try again.");
                }
            }
            catch (COMException exception)
            {
                throw new ShipRefusedException("vbomNotTrusted",
                    "Excel does not allow access to the VBA project: switch on \"Trust access to the VBA project object model\" in Excel's Trust Center (Macro Settings). " + exception.Message);
            }

            progress?.Report("Applying the steps to the copy");
            dynamic components = project.VBComponents;

            foreach (var module in job.Set.Modules.Where(m => m.Removed))
            {
                dynamic component = components.Item(module.Name);
                if ((int)component.Type == ComponentDocument)
                {
                    dynamic code = component.CodeModule;
                    int lines = code.CountOfLines;
                    if (lines > 0)
                    {
                        code.DeleteLines(1, lines);
                    }
                }
                else
                {
                    components.Remove(component);
                }
            }

            foreach (var module in job.Set.Modules.Where(m => m.Changed))
            {
                Replace(components, module, job);
            }

            if (job.Set.References.Any(r => r.Removed))
            {
                dynamic references = project.References;
                foreach (var removed in job.Set.References.Where(r => r.Removed))
                {
                    foreach (dynamic reference in ((System.Collections.IEnumerable)references).Cast<object>().ToList())
                    {
                        string guid = reference.Guid;
                        bool builtIn = reference.BuiltIn;
                        if (!builtIn && string.Equals((guid ?? string.Empty).Trim('{', '}'), (removed.Guid ?? string.Empty).Trim('{', '}'), StringComparison.OrdinalIgnoreCase))
                        {
                            references.Remove(reference);
                            break;
                        }
                    }
                }
            }
        }

        //A standard or class module comes back through an import, which is the only way to change its attribute lines; a document module or a form takes new code in place (the VBE
        //does not let anyone import over a sheet, and a form's designer must stay as it is).
        private void Replace(dynamic components, ShipModule module, ShipArtifactJob job)
        {
            dynamic component = components.Item(module.Name);
            var type = (int)component.Type;
            if (type == ComponentStandard || type == ComponentClass)
            {
                var file = _fileSystem.Path.Combine(job.WorkFolder, "modules", module.Name + ShipSource.ExtensionOf(module.Type));
                _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(file));
                _fileSystem.File.WriteAllText(file, module.Text, ShipSource.VbeEncoding);
                components.Remove(component);
                components.Import(file);
                return;
            }

            var code = _text.CodeOnly(module.Name, job.Set.ProjectId, module.Text);
            if (!code.Parsed)
            {
                job.Set.Warnings.Add($"Module {module.Name} was left as it was: {code.Problem}");
                return;
            }

            ReplaceCode(component.CodeModule, code.Text);
        }

        //The code of a module is replaced as a whole, never appended to: a document module always exists and may already say "Option Explicit", and a second one is a compile error that
        //the VBE reports in a dialog. Whatever is left after the deletion is read back and duplicated Option lines are taken out.
        private static void ReplaceCode(dynamic codeModule, string code)
        {
            for (var attempt = 0; attempt < 3; attempt++)
            {
                int existing = codeModule.CountOfLines;
                if (existing == 0)
                {
                    break;
                }

                codeModule.DeleteLines(1, existing);
            }

            codeModule.AddFromString(code);

            int count = codeModule.CountOfLines;
            var written = count == 0 ? string.Empty : (string)codeModule.Lines(1, count);
            var clean = ShipCodeText.WithoutRepeatedOptions(written);
            if (!string.Equals(clean, written, StringComparison.Ordinal))
            {
                codeModule.DeleteLines(1, count);
                codeModule.AddFromString(clean);
            }
        }

        //The parts the steps took out of the package: Excel wrote the artifact, so this is done to the finished file.
        private void RemoveParts(ShipArtifactJob job, IProgress<string> progress)
        {
            if (job.Set.RemovedResources.Count == 0)
            {
                return;
            }

            progress?.Report("Removing " + string.Join(", ", job.Set.RemovedResources));
            var edits = job.Set.RemovedResources.ToDictionary(name => name, name => (byte[])null, StringComparer.OrdinalIgnoreCase);
            var patched = job.ArtifactPath + ".patch";
            using (var source = _fileSystem.File.OpenRead(job.ArtifactPath))
            using (var target = _fileSystem.File.Create(patched))
            {
                OfficePackage.Patch(source, target, edits);
            }

            _fileSystem.File.Delete(job.ArtifactPath);
            _fileSystem.File.Move(patched, job.ArtifactPath);
        }

        private static int ProcessOf(dynamic excel)
        {
            try
            {
                GetWindowThreadProcessId(new IntPtr((int)excel.Hwnd), out var process);
                return (int)process;
            }
            catch (Exception exception) when (exception is COMException || exception is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return 0;
            }
        }

        private static void Quit(dynamic excel, int process)
        {
            if (excel == null)
            {
                return;
            }

            try
            {
                excel.Quit();
            }
            catch (COMException exception)
            {
                Logger.Debug(exception, "Excel did not quit cleanly.");
            }

            Marshal.FinalReleaseComObject((object)excel);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //Excel can linger for a moment; Ship must not leave one behind.
            if (process != 0)
            {
                try
                {
                    var running = Process.GetProcessById(process);
                    if (!running.WaitForExit(8000))
                    {
                        running.Kill();
                    }
                }
                catch (ArgumentException)
                {
                }
            }
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    }
}
