using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using Rubberduck.PackageResources;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Ship
{
    /// <summary>
    /// Ship: the rules around the pipeline and the writer. A saved Excel workbook only; the folder is checked first; an existing file is dealt with as the settings or the caller say;
    /// the artifact is built in a temp folder and put in place at the end, so a failure never leaves half a file where a good one was. The user's project is only ever read.
    /// </summary>
    public class ShipService : IShipService
    {
        private readonly IFileSystem _fileSystem;
        private readonly IPackageResourceService _resources;
        private readonly ShipPipeline _pipeline;
        private readonly IShipArtifactWriter _writer;
        private readonly ShipFolders _folders;
        private readonly string _unattendedFile;
        private readonly string _tempRoot;
        private readonly ShipSource _source;

        /// <summary>The service without tree shaking: the step is skipped, with a warning, when it is switched on.</summary>
        public ShipService(IFileSystem fileSystem, IPackageResourceService resources, IStringParser parser)
            : this(fileSystem, resources, parser, (IShipAnalysis)null)
        {
        }

        /// <summary>
        /// The service as Rubberduck runs it: the real Excel, the real machine, the user's own settings folder, and the analysis over the parser state.
        /// Windsor takes the constructor with the most arguments it can satisfy, so a container without an <see cref="IShipAnalysis"/> gets the one above.
        /// </summary>
        public ShipService(IFileSystem fileSystem, IPackageResourceService resources, IStringParser parser, IShipAnalysis analysis)
            : this(fileSystem, resources, parser, new ShipArtifactWriter(fileSystem, new ShipModuleText(parser)), new ShipEnvironment(),
                ShipUnattended.FilePath, fileSystem.Path.Combine(fileSystem.Path.GetTempPath(), "Rubberduck-Ship"), analysis)
        {
        }

        public ShipService(IFileSystem fileSystem, IPackageResourceService resources, IStringParser parser, IShipArtifactWriter writer, IShipEnvironment environment,
            string unattendedFile, string tempRoot, IShipAnalysis analysis = null)
        {
            _fileSystem = fileSystem;
            _resources = resources;
            _pipeline = new ShipPipeline(parser, analysis);
            _writer = writer;
            _folders = new ShipFolders(fileSystem, environment);
            _unattendedFile = unattendedFile;
            _tempRoot = tempRoot;
            _source = new ShipSource(fileSystem);
        }

        #region Rules

        private static string HostOf(IVBProject project)
        {
            using (var host = project.VBE?.HostApplication())
            {
                return host?.ApplicationName;
            }
        }

        public bool CanShip(IVBProject project)
        {
            try
            {
                return string.Equals(HostOf(project), "Excel", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(project.FileName);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        private void RequireShippable(IVBProject project)
        {
            var host = HostOf(project);
            if (host == null)
            {
                throw new ShipRefusedException("unsupportedHost", "Ship works on Excel workbooks only; this host has no Ship.");
            }

            if (!string.Equals(host, "Excel", StringComparison.OrdinalIgnoreCase))
            {
                throw new ShipRefusedException("notExcel", $"Ship works on Excel workbooks only, not on {host}.");
            }

            if (string.IsNullOrWhiteSpace(project.FileName) || !_fileSystem.File.Exists(project.FileName))
            {
                throw new ShipRefusedException("unsaved", "The workbook has not been saved yet. Ship makes its copy from the saved file, so save it first.");
            }

            if (!project.IsSaved)
            {
                throw new ShipRefusedException("unsaved", "The workbook has unsaved changes. Ship makes its copy from the saved file, so save it first.");
            }
        }

        private string DocumentFolder(IVBProject project)
            => string.IsNullOrWhiteSpace(project.FileName) ? null : _fileSystem.Path.GetDirectoryName(project.FileName);

        public string ResolveOutputFolder(IVBProject project, ShipSettings settings)
            => _folders.Resolve(settings?.OutputFolder, DocumentFolder(project)).Path;

        //The folder as a ready, writable path; when nothing is set, null if that is allowed.
        private string RequireFolder(IVBProject project, ShipSettings settings, bool mustBeSet, bool probeWrite = true)
        {
            var resolved = _folders.Resolve(settings.OutputFolder, DocumentFolder(project));
            if (resolved.Path == null)
            {
                if (resolved.Problem == null && !mustBeSet)
                {
                    return null;
                }

                throw new ShipRefusedException("folderMissing", resolved.Problem ?? "No output folder is set. Choose where the shipped file goes.");
            }

            if (IsAutoLoadFolder(resolved.Path))
            {
                throw new ShipRefusedException("restrictedFolder", $"'{resolved.Path}' is a folder that Windows or Excel loads files from by itself. Choose another folder.");
            }

            _folders.Validate(resolved.Path, probeWrite);
            return resolved.Path;
        }

        //Folders Excel and Windows load files from by themselves: a macro-enabled file written there would run without anyone opening it.
        private static bool IsAutoLoadFolder(string folder)
        {
            var roots = new List<string>
            {
                Environment.GetFolderPath(Environment.SpecialFolder.Startup),
                Environment.GetFolderPath(Environment.SpecialFolder.CommonStartup),
                System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Excel", "XLSTART")
            };
            var programs = new[] { Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) };
            roots.AddRange(programs.Where(p => !string.IsNullOrEmpty(p)).Select(p => System.IO.Path.Combine(p, "Microsoft Office")));

            var full = System.IO.Path.GetFullPath(folder).TrimEnd('\\') + "\\";
            return roots.Where(r => !string.IsNullOrEmpty(r)).Any(r =>
            {
                var root = System.IO.Path.GetFullPath(r).TrimEnd('\\') + "\\";
                return full.StartsWith(root, StringComparison.OrdinalIgnoreCase)
                       && (!root.EndsWith("Microsoft Office\\", StringComparison.OrdinalIgnoreCase) || full.IndexOf("XLSTART", StringComparison.OrdinalIgnoreCase) >= 0);
            });
        }

        //The settings can be edited by an agent: the target must never be the workbook that is being shipped.
        private void RequireSafeTarget(IVBProject project, string target)
        {
            if (!string.IsNullOrWhiteSpace(project.FileName)
                && string.Equals(_fileSystem.Path.GetFullPath(project.FileName), _fileSystem.Path.GetFullPath(target), StringComparison.OrdinalIgnoreCase))
            {
                throw new ShipRefusedException("targetIsSource", "The shipped file would have the same path as the workbook itself. Choose another folder or file name.");
            }
        }

        private static string FileNameOf(IVBProject project, ShipSettings settings)
        {
            var name = string.IsNullOrWhiteSpace(settings.FileName) ? project.Name : settings.FileName;
            var invalid = System.IO.Path.GetInvalidFileNameChars();
            name = new string((name ?? string.Empty).Trim().Select(c => invalid.Contains(c) ? '_' : c).ToArray()).TrimEnd('.', ' ');
            return name.Length == 0 ? "Shipped" : name;
        }

        private static string Extension(ShipFormat format) => format == ShipFormat.Xlam ? ".xlam" : ".xlsm";

        #endregion

        #region Settings

        public ShipSettings ReadSettings(IVBProject project)
        {
            var text = ReadConfig(project, out _);
            if (text == null)
            {
                return new ShipSettings();
            }

            return new ProjectConfigReader(_fileSystem).Parse(text, null).Ship ?? new ShipSettings();
        }

        private string ReadConfig(IVBProject project, out string hash)
        {
            hash = null;
            try
            {
                if (string.IsNullOrWhiteSpace(project.FileName) || !_resources.IsSupported(project.FileName))
                {
                    return null;
                }

                var content = _resources.Read(project.FileName, PackageResourceNames.ConfigName);
                hash = content.Hash;
                return content.Text;
            }
            catch (PackageResourceException)
            {
                //No config, or the file cannot be read right now: the defaults are what there is.
                return null;
            }
        }

        public void WriteSettings(IVBProject project, ShipSettings settings)
        {
            if (string.IsNullOrWhiteSpace(project.FileName) || !_resources.IsSupported(project.FileName))
            {
                throw new ShipRefusedException("unsaved", "The settings are kept in the saved workbook file (xlsm, xlam); save the workbook as one of these first.");
            }

            var text = ReadConfig(project, out var hash);
            var updated = ProjectConfigWriter.SetShip(text ?? ProjectConfigTemplate.Create(new IndenterSettings(true)), settings);
            _resources.SetConfig(project.FileName, updated, hash);
        }

        #endregion

        #region Unattended

        private static string KeyOf(IVBProject project)
            => string.IsNullOrWhiteSpace(project.FileName) ? null : project.FileName;

        public bool IsUnattended(IVBProject project) => ShipUnattended.Load(_fileSystem, _unattendedFile).IsOn(KeyOf(project));

        public void SetUnattended(IVBProject project, bool value)
        {
            var store = ShipUnattended.Load(_fileSystem, _unattendedFile);
            store.Set(KeyOf(project), value);
            store.Save(_fileSystem, _unattendedFile);
        }

        #endregion

        #region Preview and ship

        public ShipReport Preview(IVBProject project, ShipSettings settings)
        {
            settings = settings ?? new ShipSettings();
            RequireShippable(project);
            var folder = RequireFolder(project, settings, mustBeSet: false, probeWrite: false);

            var work = NewWorkFolder();
            try
            {
                var set = _source.Build(project, _fileSystem.Path.Combine(work, "export"), PackageOf(project));
                var report = _pipeline.Run(set, settings, dryRun: true, project: project);
                if (folder != null)
                {
                    report.OutputPath = _fileSystem.Path.Combine(folder, FileNameOf(project, settings) + Extension(settings.Format));
                }

                return report;
            }
            finally
            {
                Delete(work);
            }
        }

        public ShipReport Ship(IVBProject project, ShipSettings settings, ShipOnExisting onExisting, IProgress<string> progress)
        {
            settings = settings ?? new ShipSettings();
            RequireShippable(project);
            var folder = RequireFolder(project, settings, mustBeSet: true);

            var target = _fileSystem.Path.Combine(folder, FileNameOf(project, settings) + Extension(settings.Format));
            RequireSafeTarget(project, target);
            var choice =onExisting != ShipOnExisting.Ask ? onExisting : settings.OnExisting;
            var replacing = false;
            if (_fileSystem.File.Exists(target))
            {
                switch (choice)
                {
                    case ShipOnExisting.Replace:
                        replacing = true;
                        break;
                    case ShipOnExisting.KeepBoth:
                        target = NextFree(target);
                        break;
                    default:
                        throw new ShipRefusedException("exists", $"The file '{target}' already exists. Choose whether to replace it or keep both.") { Path = target };
                }
            }

            var work = NewWorkFolder();
            try
            {
                progress?.Report("Reading the project");
                var set = _source.Build(project, _fileSystem.Path.Combine(work, "export"), PackageOf(project));
                var report = _pipeline.Run(set, settings, dryRun: false, project: project);

                //Built under the name it will have: an add-in takes its name from its file.
                var artifact = _fileSystem.Path.Combine(work, "out", _fileSystem.Path.GetFileName(target));
                _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(artifact));
                _writer.Write(new ShipArtifactJob
                {
                    SourceFile = project.FileName,
                    Set = set,
                    Format = settings.Format,
                    ArtifactPath = artifact,
                    WorkFolder = work
                }, progress);

                progress?.Report("Writing " + _fileSystem.Path.GetFileName(target));
                Deliver(artifact, target, replacing);
                report.OutputPath = target;
                progress?.Report("Done");
                return report;
            }
            finally
            {
                Delete(work);
            }
        }

        private Func<System.IO.Stream> PackageOf(IVBProject project)
        {
            if (!_resources.IsSupported(project.FileName))
            {
                return null;
            }

            //Excel has the workbook open and shares reads.
            return () => _fileSystem.File.Open(project.FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite | System.IO.FileShare.Delete);
        }

        private string NextFree(string path)
        {
            var folder = _fileSystem.Path.GetDirectoryName(path);
            var name = _fileSystem.Path.GetFileNameWithoutExtension(path);
            var extension = _fileSystem.Path.GetExtension(path);
            for (var number = 2; ; number++)
            {
                var candidate = _fileSystem.Path.Combine(folder, $"{name} ({number}){extension}");
                if (!_fileSystem.File.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        //The file goes into the folder under a temporary name first, so the swap happens between two files on the same volume. The old file is the backup until the swap is done.
        private void Deliver(string artifact, string target, bool replacing)
        {
            var temporary = target + ".rd-ship.tmp";
            var backup = FreeBackupName(target);
            try
            {
                _fileSystem.File.Copy(artifact, temporary, true);
                if (replacing && _fileSystem.File.Exists(target))
                {
                    //The old file stays as the backup: a backup the user made by hand is never overwritten, the next free name is used.
                    _fileSystem.File.Replace(temporary, target, backup);
                }
                else
                {
                    _fileSystem.File.Move(temporary, target);
                }
            }
            catch (Exception exception) when (exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                TryDelete(temporary);
                throw new System.IO.IOException($"The file '{target}' could not be written; it may be open in Excel or another program. {exception.Message}", exception);
            }
        }

        private string FreeBackupName(string target)
        {
            var backup = target + ".bak";
            for (var number = 2; _fileSystem.File.Exists(backup); number++)
            {
                backup = target + ".bak" + number;
            }

            return backup;
        }

        //A work folder holds a full copy of the user's workbook: one left behind by a failed delete is removed by a later ship.
        private void SweepOldWorkFolders()
        {
            try
            {
                if (!_fileSystem.Directory.Exists(_tempRoot))
                {
                    return;
                }

                foreach (var old in _fileSystem.Directory.GetDirectories(_tempRoot).Where(d => _fileSystem.Directory.GetCreationTimeUtc(d) < DateTime.UtcNow.AddHours(-12)))
                {
                    Delete(old);
                }
            }
            catch (Exception exception) when (exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                //Best effort.
            }
        }

        private string NewWorkFolder()
        {
            SweepOldWorkFolders();
            var folder = _fileSystem.Path.Combine(_tempRoot, Guid.NewGuid().ToString("N"));
            _fileSystem.Directory.CreateDirectory(folder);
            return folder;
        }

        private void Delete(string folder)
        {
            try
            {
                if (_fileSystem.Directory.Exists(folder))
                {
                    _fileSystem.Directory.Delete(folder, true);
                }
            }
            catch (Exception exception) when (exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                //Left in the temp folder; the next ship does not depend on it.
            }
        }

        private void TryDelete(string path)
        {
            try
            {
                if (_fileSystem.File.Exists(path))
                {
                    _fileSystem.File.Delete(path);
                }
            }
            catch (Exception exception) when (exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                //Nothing more to do.
            }
        }

        #endregion
    }
}
