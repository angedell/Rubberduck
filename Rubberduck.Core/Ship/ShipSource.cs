using System;
using System.IO.Abstractions;
using System.Text;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Ship
{
    /// <summary>
    /// Reads the user's project into a working set without touching it: each component is exported to a folder (the VBE only reads while it does that)
    /// and the text of the exported file is what the steps work on.
    /// </summary>
    internal class ShipSource
    {
        private readonly IFileSystem _fileSystem;

        public ShipSource(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
        }

        //The VBE writes and reads its files in the ANSI code page of the machine.
        public static Encoding VbeEncoding => Encoding.Default;

        public ShipWorkingSet Build(IVBProject project, string exportFolder, Func<System.IO.Stream> openPackage)
        {
            var set = new ShipWorkingSet(project.ProjectId) { OpenPackage = openPackage };
            _fileSystem.Directory.CreateDirectory(exportFolder);

            using (var components = project.VBComponents)
            {
                foreach (var component in components)
                {
                    using (component)
                    {
                        var extension = ExtensionOf(component.Type);
                        if (extension == null)
                        {
                            continue;
                        }

                        var file = _fileSystem.Path.Combine(exportFolder, component.Name + extension);
                        component.Export(file);
                        set.Modules.Add(new ShipModule(component.Name, component.Type, _fileSystem.File.ReadAllText(file, VbeEncoding)) { ExportFile = file });
                    }
                }
            }

            using (var references = project.References)
            {
                foreach (var reference in references)
                {
                    using (reference)
                    {
                        set.References.Add(new ShipReference { Name = reference.Name, Guid = reference.Guid, Major = reference.Major, Minor = reference.Minor });
                    }
                }
            }

            return set;
        }

        /// <summary>The extension the VBE gives an exported component, or null for the kinds Ship leaves alone.</summary>
        public static string ExtensionOf(ComponentType type)
        {
            switch (type)
            {
                case ComponentType.StandardModule: return ".bas";
                case ComponentType.ClassModule:
                case ComponentType.Document: return ".cls";
                case ComponentType.UserForm: return ".frm";
                default: return null;
            }
        }
    }
}
