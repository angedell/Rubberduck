using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.PackageResources;

namespace Rubberduck.Api
{
    /// <summary>The file a project is saved in, which is where its ribbon XML and pictures are.</summary>
    public interface IProjectFiles
    {
        /// <summary>The path of the project's document, or null if it has not been saved.</summary>
        /// <exception cref="ApiException">No such open project, or more than one.</exception>
        string DocumentOf(string project);
    }

    /// <summary>
    /// What agents can do with a document's ribbon XML and pictures. Changes are staged and reach the file when the document is closed,
    /// because the application keeps it locked while it is open; every answer says so.
    /// </summary>
    public class ResourceApi
    {
        private const string ApplyNote = "Staged: the document is open, so the change goes into its file when the document is closed. Reads already include it.";

        private readonly IProjectFiles _files;
        private readonly IPackageResourceService _resources;
        private readonly ApiWriteApproval _approval;

        public ResourceApi(IProjectFiles files, IPackageResourceService resources, ApiWriteApproval approval)
        {
            _files = files;
            _resources = resources;
            _approval = approval;
        }

        public ResourceList List(string project)
        {
            var document = _files.DocumentOf(project);
            if (!_resources.IsSupported(document))
            {
                return new ResourceList
                {
                    Project = project,
                    Supported = false,
                    Note = "This project's document is not saved as an Open XML file (xlsm, docm, pptm...), so it has no ribbon XML or pictures to show."
                };
            }

            return Translate(() =>
            {
                var pending = _resources.HasPending(document);
                return new ResourceList
                {
                    Project = project,
                    Supported = true,
                    PendingChanges = pending,
                    Note = pending ? ApplyNote : null,
                    Resources = _resources.List(document).Select(Info).ToList()
                };
            });
        }

        public ResourceContent Read(string project, string name)
        {
            var document = Document(project);
            return Translate(() =>
            {
                var content = _resources.Read(document, name);
                var result = new ResourceContent
                {
                    Name = content.Name,
                    Kind = KindName(content.Kind),
                    Size = content.Size,
                    Hash = content.Hash,
                    Pending = Pending(content.Change)
                };

                if (content.Kind == PackageResourceKind.RibbonXml)
                {
                    result.Xml = content.Text;
                }
                else if (content.Kind == PackageResourceKind.Config)
                {
                    result.Text = content.Text;
                }
                else
                {
                    result.ContentBase64 = Convert.ToBase64String(content.Bytes);
                }

                return result;
            });
        }

        /// <summary>Stages ribbon XML (<paramref name="xml"/>), the config (<paramref name="text"/>) or a picture (<paramref name="contentBase64"/>), adding or replacing.</summary>
        public ResourceChangeResult Write(string client, string project, string name, string xml, string text, string contentBase64, string expectedHash)
        {
            if (new[] { xml, text, contentBase64 }.Count(value => value != null) != 1)
            {
                throw ApiException.BadRequest("Send exactly one of 'xml' (customUI.xml / customUI14.xml), 'text' (rubberduck.json) or 'contentBase64' (a picture).");
            }

            //Check what can be checked before the user is asked anything.
            var kind = PackageResourceNames.KindOf(name);
            var expected = xml != null ? PackageResourceKind.RibbonXml : text != null ? PackageResourceKind.Config : PackageResourceKind.Image;
            if (kind != expected)
            {
                throw ApiException.BadRequest(
                    kind == PackageResourceKind.RibbonXml ? $"{name} is ribbon XML: send it as 'xml'."
                    : kind == PackageResourceKind.Config ? $"{name} is the config: send it as 'text'."
                    : $"'{name}' is a picture: send it as 'contentBase64'. XML goes to {PackageResourceNames.Ribbon2007} or {PackageResourceNames.Ribbon2010}, JSON to {PackageResourceNames.ConfigName}.");
            }

            byte[] bytes = null;
            if (contentBase64 != null)
            {
                try
                {
                    bytes = Convert.FromBase64String(contentBase64);
                }
                catch (FormatException)
                {
                    throw ApiException.BadRequest("'contentBase64' is not valid base64.");
                }
            }

            var document = Document(project);
            _approval.Require(client);
            return Translate(() =>
            {
                PackageResource resource;
                List<string> warnings = null;
                switch (expected)
                {
                    case PackageResourceKind.RibbonXml:
                        resource = _resources.SetXml(document, name, xml, expectedHash);
                        break;
                    case PackageResourceKind.Config:
                        resource = _resources.SetConfig(document, text, expectedHash);
                        warnings = _resources.CheckConfig(text).ToList();
                        break;
                    default:
                        resource = _resources.SetImage(document, name, bytes, expectedHash);
                        break;
                }

                return new ResourceChangeResult { Resource = Info(resource), Note = ApplyNote, Warnings = warnings };
            });
        }

        public ResourceChangeResult Delete(string client, string project, string name, string expectedHash)
        {
            if (string.IsNullOrWhiteSpace(expectedHash))
            {
                throw ApiException.PreconditionRequired("Send the hash from your last read of the picture (or the If-Match header), so a picture changed in the meantime is not removed.");
            }

            var document = Document(project);
            _approval.Require(client);
            return Translate(() =>
            {
                var current = _resources.Read(document, name);
                if (!string.Equals(current.Hash, expectedHash, StringComparison.OrdinalIgnoreCase))
                {
                    throw new PackageResourceException(PackageResourceError.Stale, $"'{name}' changed since you read it. Read it again.");
                }

                _resources.RemoveImages(document, new[] { current.Name });
                return new ResourceChangeResult
                {
                    Resource = new ResourceInfo { Name = current.Name, Kind = KindName(current.Kind), Size = current.Size, Pending = Pending(PackageResourceChange.Removed) },
                    Note = ApplyNote
                };
            });
        }

        /// <summary>Throws away every staged change of the project's document, the way the Code Explorer's discard does. What the file holds stays as it is.</summary>
        public ResourceList DiscardPending(string client, string project)
        {
            var document = Document(project);
            if (!_resources.HasPending(document))
            {
                return List(project);
            }

            _approval.Require(client);
            Translate(() =>
            {
                _resources.Discard(document);
                return true;
            });
            return List(project);
        }

        private string Document(string project)
        {
            var document = _files.DocumentOf(project);
            if (!_resources.IsSupported(document))
            {
                throw ApiException.Conflict("This project's document is not saved as an Open XML file (xlsm, docm, pptm...), so it has no ribbon XML or pictures.");
            }

            return document;
        }

        private static T Translate<T>(Func<T> work)
        {
            try
            {
                return work();
            }
            catch (PackageResourceException exception)
            {
                switch (exception.Error)
                {
                    case PackageResourceError.NotFound: throw ApiException.NotFound(exception.Message);
                    case PackageResourceError.Conflict: throw ApiException.Conflict(exception.Message);
                    case PackageResourceError.Stale: throw ApiException.PreconditionFailed(exception.Message);
                    case PackageResourceError.HashRequired: throw ApiException.PreconditionRequired(exception.Message);
                    case PackageResourceError.Unavailable: throw ApiException.Unavailable(exception.Message);
                    default: throw ApiException.BadRequest(exception.Message);
                }
            }
        }

        private static ResourceInfo Info(PackageResource resource)
            => new ResourceInfo { Name = resource.Name, Kind = KindName(resource.Kind), Size = resource.Size, Pending = Pending(resource.Change) };

        private static string KindName(PackageResourceKind kind) => kind == PackageResourceKind.RibbonXml ? "ribbon" : kind == PackageResourceKind.Config ? "config" : "image";

        private static string Pending(PackageResourceChange change)
        {
            switch (change)
            {
                case PackageResourceChange.Added: return "added";
                case PackageResourceChange.Modified: return "modified";
                case PackageResourceChange.Removed: return "removed";
                default: return null;
            }
        }
    }
}
