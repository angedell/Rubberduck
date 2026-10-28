using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Rubberduck.Parsing.UIContext;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.Api
{
    /// <summary><see cref="IProjectFiles"/> on top of the VBE. The VBE can only be talked to from its UI thread, so the lookup is handed over to it.</summary>
    public class VbeProjectFiles : IProjectFiles
    {
        private static readonly TimeSpan UiTimeout = TimeSpan.FromSeconds(60);

        private readonly IProjectsProvider _projects;
        private readonly IUiDispatcher _dispatcher;

        public VbeProjectFiles(IProjectsProvider projects, IUiDispatcher dispatcher)
        {
            _projects = projects;
            _dispatcher = dispatcher;
        }

        public string DocumentOf(string project)
        {
            var task = _dispatcher.StartTask(() => FileNameOf(project), CancellationToken.None);
            try
            {
                if (!task.Wait(UiTimeout))
                {
                    throw ApiException.Unavailable("The editor did not respond in time. It may be showing a dialog or running code.");
                }

                return task.Result;
            }
            catch (AggregateException exception)
            {
                var inner = exception.InnerException ?? exception;
                if (inner is ApiException api)
                {
                    throw api;
                }

                if (inner is COMException)
                {
                    throw ApiException.Unavailable($"The editor refused the request ({inner.Message}). It may be in break mode or running code.");
                }

                throw;
            }
        }

        private string FileNameOf(string key)
        {
            var all = _projects.Projects().ToList();
            var byId = all.Where(p => string.Equals(p.ProjectId, key, StringComparison.OrdinalIgnoreCase)).ToList();
            var matches = byId.Count > 0 ? byId : all.Where(p => string.Equals(p.Project.Name, key, StringComparison.OrdinalIgnoreCase)).ToList();
            switch (matches.Count)
            {
                case 0:
                    throw ApiException.NotFound($"No open project '{key}'.");
                case 1:
                    return matches[0].Project.FileName;
                default:
                    throw ApiException.Conflict($"More than one open project is named '{key}'; use its id (see /projects).");
            }
        }
    }
}
