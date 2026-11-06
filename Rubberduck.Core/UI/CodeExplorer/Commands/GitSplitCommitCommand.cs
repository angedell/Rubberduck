using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Git;
using Rubberduck.Interaction;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.UI.CodeExplorer.GitClient;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>Saves the project's changes as several commits, one per functionality, by sorting the changed blocks of code into groups.</summary>
    public class GitSplitCommitCommand : GitClientCommandBase
    {
        public GitSplitCommitCommand(IVbeEvents vbeEvents, IProjectsProvider projectsProvider, GitClientService git, IMessageBox messageBox)
            : base(vbeEvents, projectsProvider, git, messageBox)
        {
        }

        protected override void OnExecute(object parameter)
        {
            var project = TargetProject(parameter);
            if (project == null)
            {
                return;
            }

            Guarded(() =>
            {
                var changes = ReadChanges(project, out _);
                if (changes.Count == 0)
                {
                    MessageBox.Message(string.Format(CodeExplorerUI.GitClient_NothingToCommit, project.Name));
                    return;
                }

                var hunks = Git.SplitHunks(project);
                Func<System.Windows.Forms.IWin32Window, GitGroupSuggestion> group = null;
                if (Git.AiAvailable)
                {
                    group = owner => GroupWithAi(hunks);
                }

                var parts = GitSplitDialog.Ask(project.Name, hunks, group);
                if (parts == null)
                {
                    return;
                }

                var ids = Git.CommitSplit(project, hunks, parts);
                MessageBox.Message(string.Format(CodeExplorerUI.GitClient_CommittedInParts, ids.Count, changes.Count));
            });
        }

        //Shows what would leave the computer and sends it only if the user agrees.
        private GitGroupSuggestion GroupWithAi(IReadOnlyList<GitHunk> hunks)
        {
            GitGroupSuggestion suggestion = null;
            Guarded(() =>
            {
                try
                {
                    var preview = Git.PrepareGrouping(hunks);
                    if (ConfirmSending(preview))
                    {
                        suggestion = Git.SuggestGroups(hunks, preview);
                    }
                }
                catch (AiException exception)
                {
                    Logger.Warn(exception, "Git client: the AI service did not help.");
                    MessageBox.NotifyWarn(exception.Message, CodeExplorerUI.GitClient_DialogTitle);
                }
            });
            return suggestion;
        }
    }
}
