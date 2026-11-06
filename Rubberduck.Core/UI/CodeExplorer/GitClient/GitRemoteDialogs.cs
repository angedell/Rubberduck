using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Rubberduck.Git;

namespace Rubberduck.UI.CodeExplorer.GitClient
{
    /// <summary>Asks for the address of the remote repository.</summary>
    public sealed class GitRemoteDialog : Form
    {
        private readonly TextBox _url;

        private GitRemoteDialog(string projectName, string current)
        {
            Text = "Git: remote of '" + projectName + "'";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(520, 130);

            Controls.Add(new Label { Text = "Address of the repository (https), for example https://github.com/owner/repository.git", Location = new Point(12, 10), Size = new Size(496, 36) });
            _url = new TextBox { Text = current ?? string.Empty, Location = new Point(12, 50), Size = new Size(496, 24) };
            Controls.Add(_url);

            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Size = new Size(90, 28), Location = new Point(ClientSize.Width - 12 - 90 - 8 - 90, ClientSize.Height - 40) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(90, 28), Location = new Point(ClientSize.Width - 12 - 90, ClientSize.Height - 40) };
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        /// <returns>The address, or null if the user cancelled.</returns>
        public static string Ask(string projectName, string current)
        {
            using (var dialog = new GitRemoteDialog(projectName, current))
            {
                return dialog.ShowDialog() == DialogResult.OK ? dialog._url.Text.Trim() : null;
            }
        }
    }

    /// <summary>Asks for the personal access token a host wants. It goes to the Credential Manager; nothing else keeps it.</summary>
    public sealed class GitTokenDialog : Form
    {
        private readonly TextBox _user;
        private readonly TextBox _token;

        private GitTokenDialog(string host, string reason)
        {
            Text = "Git: sign in to " + host;
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(520, 250);

            Controls.Add(new Label { Text = reason, Location = new Point(12, 10), Size = new Size(496, 36) });

            if (host == "github.com")
            {
                var link = new LinkLabel { Text = "Create a token on GitHub (it needs read and write access to the repository's contents)", Location = new Point(12, 50), Size = new Size(496, 36) };
                link.LinkClicked += (sender, args) => Process.Start("https://github.com/settings/personal-access-tokens/new");
                Controls.Add(link);
            }

            Controls.Add(new Label { Text = "User name (optional):", Location = new Point(12, 96), Size = new Size(496, 18) });
            _user = new TextBox { Location = new Point(12, 116), Size = new Size(496, 24) };
            Controls.Add(_user);

            Controls.Add(new Label { Text = "Token:", Location = new Point(12, 146), Size = new Size(496, 18) });
            _token = new TextBox { Location = new Point(12, 166), Size = new Size(496, 24), UseSystemPasswordChar = true };
            Controls.Add(_token);

            var ok = new Button { Text = "Save", DialogResult = DialogResult.OK, Size = new Size(90, 28), Location = new Point(ClientSize.Width - 12 - 90 - 8 - 90, ClientSize.Height - 40) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(90, 28), Location = new Point(ClientSize.Width - 12 - 90, ClientSize.Height - 40) };
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
            Shown += (sender, args) => _token.Focus();
        }

        /// <returns>The user name (may be empty) and token, or null if the user cancelled or left the token empty.</returns>
        public static StoredCredential Ask(string host, string reason)
        {
            using (var dialog = new GitTokenDialog(host, reason))
            {
                return dialog.ShowDialog() == DialogResult.OK && dialog._token.Text.Trim().Length > 0
                    ? new StoredCredential(dialog._user.Text.Trim(), dialog._token.Text.Trim())
                    : null;
            }
        }
    }
}
