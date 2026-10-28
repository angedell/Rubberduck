using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.Api
{
    /// <summary>The commands that give Rubberduck the right to listen on the network, written out so the user can read, copy and run them.</summary>
    public static class ApiSetupCommands
    {
        public static string CurrentUser => $"{Environment.UserDomainName}\\{Environment.UserName}";

        public static string UrlAcl(int port, string user) => $"netsh http add urlacl url=http://+:{port}/ user=\"{user}\"";

        public static string Firewall(int port) => $"netsh advfirewall firewall add rule name=\"Rubberduck API\" dir=in action=allow protocol=TCP localport={port} profile=private,domain";
    }

    public enum ApiSetupChoice
    {
        TryAgain,
        LocalOnly,
        Cancel
    }

    /// <summary>
    /// Shown when the API is set to accept agents from the network and Windows does not allow it yet.
    /// Explains why, shows the exact commands with a button to copy each, and can run them elevated, so the user only has to say yes to the Windows prompt.
    /// </summary>
    public sealed class ApiSetupDialog : Form
    {
        private readonly string _urlAcl;
        private readonly string _firewall;
        private readonly CheckBox _withFirewall;
        private readonly Button _run;
        private readonly Label _status;

        private ApiSetupDialog(int port)
        {
            _urlAcl = ApiSetupCommands.UrlAcl(port, ApiSetupCommands.CurrentUser);
            _firewall = ApiSetupCommands.Firewall(port);

            Text = "Rubberduck API - one-time setup";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(640, 424);

            var y = 14;
            Add(new Label { Text = "Windows needs your permission before agents can connect over the network", Font = new Font(Font.FontFamily, 11f, FontStyle.Bold) }, 16, ref y, 608, 24);
            Add(new Label
            {
                Text = "You switched on \"Also accept agents from the network\". Windows only lets a program listen on the network when an administrator has allowed it, "
                       + "and Rubberduck cannot allow itself. This is needed once for this port; you will not see this window again."
            }, 16, ref y, 608, 48);
            y += 4;

            AddStepTitle("Step 1.  Allow Rubberduck to listen on port " + port, ref y);
            AddCommand(_urlAcl, ref y);
            Add(new Label { Text = "Run it in a Command Prompt opened with \"Run as administrator\", or let Rubberduck run it for you:" }, 16, ref y, 608, 20);
            _run = new Button { Text = "Run as administrator", AutoSize = true, Padding = new Padding(8, 2, 8, 2) };
            _run.Click += async (sender, args) => await RunElevated();
            Add(_run, 16, ref y, 190, 30);
            y += 6;

            AddStepTitle("Step 2.  Only if your sandbox still cannot connect: open the firewall", ref y);
            AddCommand(_firewall, ref y);
            _withFirewall = new CheckBox { Text = "Also run this when I click \"Run as administrator\" (private and domain networks only)", AutoSize = true };
            Add(_withFirewall, 16, ref y, 608, 22);
            y += 6;

            AddStepTitle("Step 3.  Click \"Try again\"", ref y);
            _status = new Label { Text = string.Empty, ForeColor = Color.DimGray };
            Add(_status, 16, ref y, 608, 36);

            var tryAgain = new Button { Text = "Try again", DialogResult = DialogResult.Retry, Size = new Size(100, 28), Location = new Point(ClientSize.Width - 16 - 100, ClientSize.Height - 44) };
            var local = new Button { Text = "Use this PC only", DialogResult = DialogResult.Ignore, Size = new Size(120, 28), Location = new Point(tryAgain.Left - 8 - 120, tryAgain.Top) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(80, 28), Location = new Point(local.Left - 8 - 80, tryAgain.Top) };
            Controls.AddRange(new Control[] { tryAgain, local, cancel });
            AcceptButton = tryAgain;
            CancelButton = cancel;
        }

        /// <summary>Asks the user what to do about the missing permission.</summary>
        public static ApiSetupChoice Ask(int port)
        {
            using (var dialog = new ApiSetupDialog(port))
            {
                switch (dialog.ShowDialog())
                {
                    case DialogResult.Retry:
                        return ApiSetupChoice.TryAgain;
                    case DialogResult.Ignore:
                        return ApiSetupChoice.LocalOnly;
                    default:
                        return ApiSetupChoice.Cancel;
                }
            }
        }

        private void Add(Control control, int x, ref int y, int width, int height)
        {
            control.Location = new Point(x, y);
            control.Size = new Size(width, height);
            if (control is Label label)
            {
                label.AutoSize = false;
            }

            Controls.Add(control);
            y += height + 4;
        }

        private void AddStepTitle(string text, ref int y)
            => Add(new Label { Text = text, Font = new Font(Font, FontStyle.Bold) }, 16, ref y, 608, 20);

        //The command in a box that can be selected, next to a button that puts it on the clipboard.
        private void AddCommand(string command, ref int y)
        {
            var box = new TextBox { Text = command, ReadOnly = true, Font = new Font("Consolas", 9f), BackColor = SystemColors.Window };
            Add(box, 16, ref y, 520, 24);
            y -= 28;
            var copy = new Button { Text = "Copy" };
            copy.Click += (sender, args) =>
            {
                try
                {
                    Clipboard.SetText(command);
                    _status.ForeColor = Color.DimGray;
                    _status.Text = "Copied to the clipboard.";
                }
                catch (System.Runtime.InteropServices.ExternalException)
                {
                    //Another program holds the clipboard; selecting the text works as well.
                    _status.ForeColor = Color.Firebrick;
                    _status.Text = "The clipboard is in use by another program. Select the command and press Ctrl+C instead.";
                }
            };
            Add(copy, 544, ref y, 80, 24);
        }

        private async Task RunElevated()
        {
            _run.Enabled = false;
            _status.ForeColor = Color.DimGray;
            _status.Text = "Waiting for the Windows permission prompt...";
            try
            {
                var commands = _withFirewall.Checked ? _urlAcl + " && " + _firewall : _urlAcl;
                var code = await Task.Run(() =>
                {
                    var start = new ProcessStartInfo("cmd.exe", "/c " + commands) { Verb = "runas", UseShellExecute = true, WindowStyle = ProcessWindowStyle.Hidden };
                    using (var process = Process.Start(start))
                    {
                        process.WaitForExit();
                        return process.ExitCode;
                    }
                });

                if (code == 0)
                {
                    _status.ForeColor = Color.SeaGreen;
                    _status.Text = "Done. Click \"Try again\".";
                }
                else
                {
                    _status.ForeColor = Color.Firebrick;
                    _status.Text = $"Windows reported an error (code {code}). If the permission was already there, that is fine: click \"Try again\" to check.";
                }
            }
            catch (Win32Exception exception) when (exception.NativeErrorCode == 1223)
            {
                _status.ForeColor = Color.Firebrick;
                _status.Text = "The permission prompt was cancelled, so nothing was changed.";
            }
            catch (Exception exception) when (exception is Win32Exception || exception is InvalidOperationException)
            {
                _status.ForeColor = Color.Firebrick;
                _status.Text = "Could not start the command: " + exception.Message;
            }
            finally
            {
                _run.Enabled = true;
            }
        }
    }
}
