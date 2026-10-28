using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>
    /// Picks a virtual folder (the argument of '@Folder) from a tree of the folders the project has, the way a file dialog does: select one, or make a subfolder in it.
    /// A new folder is only a name until something is saved in it; that is all a virtual folder is.
    /// </summary>
    public class FolderPickerDialog : Form
    {
        private readonly bool _allowNone;
        private readonly TreeView _tree;
        private readonly Label _path;
        private readonly Button _ok;
        private readonly TreeNode _none;

        /// <param name="allowNone">Whether "no folder" can be chosen. Where a folder is required, the top of the tree only stands for the top level, to make folders in.</param>
        public FolderPickerDialog(IEnumerable<string> folders, string current, bool allowNone = true)
        {
            _allowNone = allowNone;
            Text = CodeExplorerUI.FolderPicker_Title;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(380, 420);
            MinimumSize = new Size(300, 300);
            ShowIcon = false;
            ShowInTaskbar = false;
            KeyPreview = true;

            var images = new ImageList { ColorDepth = ColorDepth.Depth32Bit, ImageSize = new Size(16, 16) };
            images.Images.Add("closed", CodeExplorerUI.FolderClosed);
            images.Images.Add("open", CodeExplorerUI.FolderOpen);

            _tree = new TreeView
            {
                Dock = DockStyle.Fill,
                ImageList = images,
                ImageKey = "closed",
                SelectedImageKey = "open",
                HideSelection = false,
                LabelEdit = true,
                ShowLines = true,
                ShowRootLines = true,
                BorderStyle = BorderStyle.FixedSingle,
                Sorted = true
            };
            _tree.AfterSelect += (sender, args) => ShowPath();
            _tree.BeforeLabelEdit += (sender, args) => args.CancelEdit = args.Node == _none;
            _tree.AfterLabelEdit += AfterLabelEdit;
            _tree.NodeMouseDoubleClick += (sender, args) =>
            {
                if (args.Button == MouseButtons.Left && args.Node.Tag != null)
                {
                    _ok.PerformClick();
                }
            };
            _tree.KeyDown += (sender, args) =>
            {
                if (args.KeyCode == Keys.Insert || args.Control && args.KeyCode == Keys.N)
                {
                    args.Handled = true;
                    NewSubfolder();
                }
                else if (args.KeyCode == Keys.F2 && _tree.SelectedNode != null && _tree.SelectedNode != _none)
                {
                    args.Handled = true;
                    _tree.SelectedNode.BeginEdit();
                }
            };

            //The root stands for "no @Folder": the module sits in the project's own folder.
            _none = new TreeNode(allowNone ? CodeExplorerUI.FolderPicker_None : CodeExplorerUI.FolderPicker_TopLevel) { Tag = string.Empty, ImageKey = "open", SelectedImageKey = "open", NodeFont = new Font(Font, FontStyle.Italic) };
            _tree.Nodes.Add(_none);
            foreach (var folder in (folders ?? Enumerable.Empty<string>()).Concat(new[] { current }).Where(f => !string.IsNullOrWhiteSpace(f)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                Ensure(folder);
            }

            _tree.ExpandAll();
            _tree.SelectedNode = Find(current) ?? _none;
            _tree.SelectedNode.EnsureVisible();

            _path = new Label { Dock = DockStyle.Top, Height = 24, Padding = new Padding(2, 5, 0, 0), AutoEllipsis = true };

            var newFolder = new Button { Text = CodeExplorerUI.FolderPicker_New, Size = new Size(110, 26), Location = new Point(0, 7) };
            newFolder.Click += (sender, args) => NewSubfolder();
            _ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Size = new Size(84, 26) };
            var cancel = new Button { Text = CodeExplorerUI.ResourceEditor_Cancel, DialogResult = DialogResult.Cancel, Size = new Size(84, 26) };
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, Padding = new Padding(0, 6, 0, 0) };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_ok);
            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(10, 0, 10, 0) };
            bottom.Controls.Add(buttons);
            bottom.Controls.Add(newFolder);

            var body = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 8, 10, 0) };
            body.Controls.Add(_tree);
            body.Controls.Add(_path);
            Controls.Add(body);
            Controls.Add(bottom);

            AcceptButton = _ok;
            CancelButton = cancel;
            ShowPath();
        }

        /// <summary>The folder the user chose: "" for none. Only meaningful when the dialog ended with OK.</summary>
        public string SelectedFolder => PathOf(_tree.SelectedNode);

        /// <summary>Shows the dialog. Null if the user cancelled; otherwise the folder, "" for none.</summary>
        public static string Pick(IWin32Window owner, IEnumerable<string> folders, string current, bool allowNone = true)
        {
            using (var dialog = new FolderPickerDialog(folders, current, allowNone))
            {
                return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.SelectedFolder : null;
            }
        }

        /// <summary>For the Move to Folder dialogs, which are WPF: lets the user pick the folder in a tree instead of typing it, and puts it in the view model.</summary>
        public static void Browse(System.Windows.FrameworkElement view)
        {
            if (!(view?.DataContext is IFolderBrowsing browsing))
            {
                return;
            }

            var window = System.Windows.Window.GetWindow(view);
            var owner = window == null ? null : new WindowHandle(new System.Windows.Interop.WindowInteropHelper(window).Handle);
            var chosen = Pick(owner, browsing.KnownFolders, browsing.NewFolder, false);
            if (!string.IsNullOrEmpty(chosen))
            {
                browsing.NewFolder = chosen;
            }
        }

        private class WindowHandle : IWin32Window
        {
            public WindowHandle(IntPtr handle) => Handle = handle;

            public IntPtr Handle { get; }
        }

        private static string PathOf(TreeNode node) => node?.Tag as string ?? string.Empty;

        private TreeNode Find(string folder)
        {
            var node = _none;
            foreach (var segment in (folder ?? string.Empty).Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries))
            {
                node = node.Nodes.Cast<TreeNode>().FirstOrDefault(n => string.Equals(n.Text, segment.Trim(), StringComparison.OrdinalIgnoreCase));
                if (node == null)
                {
                    return null;
                }
            }

            return node;
        }

        private void Ensure(string folder)
        {
            var node = _none;
            var path = string.Empty;
            foreach (var raw in folder.Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var segment = raw.Trim();
                path = path.Length == 0 ? segment : path + "." + segment;
                var child = node.Nodes.Cast<TreeNode>().FirstOrDefault(n => string.Equals(n.Text, segment, StringComparison.OrdinalIgnoreCase));
                if (child == null)
                {
                    child = node.Nodes.Add(segment);
                    child.Tag = path;
                }

                node = child;
            }
        }

        private void ShowPath()
        {
            //The tree selects a node while the dialog is still being built.
            if (_path == null)
            {
                return;
            }

            var folder = SelectedFolder;
            _path.Text = string.Format(CodeExplorerUI.FolderPicker_Selected, folder.Length == 0 ? _none.Text : folder.Replace(".", " › "));
            _ok.Enabled = _allowNone || folder.Length > 0;
        }

        //A folder that is still being named has no path yet: its Tag is null until the name is accepted.
        private void NewSubfolder()
        {
            var parent = _tree.SelectedNode ?? _none;
            var node = parent.Nodes.Add(CodeExplorerUI.FolderPicker_NewNode);
            node.Tag = null;
            parent.Expand();
            _tree.SelectedNode = node;
            node.BeginEdit();
        }

        private void AfterLabelEdit(object sender, NodeLabelEditEventArgs args)
        {
            var node = args.Node;
            var isNew = node.Tag == null;

            //Escape on a folder that was just made removes it again; on a rename it keeps the old name.
            if (args.Label == null)
            {
                if (isNew)
                {
                    BeginInvoke((Action)(() => node.Remove()));
                }

                return;
            }

            var text = args.Label.Trim();
            args.CancelEdit = true;
            if (text.Length == 0 || text.IndexOf('.') >= 0 || !NoteModule.IsValidFolder(text))
            {
                MessageBox.Show(this, CodeExplorerUI.FolderPicker_Invalid, Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (isNew)
                {
                    BeginInvoke((Action)(() => node.Remove()));
                }

                return;
            }

            var parent = node.Parent ?? _none;
            var twin = parent.Nodes.Cast<TreeNode>().FirstOrDefault(n => n != node && string.Equals(n.Text, text, StringComparison.OrdinalIgnoreCase));
            if (twin != null)
            {
                //The folder is there already: take the user to it.
                BeginInvoke((Action)(() =>
                {
                    if (isNew)
                    {
                        node.Remove();
                    }

                    _tree.SelectedNode = twin;
                }));
                return;
            }

            //The edit is cancelled so that the new name is set by hand, along with the paths below it.
            node.Text = text;
            Retag(node, PathOf(parent));
            BeginInvoke((Action)(() =>
            {
                _tree.Sort();
                _tree.SelectedNode = node;
                ShowPath();
            }));
        }

        private static void Retag(TreeNode node, string parentPath)
        {
            node.Tag = parentPath.Length == 0 ? node.Text : parentPath + "." + node.Text;
            foreach (TreeNode child in node.Nodes)
            {
                Retag(child, (string)node.Tag);
            }
        }
    }
}

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>A view model with a folder to type that <see cref="FolderPickerDialog.Browse"/> can fill in for the user.</summary>
    public interface IFolderBrowsing
    {
        string NewFolder { get; set; }

        /// <summary>The folders the project has.</summary>
        IReadOnlyList<string> KnownFolders { get; }
    }
}
