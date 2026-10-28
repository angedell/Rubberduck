using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Rubberduck.Parsing.UIContext;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Api
{
    /// <summary>A dialog the editor is showing, which nobody but a human could answer.</summary>
    public class BlockingDialog
    {
        public string Title { get; set; }

        /// <summary>The dialog's text, e.g. "Run-time error '11': Division by zero".</summary>
        public string Message { get; set; }

        public List<string> Buttons { get; set; } = new List<string>();

        /// <summary>The editor's own error dialog (run-time or compile error), as opposed to a message box raised by the user's code.</summary>
        [Newtonsoft.Json.JsonIgnore]
        public bool IsEditorError { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public IntPtr Handle { get; set; }
    }

    public interface IDialogWatcher
    {
        /// <summary>The dialog currently blocking the editor, or null.</summary>
        BlockingDialog Find();

        /// <summary>Presses the button that ends the error (End, or OK). False if no suitable button was found or the dialog did not go away.</summary>
        bool Acknowledge(BlockingDialog dialog);
    }

    /// <summary>
    /// <see cref="IDialogWatcher"/> on the dialogs of the editor's process. The editor's error dialogs are ordinary Windows dialogs: owned by the editor's window,
    /// with a Help button (control id 9), and their text and buttons can be read like any other. Only dialogs of this process are looked at.
    /// </summary>
    public class Win32DialogWatcher : IDialogWatcher
    {
        private const string DialogClass = "#32770";
        private const int HelpButtonId = 9;
        private const int BmClick = 0x00F5;
        private const uint GwOwner = 4;

        //The class of the editor's main window; its error dialogs are owned by it.
        private const string EditorWindowClass = "wndclass_desked_gsk";

        //What the button that ends the error is called; the first one present is pressed. Never Debug: that would leave the editor in break mode.
        private static readonly string[] Acknowledging = { "End", "OK" };

        private readonly IVBE _vbe;
        private readonly IUiDispatcher _dispatcher;

        public Win32DialogWatcher(IVBE vbe, IUiDispatcher dispatcher)
        {
            _vbe = vbe;
            _dispatcher = dispatcher;
        }

        public BlockingDialog Find()
        {
            //Does not ask the editor's UI thread for anything: while it sits in a dialog's message loop, that request may never be answered, and the dialog would never be found.
            //The add-in lives in the editor's process, and the editor's own windows are recognised by their class.
            var process = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;

            foreach (var window in TopLevelWindows())
            {
                GetWindowThreadProcessId(window, out var owner);
                if (owner != process || !IsWindowVisible(window) || ClassOf(window) != DialogClass)
                {
                    continue;
                }

                var dialog = new BlockingDialog { Handle = window, Title = TextOf(window) };
                var hasHelp = false;
                foreach (var child in Children(window))
                {
                    var kind = ClassOf(child);
                    if (kind == "Button")
                    {
                        dialog.Buttons.Add(TextOf(child).Replace("&", string.Empty));
                        hasHelp |= GetDlgCtrlID(child) == HelpButtonId;
                    }
                    else if (kind == "Static" && dialog.Message == null)
                    {
                        var text = TextOf(child);
                        if (text.Length > 0)
                        {
                            dialog.Message = text.Replace("\r\n", "\n").Trim();
                        }
                    }
                }

                dialog.IsEditorError = hasHelp && ClassOf(GetWindow(window, GwOwner)) == EditorWindowClass;
                return dialog;
            }

            return null;
        }

        public bool Acknowledge(BlockingDialog dialog)
        {
            foreach (var name in Acknowledging)
            {
                foreach (var child in Children(dialog.Handle))
                {
                    if (ClassOf(child) == "Button" && string.Equals(TextOf(child).Replace("&", string.Empty), name, StringComparison.OrdinalIgnoreCase))
                    {
                        //Sent from this thread, answered by the editor's own: it is inside the dialog's message loop and takes it.
                        SendMessage(child, BmClick, IntPtr.Zero, IntPtr.Zero);
                        return Gone(dialog.Handle);
                    }
                }
            }

            return false;
        }

        private static bool Gone(IntPtr handle)
        {
            for (var i = 0; i < 20; i++)
            {
                if (!IsWindowVisible(handle))
                {
                    return true;
                }

                Thread.Sleep(50);
            }

            return false;
        }

        private IntPtr MainWindow()
        {
            var task = _dispatcher.StartTask(() =>
            {
                using (var main = _vbe.MainWindow)
                {
                    return main.Handle();
                }
            }, CancellationToken.None);

            if (!task.Wait(TimeSpan.FromSeconds(30)))
            {
                throw ApiException.Unavailable("The editor did not respond in time.");
            }

            return task.Result;
        }

        #region Native

        private delegate bool EnumProc(IntPtr window, IntPtr parameter);

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

        private static List<IntPtr> TopLevelWindows()
        {
            var found = new List<IntPtr>();
            EnumWindows((window, _) => { found.Add(window); return true; }, IntPtr.Zero);
            return found;
        }

        private static List<IntPtr> Children(IntPtr parent)
        {
            var found = new List<IntPtr>();
            EnumChildWindows(parent, (window, _) => { found.Add(window); return true; }, IntPtr.Zero);
            return found;
        }

        private static string ClassOf(IntPtr window)
        {
            var name = new StringBuilder(256);
            GetClassName(window, name, name.Capacity);
            return name.ToString();
        }

        private static string TextOf(IntPtr window)
        {
            var text = new StringBuilder(2048);
            GetWindowText(window, text, text.Capacity);
            return text.ToString();
        }

        #endregion
    }
}
