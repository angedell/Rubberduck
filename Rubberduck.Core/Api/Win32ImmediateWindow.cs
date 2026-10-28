using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using Rubberduck.Parsing.UIContext;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Api
{
    /// <summary>
    /// <see cref="IImmediateWindow"/> through the editor's own window.
    /// The window draws its text itself, so it has no handle to read from: the text is read through UI Automation, which the editor supports for the sake of screen readers.
    /// Typing is done by posting the characters, and then Enter as a key press, to that one window. Nothing goes through the keyboard, so nothing needs focus
    /// and no other application can get the keystrokes.
    /// </summary>
    public class Win32ImmediateWindow : IImmediateWindow
    {
        private const string WindowClass = "VbaWindow";
        private const int WmKeyDown = 0x0100;
        private const int WmKeyUp = 0x0101;
        private const int WmChar = 0x0102;
        private const int VkReturn = 0x0D;

        private readonly IVBE _vbe;
        private readonly IUiDispatcher _dispatcher;

        public Win32ImmediateWindow(IVBE vbe, IUiDispatcher dispatcher)
        {
            _vbe = vbe;
            _dispatcher = dispatcher;
        }

        public string Read()
        {
            var handle = Locate();

            //Called from a thread that is not the editor's, so the editor can answer while this waits.
            var window = AutomationElement.FromHandle(handle);
            var document = window.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Document));
            if (document == null || !document.TryGetCurrentPattern(TextPattern.Pattern, out var pattern))
            {
                throw ApiException.Unavailable("The Immediate window does not expose its text on this version of the editor.");
            }

            return ((TextPattern)pattern).DocumentRange.GetText(-1);
        }

        public void Send(string line)
        {
            var handle = Locate();
            foreach (var character in line)
            {
                PostMessage(handle, WmChar, (IntPtr)character, (IntPtr)1);
            }

            //Enter is what runs the line, and the editor looks for it as a key, not as a character.
            PostMessage(handle, WmKeyDown, (IntPtr)VkReturn, (IntPtr)0x001C0001);
            PostMessage(handle, WmKeyUp, (IntPtr)VkReturn, unchecked((IntPtr)(int)0xC01C0001));
        }

        //The window can be closed or hidden by the user at any time, so it is looked up for every call. The editor's objects are only touched on its own thread.
        private IntPtr Locate()
        {
            var task = _dispatcher.StartTask(() =>
            {
                string caption = null;
                using (var windows = _vbe.Windows)
                {
                    foreach (var window in windows)
                    {
                        using (window)
                        {
                            if (caption == null && window.Type == WindowKind.Immediate)
                            {
                                caption = window.Caption;
                                if (!window.IsVisible)
                                {
                                    window.IsVisible = true;
                                }
                            }
                        }
                    }
                }

                using (var main = _vbe.MainWindow)
                {
                    return (Caption: caption, Main: main.Handle());
                }
            }, CancellationToken.None);

            if (!task.Wait(TimeSpan.FromSeconds(30)))
            {
                throw ApiException.Unavailable("The editor did not respond in time.");
            }

            var (title, mainHandle) = task.Result;
            if (title == null)
            {
                throw ApiException.Unavailable("The editor has no Immediate window.");
            }

            foreach (var child in Descendants(mainHandle))
            {
                if (ClassOf(child) == WindowClass && TextOf(child) == title)
                {
                    return child;
                }
            }

            throw ApiException.Unavailable("The Immediate window could not be found. Is the editor open?");
        }

        #region Native

        private delegate bool EnumProc(IntPtr window, IntPtr parameter);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr parameter);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern bool PostMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

        private static List<IntPtr> Descendants(IntPtr parent)
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
            var text = new StringBuilder(256);
            GetWindowText(window, text, text.Capacity);
            return text.ToString();
        }

        #endregion
    }
}
