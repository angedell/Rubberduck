using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>How long it is since the user last touched the keyboard or the mouse.</summary>
    public static class UserIdle
    {
        public static TimeSpan Time()
        {
            var info = new LastInputInfo { Size = (uint)Marshal.SizeOf(typeof(LastInputInfo)) };
            if (!GetLastInputInfo(ref info))
            {
                //Cannot tell: say the user is busy, so that nothing optional goes on behind their back.
                return TimeSpan.Zero;
            }

            //Both are milliseconds since the start and wrap together; the unsigned difference is right across the wrap.
            return TimeSpan.FromMilliseconds(unchecked((uint)Environment.TickCount - info.Time));
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LastInputInfo
        {
            public uint Size;
            public uint Time;
        }

        [DllImport("user32.dll")]
        private static extern bool GetLastInputInfo(ref LastInputInfo info);
    }
}
