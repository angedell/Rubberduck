using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Rubberduck.Git
{
    public class StoredCredential
    {
        public StoredCredential(string user, string secret)
        {
            User = user;
            Secret = secret;
        }

        public string User { get; }
        public string Secret { get; }
    }

    /// <summary>Where a user's secrets are kept: a git token per host, an API key per provider. Never the settings file, never the repository.</summary>
    public interface ICredentialStore
    {
        /// <returns>The credential saved under the target, or null.</returns>
        StoredCredential Read(string target);

        void Write(string target, string user, string secret);

        /// <returns>Whether there was anything to delete.</returns>
        bool Delete(string target);

        /// <summary>The targets saved by Rubberduck whose name starts with the prefix, so settings can list and delete them.</summary>
        IReadOnlyList<string> List(string prefix);
    }

    /// <summary>Names the targets Rubberduck saves under, so that they cannot be confused with anyone else's entries.</summary>
    public static class CredentialTargets
    {
        public const string Prefix = "Rubberduck/";

        public static string GitHost(string host) => Prefix + "git/" + host.ToLowerInvariant();

        public static string Ai(string provider) => Prefix + "ai/" + provider.ToLowerInvariant();
    }

    /// <summary>
    /// The Windows Credential Manager. A secret is kept encrypted for the Windows user who saved it, in the same store the system's own
    /// sign-ins use, so the user can also see and delete it in Control Panel.
    /// </summary>
    public class WindowsCredentialStore : ICredentialStore
    {
        private const int TypeGeneric = 1;
        private const int PersistLocalMachine = 2;
        private const int ErrorNotFound = 1168;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct NativeCredential
        {
            public int Flags;
            public int Type;
            public string TargetName;
            public string Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public int CredentialBlobSize;
            public IntPtr CredentialBlob;
            public int Persist;
            public int AttributeCount;
            public IntPtr Attributes;
            public string TargetAlias;
            public string UserName;
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CredReadW")]
        private static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credential);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CredWriteW")]
        private static extern bool CredWrite(ref NativeCredential credential, int flags);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CredDeleteW")]
        private static extern bool CredDelete(string target, int type, int flags);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CredEnumerateW")]
        private static extern bool CredEnumerate(string filter, int flags, out int count, out IntPtr credentials);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern void CredFree(IntPtr buffer);

        public StoredCredential Read(string target)
        {
            if (!CredRead(target, TypeGeneric, 0, out var pointer))
            {
                var error = Marshal.GetLastWin32Error();
                if (error == ErrorNotFound)
                {
                    return null;
                }

                throw new System.ComponentModel.Win32Exception(error);
            }

            try
            {
                return ToCredential((NativeCredential)Marshal.PtrToStructure(pointer, typeof(NativeCredential)));
            }
            finally
            {
                CredFree(pointer);
            }
        }

        public void Write(string target, string user, string secret)
        {
            var blob = Encoding.Unicode.GetBytes(secret);
            var handle = GCHandle.Alloc(blob, GCHandleType.Pinned);
            try
            {
                var credential = new NativeCredential
                {
                    Type = TypeGeneric,
                    TargetName = target,
                    UserName = user,
                    CredentialBlobSize = blob.Length,
                    CredentialBlob = handle.AddrOfPinnedObject(),
                    Persist = PersistLocalMachine
                };

                if (!CredWrite(ref credential, 0))
                {
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                }
            }
            finally
            {
                Array.Clear(blob, 0, blob.Length);
                handle.Free();
            }
        }

        public bool Delete(string target)
        {
            if (CredDelete(target, TypeGeneric, 0))
            {
                return true;
            }

            var error = Marshal.GetLastWin32Error();
            if (error == ErrorNotFound)
            {
                return false;
            }

            throw new System.ComponentModel.Win32Exception(error);
        }

        public IReadOnlyList<string> List(string prefix)
        {
            var targets = new List<string>();
            if (!CredEnumerate(prefix + "*", 0, out var count, out var array))
            {
                var error = Marshal.GetLastWin32Error();
                if (error == ErrorNotFound)
                {
                    return targets;
                }

                throw new System.ComponentModel.Win32Exception(error);
            }

            try
            {
                for (var i = 0; i < count; i++)
                {
                    var pointer = Marshal.ReadIntPtr(array, i * IntPtr.Size);
                    var credential = (NativeCredential)Marshal.PtrToStructure(pointer, typeof(NativeCredential));
                    targets.Add(credential.TargetName);
                }
            }
            finally
            {
                CredFree(array);
            }

            return targets;
        }

        private static StoredCredential ToCredential(NativeCredential native)
        {
            var secret = native.CredentialBlobSize > 0
                ? Marshal.PtrToStringUni(native.CredentialBlob, native.CredentialBlobSize / 2)
                : string.Empty;
            return new StoredCredential(native.UserName, secret);
        }
    }
}
