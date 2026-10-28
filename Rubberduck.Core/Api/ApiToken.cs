using System;
using System.IO.Abstractions;
using System.Security.Cryptography;
using System.Text;

namespace Rubberduck.Api
{
    /// <summary>
    /// The secret an agent has to present. It is kept in the user's Rubberduck folder, so the same one works after a restart and the user can hand it to a sandbox once.
    /// </summary>
    public class ApiToken
    {
        private readonly IFileSystem _fileSystem;
        private readonly string _path;
        private readonly object _lock = new object();
        private string _value;

        public ApiToken(IFileSystem fileSystem, string path)
        {
            _fileSystem = fileSystem;
            _path = path;
        }

        public string Value
        {
            get
            {
                lock (_lock)
                {
                    return _value ?? (_value = Load() ?? Generate());
                }
            }
        }

        /// <summary>Makes a new token. Whoever held the old one is locked out.</summary>
        public string Regenerate()
        {
            lock (_lock)
            {
                return Generate();
            }
        }

        public bool Matches(string presented)
        {
            if (string.IsNullOrEmpty(presented))
            {
                return false;
            }

            return FixedTimeEquals(Encoding.UTF8.GetBytes(Value), Encoding.UTF8.GetBytes(presented));
        }

        private string Load()
        {
            try
            {
                if (_fileSystem.File.Exists(_path))
                {
                    var stored = _fileSystem.File.ReadAllText(_path).Trim();
                    if (stored.Length >= 32)
                    {
                        return stored;
                    }
                }
            }
            catch (System.IO.IOException)
            {
                //A token that cannot be read is replaced by a new one.
            }

            return null;
        }

        private string Generate()
        {
            var bytes = new byte[32];
            using (var random = RandomNumberGenerator.Create())
            {
                random.GetBytes(bytes);
            }

            _value = "rd_" + Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');
            _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(_path));
            _fileSystem.File.WriteAllText(_path, _value);
            return _value;
        }

        //net462 has no CryptographicOperations.FixedTimeEquals.
        private static bool FixedTimeEquals(byte[] left, byte[] right)
        {
            var difference = left.Length ^ right.Length;
            for (var i = 0; i < left.Length && i < right.Length; i++)
            {
                difference |= left[i] ^ right[i];
            }

            return difference == 0;
        }
    }
}
