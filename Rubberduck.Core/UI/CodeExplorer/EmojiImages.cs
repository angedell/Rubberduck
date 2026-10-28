using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using NLog;
using SharpVectors.Converters;
using SharpVectors.Renderers.Wpf;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>
    /// The pictures of the emoji (Noto Color Emoji, Apache 2.0; see Resources/Emoji). WPF cannot draw colour fonts, so the renderer puts these pictures in the text instead.
    /// <list type="bullet">
    /// <item>The assembly carries the index of every emoji there is, and the 128 px PNGs of the core: the single characters and the keycaps.</item>
    /// <item>The others (skin tones, families and other sequences) are fetched in the background into a folder next to the settings: while the user is idle, a little at a time,
    /// and only if Windows does not count the connection as metered. An emoji a note is waiting for goes first. Until its picture is there the emoji is left to the font.</item>
    /// <item>The SVGs, for a zoom beyond what the PNG can give, are fetched the first time they are needed, and kept the same way.</item>
    /// </list>
    /// </summary>
    public class EmojiImages
    {
        private const string PackageResource = "Rubberduck.Emoji.noto-emoji-128.zip";
        private const string IndexResource = "Rubberduck.Emoji.noto-emoji-index.txt";
        private const string FilePrefix = "emoji_u";
        private const string SvgAddress = "https://raw.githubusercontent.com/googlefonts/noto-emoji/main/2D/svg/emoji_u{0}.svg";
        private const string PngAddress = "https://raw.githubusercontent.com/googlefonts/noto-emoji/main/2D/png/128/emoji_u{0}.png";
        private const int FailuresBeforeGivingUp = 5;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly TimeSpan UserIdleFor = TimeSpan.FromSeconds(5);
        private static readonly Lazy<EmojiImages> SharedInstance = new Lazy<EmojiImages>(() => new EmojiImages(
            () => typeof(EmojiImages).Assembly.GetManifestResourceStream(PackageResource),
            () => typeof(EmojiImages).Assembly.GetManifestResourceStream(IndexResource),
            System.IO.Path.Combine(Rubberduck.Resources.ApplicationConstants.RUBBERDUCK_FOLDER_PATH, "emoji"),
            NetworkCost.CanDownloadOptionalData));

        private readonly IFileSystem _fileSystem;
        private readonly Func<System.IO.Stream> _openPackage;
        private readonly Func<System.IO.Stream> _openIndex;
        private readonly string _pngFolder;
        private readonly string _svgFolder;
        private readonly Func<bool> _canDownload;
        private readonly Func<string, Task<byte[]>> _download;
        private readonly Func<bool> _isIdle;
        private readonly TimeSpan _pause;
        private readonly object _lock = new object();
        private readonly Dictionary<string, ImageSource> _png = new Dictionary<string, ImageSource>();
        private readonly Dictionary<string, Task<byte[]>> _svgDownloads = new Dictionary<string, Task<byte[]>>();
        private readonly LinkedList<string> _wanted = new LinkedList<string>();
        private readonly HashSet<string> _asked = new HashSet<string>();
        private ZipArchive _package;
        private HashSet<string> _packaged;
        private HashSet<string> _keys;
        private SynchronizationContext _context;
        private Thread _worker;
        private volatile bool _stop;

        /// <param name="openPackage">The zip of the core PNGs, named emoji_u&lt;key&gt;.png.</param>
        /// <param name="openIndex">The text with the key of every emoji there is, one per line.</param>
        /// <param name="folder">Where fetched pictures are kept: png and svg folders are made in it.</param>
        /// <param name="canDownload">Whether the connection allows a download now (see <see cref="NetworkCost"/>).</param>
        /// <param name="download">Gets the bytes of an address; the default is an HTTP GET.</param>
        /// <param name="isIdle">Whether the user is away enough for the background fetching to go on; the default looks at the last keystroke or mouse move.</param>
        /// <param name="pause">The wait between two pictures fetched in the background.</param>
        public EmojiImages(
            Func<System.IO.Stream> openPackage,
            Func<System.IO.Stream> openIndex,
            string folder,
            Func<bool> canDownload,
            Func<string, Task<byte[]>> download = null,
            IFileSystem fileSystem = null,
            Func<bool> isIdle = null,
            TimeSpan? pause = null)
        {
            _fileSystem = fileSystem ?? new FileSystem();
            _openPackage = openPackage;
            _openIndex = openIndex;
            _pngFolder = System.IO.Path.Combine(folder, "png");
            _svgFolder = System.IO.Path.Combine(folder, "svg");
            _canDownload = canDownload;
            _download = download ?? HttpDownload;
            _isIdle = isIdle ?? (() => UserIdle.Time() >= UserIdleFor);
            _pause = pause ?? TimeSpan.FromMilliseconds(300);
        }

        public static EmojiImages Shared => SharedInstance.Value;

        /// <summary>Raised, on the thread that started the fetching, when the picture of an emoji a document was waiting for (see <see cref="Png"/>) has arrived.</summary>
        public event Action<string> PictureArrived;

        /// <summary>Whether there is such an emoji, with its picture here or still to come (see <see cref="EmojiPiece.Key"/>).</summary>
        public bool IsKnown(string key)
        {
            EnsurePackage();
            return _keys.Contains(key);
        }

        /// <summary>
        /// The 128 px picture, frozen so that every document can share it. Null if it is not here yet: the emoji is moved to the front of the fetching (if that is running,
        /// see <see cref="StartFetching"/>) and <see cref="PictureArrived"/> says when it comes.
        /// </summary>
        public ImageSource Png(string key)
        {
            lock (_lock)
            {
                EnsurePackage();
                if (_png.TryGetValue(key, out var cached))
                {
                    return cached;
                }

                var bytes = ReadPackaged(key) ?? ReadFile(PngPath(key));
                if (bytes == null)
                {
                    if (_keys.Contains(key) && _asked.Add(key))
                    {
                        _wanted.AddLast(key);
                    }

                    return null;
                }

                var image = Decode(bytes, key);
                if (image != null)
                {
                    _png[key] = image;
                }

                return image;
            }
        }

        /// <summary>
        /// Starts fetching the pictures that are not here, if it is not going already: one at a time, only while the user is idle (but at once for an emoji a document waits for) and
        /// the connection is not metered. Gives up for the session after a few failures in a row. Call it on the UI thread: <see cref="PictureArrived"/> comes back there.
        /// </summary>
        public void StartFetching()
        {
            lock (_lock)
            {
                if (_worker != null)
                {
                    return;
                }

                _context = SynchronizationContext.Current;
                _stop = false;
                _worker =new Thread(Fetch) { IsBackground = true, Priority = ThreadPriority.BelowNormal, Name = "Rubberduck emoji pictures" };
                _worker.Start();
            }
        }

        /// <summary>Stops the fetching; what was fetched stays.</summary>
        public void StopFetching() => _stop = true;

        /// <summary>The keys that have no picture here yet, in the order they will be fetched.</summary>
        public IReadOnlyList<string> Missing()
        {
            lock (_lock)
            {
                EnsurePackage();
                return _keys.Where(key => !_packaged.Contains(key) && !_fileSystem.File.Exists(PngPath(key))).OrderBy(key => key, StringComparer.Ordinal).ToList();
            }
        }

        private void Fetch()
        {
            var failures = 0;
            var rest = new Queue<string>(Missing());
            Logger.Debug("Fetching the emoji pictures that are not here yet: {0}.", rest.Count);

            while (!_stop)
            {
                string key = null;
                var wanted = false;
                lock (_lock)
                {
                    if (_wanted.Count > 0)
                    {
                        key = _wanted.First.Value;
                        _wanted.RemoveFirst();
                        wanted = true;
                    }
                }

                if (key == null)
                {
                    if (rest.Count == 0)
                    {
                        //Everything is here; a document may still ask for something new later, so stay for a while.
                        Thread.Sleep(TimeSpan.FromSeconds(5));
                        lock (_lock)
                        {
                            if (_wanted.Count == 0)
                            {
                                break;
                            }
                        }

                        continue;
                    }

                    key = rest.Dequeue();
                }

                if (_fileSystem.File.Exists(PngPath(key)))
                {
                    continue;
                }

                //A document waiting for a picture cannot wait for the user to be away; the background filling can, and does not go on over a metered connection.
                while (!_stop && (!_canDownload() || (!wanted && !_isIdle())))
                {
                    Thread.Sleep(TimeSpan.FromSeconds(_canDownload() ? 2 : 60));
                }

                if (_stop)
                {
                    break;
                }

                try
                {
                    var bytes = _download(string.Format(PngAddress, key)).GetAwaiter().GetResult();
                    if (bytes == null || bytes.Length == 0)
                    {
                        throw new System.IO.InvalidDataException("empty");
                    }

                    Save(PngPath(key), bytes);
                    failures = 0;
                    bool asked;
                    lock (_lock)
                    {
                        asked = _asked.Contains(key);
                    }

                    //Said only for the pictures some document asked for: the rest of the background filling needs no one's attention.
                    if (asked)
                    {
                        _context?.Post(state => PictureArrived?.Invoke((string)state), key);
                    }
                }
                catch (Exception exception) when (exception is HttpRequestException || exception is TaskCanceledException || exception is System.IO.IOException || exception is System.Net.WebException)
                {
                    Logger.Debug("The picture of emoji '{0}' could not be fetched: {1}", key, exception.Message);
                    if (++failures >= FailuresBeforeGivingUp)
                    {
                        Logger.Info("Fetching the emoji pictures stopped after {0} failures in a row; it starts again with the next session.", failures);
                        break;
                    }
                }

                Thread.Sleep(_pause);
            }

            lock (_lock)
            {
                _worker = null;
                _stop = false;
            }
        }

        /// <summary>
        /// The SVG of the emoji, from the folder if it was fetched before, otherwise fetched now when <paramref name="allowDownload"/> says so and the connection is not metered.
        /// <paramref name="ready"/> is called on the thread that asked, once, with the picture; it is never called if there is none (no connection, a metered one, not found).
        /// </summary>
        public void RequestSvg(string key, bool allowDownload, Action<ImageSource> ready)
        {
            var context = SynchronizationContext.Current;
            var path = System.IO.Path.Combine(_svgFolder, FilePrefix + key + ".svg");

            var cachedBytes = ReadFile(path);
            if (cachedBytes != null)
            {
                Deliver(ready, ReadSvg(cachedBytes, key));
                return;
            }

            if (!allowDownload || !_canDownload())
            {
                return;
            }

            Task<byte[]> pending;
            lock (_lock)
            {
                if (!_svgDownloads.TryGetValue(key, out pending))
                {
                    pending = _svgDownloads[key] = Task.Run(() => _download(string.Format(SvgAddress, key)));
                }
            }

            pending.ContinueWith(task =>
            {
                lock (_lock)
                {
                    _svgDownloads.Remove(key);
                }

                if (task.IsFaulted || task.IsCanceled || task.Result == null || task.Result.Length == 0)
                {
                    Logger.Debug("The SVG of emoji '{0}' could not be fetched: {1}", key, task.Exception?.GetBaseException().Message ?? "no content");
                    return;
                }

                var bytes = task.Result;
                Save(path, bytes);
                //Parsed where it is shown: the reader builds WPF objects that belong to the thread that made them.
                context?.Post(state => Deliver(ready, ReadSvg(bytes, key)), null);
            }, TaskScheduler.Default);
        }

        private static void Deliver(Action<ImageSource> ready, ImageSource image)
        {
            if (image != null)
            {
                ready(image);
            }
        }

        private static ImageSource Decode(byte[] bytes, string key)
        {
            try
            {
                using (var memory = new System.IO.MemoryStream(bytes))
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.StreamSource = memory;
                    bitmap.EndInit();
                    bitmap.Freeze();
                    return bitmap;
                }
            }
            catch (Exception exception) when (exception is NotSupportedException || exception is System.IO.FileFormatException || exception is ArgumentException)
            {
                //A file that was cut short or is not a picture: it is not used, and is fetched again.
                Logger.Debug(exception, "The picture of emoji '{0}' could not be read.", key);
                return null;
            }
        }

        private static ImageSource ReadSvg(byte[] bytes, string key)
        {
            try
            {
                var settings = new WpfDrawingSettings { IncludeRuntime = false, TextAsGeometry = true, OptimizePath = true };
                using (var stream = new System.IO.MemoryStream(bytes))
                {
                    var reader = new FileSvgReader(settings);
                    var drawing = reader.Read(stream);
                    if (drawing == null)
                    {
                        return null;
                    }

                    var image = new DrawingImage(drawing);
                    image.Freeze();
                    return image;
                }
            }
            catch (Exception exception) when (!(exception is OutOfMemoryException))
            {
                Logger.Debug(exception, "The SVG of emoji '{0}' could not be read.", key);
                return null;
            }
        }

        private string PngPath(string key) => System.IO.Path.Combine(_pngFolder, FilePrefix + key + ".png");

        private byte[] ReadPackaged(string key)
        {
            var entry = _package?.GetEntry(FilePrefix + key + ".png");
            if (entry == null)
            {
                return null;
            }

            using (var stream = entry.Open())
            using (var memory = new System.IO.MemoryStream())
            {
                stream.CopyTo(memory);
                return memory.ToArray();
            }
        }

        private byte[] ReadFile(string path)
        {
            try
            {
                if (_fileSystem.File.Exists(path))
                {
                    var bytes = _fileSystem.File.ReadAllBytes(path);
                    return bytes.Length > 0 ? bytes : null;
                }
            }
            catch (Exception exception) when (exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                //Treated as not there.
            }

            return null;
        }

        private void Save(string path, byte[] bytes)
        {
            try
            {
                _fileSystem.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path));
                var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                _fileSystem.File.WriteAllBytes(temporary, bytes);
                if (_fileSystem.File.Exists(path))
                {
                    _fileSystem.File.Delete(temporary);
                }
                else
                {
                    _fileSystem.File.Move(temporary, path);
                }
            }
            catch (Exception exception) when (exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                //Not kept; it is fetched again next time.
            }
        }

        private void EnsurePackage()
        {
            lock (_lock)
            {
                if (_keys != null)
                {
                    return;
                }

                try
                {
                    var stream = _openPackage();
                    _package = stream == null ? null : new ZipArchive(stream, ZipArchiveMode.Read);
                }
                catch (Exception exception) when (exception is System.IO.InvalidDataException || exception is System.IO.IOException)
                {
                    Logger.Warn(exception, "The emoji pictures could not be opened; emoji are left to the font.");
                }

                _packaged = new HashSet<string>(
                    (_package?.Entries ?? Enumerable.Empty<ZipArchiveEntry>())
                    .Select(entry => entry.Name)
                    .Where(name => name.StartsWith(FilePrefix, StringComparison.Ordinal) && name.EndsWith(".png", StringComparison.Ordinal))
                    .Select(name => name.Substring(FilePrefix.Length, name.Length - FilePrefix.Length - ".png".Length)));

                _keys = new HashSet<string>(_packaged);
                try
                {
                    using (var stream = _openIndex?.Invoke())
                    {
                        if (stream != null)
                        {
                            using (var reader = new System.IO.StreamReader(stream))
                            {
                                string line;
                                while ((line = reader.ReadLine()) != null)
                                {
                                    if (line.Trim().Length > 0)
                                    {
                                        _keys.Add(line.Trim());
                                    }
                                }
                            }
                        }
                    }
                }
                catch (System.IO.IOException exception)
                {
                    Logger.Warn(exception, "The index of the emoji could not be read; only the packaged ones are used.");
                }
            }
        }

        private static readonly Lazy<HttpClient> Client = new Lazy<HttpClient>(() =>
        {
            //The runtime of an add-in built for .NET Framework 4.6 does not offer TLS 1.2 by default; GitHub requires it.
            System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
            return new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        });

        private static async Task<byte[]> HttpDownload(string address)
        {
            using (var response = await Client.Value.GetAsync(address).ConfigureAwait(false))
            {
                return response.IsSuccessStatusCode ? await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false) : null;
            }
        }
    }
}
