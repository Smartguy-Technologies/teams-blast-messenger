// Teams Blast Messenger - shared .NET types.
// Compiled once per AppDomain with Add-Type so every runspace (UI + background workers)
// sees the same type identity. Written for C# 5 (the compiler used by Windows PowerShell 5.1):
// no string interpolation, no null-conditional operators, no auto-property initialisers.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace TeamsBlastMessenger
{
    public enum TbmLogLevel
    {
        Debug = 0,
        Info = 1,
        Warn = 2,
        Error = 3
    }

    public enum TbmRecipientStatus
    {
        Pending,
        Sent,
        Failed,
        Skipped,
        Cancelled
    }

    public sealed class TbmLogEntry
    {
        public DateTime TimestampUtc { get; set; }
        public TbmLogLevel Level { get; set; }
        public string Source { get; set; }
        public string Message { get; set; }
        public int ThreadId { get; set; }

        public string ToFileLine()
        {
            return string.Format(CultureInfo.InvariantCulture, "{0:yyyy-MM-ddTHH:mm:ss.fffZ} [{1,-5}] [T{2:D3}] [{3}] {4}",
                TimestampUtc, Level.ToString().ToUpperInvariant(), ThreadId, Source, Message);
        }

        public string ToDisplayLine()
        {
            return string.Format(CultureInfo.CurrentCulture, "{0:HH:mm:ss} {1,-5} {2}",
                TimestampUtc.ToLocalTime(), Level.ToString().ToUpperInvariant(), Message);
        }
    }

    /// <summary>
    /// Thread-safe, process-wide file logger. Each entry is appended under a lock and also queued
    /// so the UI thread can drain and display entries (no cross-thread UI access, no PowerShell
    /// script blocks invoked from foreign threads).
    /// </summary>
    public static class TbmLogger
    {
        private static readonly object Sync = new object();
        private static readonly ConcurrentQueue<TbmLogEntry> Pending = new ConcurrentQueue<TbmLogEntry>();
        private static string _directory;
        private static TbmLogLevel _minimumLevel = TbmLogLevel.Info;
        private const int MaxPending = 5000;

        public static string Directory { get { return _directory; } }
        public static TbmLogLevel MinimumLevel { get { return _minimumLevel; } set { _minimumLevel = value; } }
        public static bool IsInitialized { get { return _directory != null; } }

        public static string CurrentLogPath
        {
            get
            {
                if (_directory == null) { return null; }
                return Path.Combine(_directory, "TBM_" + DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".log");
            }
        }

        public static void Initialize(string directory, TbmLogLevel minimumLevel, int retentionDays)
        {
            if (string.IsNullOrEmpty(directory)) { throw new ArgumentNullException("directory"); }
            lock (Sync)
            {
                System.IO.Directory.CreateDirectory(directory);
                _directory = directory;
                _minimumLevel = minimumLevel;
            }
            if (retentionDays > 0) { Purge(retentionDays); }
        }

        public static void Purge(int retentionDays)
        {
            if (_directory == null) { return; }
            DateTime cutoff = DateTime.Now.AddDays(-retentionDays);
            foreach (string file in System.IO.Directory.GetFiles(_directory, "TBM_*.log"))
            {
                try
                {
                    if (File.GetLastWriteTime(file) < cutoff) { File.Delete(file); }
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }

        public static void Write(TbmLogLevel level, string source, string message)
        {
            if (level < _minimumLevel) { return; }
            TbmLogEntry entry = new TbmLogEntry();
            entry.TimestampUtc = DateTime.UtcNow;
            entry.Level = level;
            entry.Source = string.IsNullOrEmpty(source) ? "General" : source;
            entry.Message = message ?? string.Empty;
            entry.ThreadId = Thread.CurrentThread.ManagedThreadId;

            if (Pending.Count < MaxPending) { Pending.Enqueue(entry); }

            string path = CurrentLogPath;
            if (path == null) { return; }
            lock (Sync)
            {
                try
                {
                    File.AppendAllText(path, entry.ToFileLine() + Environment.NewLine, new UTF8Encoding(false));
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }

        public static void Debug(string source, string message) { Write(TbmLogLevel.Debug, source, message); }
        public static void Info(string source, string message) { Write(TbmLogLevel.Info, source, message); }
        public static void Warn(string source, string message) { Write(TbmLogLevel.Warn, source, message); }
        public static void Error(string source, string message) { Write(TbmLogLevel.Error, source, message); }

        public static void Error(string source, string message, Exception exception)
        {
            string detail = exception == null ? string.Empty : Environment.NewLine + exception.ToString();
            Write(TbmLogLevel.Error, source, message + detail);
        }

        /// <summary>Returns and removes all entries queued since the last call.</summary>
        public static List<TbmLogEntry> DrainPending()
        {
            List<TbmLogEntry> result = new List<TbmLogEntry>();
            TbmLogEntry entry;
            while (Pending.TryDequeue(out entry)) { result.Add(entry); }
            return result;
        }
    }

    /// <summary>A failed Microsoft Graph call with the details Graph returned.</summary>
    public sealed class TbmGraphException : Exception
    {
        public int StatusCode { get; private set; }
        public string ErrorCode { get; private set; }
        public string RequestId { get; private set; }
        public string Method { get; private set; }
        public string Path { get; private set; }

        public TbmGraphException(int statusCode, string errorCode, string message, string requestId, string method, string path)
            : base(string.Format(CultureInfo.InvariantCulture, "Graph {0} {1} failed ({2} {3}): {4}", method, path, statusCode, errorCode, message))
        {
            StatusCode = statusCode;
            ErrorCode = errorCode;
            RequestId = requestId;
            Method = method;
            Path = path;
        }
    }

    public sealed class TbmRecipientResult
    {
        public TbmRecipientResult()
        {
            Status = TbmRecipientStatus.Pending;
        }

        public int RowIndex { get; set; }
        public string Email { get; set; }
        public string UserId { get; set; }
        public string DisplayName { get; set; }
        public string ChatId { get; set; }
        public string MessageId { get; set; }
        public DateTime? SentUtc { get; set; }
        public TbmRecipientStatus Status { get; set; }
        public string Error { get; set; }
        /// <summary>Template after token expansion, in the blast's Format (Text or Html).</summary>
        public string RenderedMessage { get; set; }

        public string SentLocal { get { return TbmText.FormatLocal(SentUtc); } }
        public string StatusText { get { return Status.ToString(); } }
    }

    public sealed class TbmReplyMessage
    {
        public string MessageId { get; set; }
        public string FromUserId { get; set; }
        public string FromName { get; set; }
        public DateTime CreatedUtc { get; set; }
        public string Text { get; set; }
        public bool IsDirectReply { get; set; }

        public string CreatedLocal { get { return TbmText.FormatLocal(CreatedUtc); } }
    }

    public sealed class TbmInteractionResult
    {
        public TbmInteractionResult()
        {
            Replies = new List<TbmReplyMessage>();
        }

        public string Email { get; set; }
        public string DisplayName { get; set; }
        public string UserId { get; set; }
        public string ChatId { get; set; }
        public DateTime? SentUtc { get; set; }
        public bool Responded { get; set; }
        public int ReplyCount { get; set; }
        public DateTime? FirstReplyUtc { get; set; }
        public DateTime? LastReplyUtc { get; set; }
        public string FirstReplyText { get; set; }
        public bool DirectReply { get; set; }
        public string Reactions { get; set; }
        public string Error { get; set; }
        public List<TbmReplyMessage> Replies { get; set; }

        public TimeSpan? ResponseTime
        {
            get
            {
                if (!SentUtc.HasValue || !FirstReplyUtc.HasValue) { return null; }
                TimeSpan span = FirstReplyUtc.Value - SentUtc.Value;
                return span < TimeSpan.Zero ? TimeSpan.Zero : span;
            }
        }

        public string ResponseTimeDisplay { get { return TbmText.FormatDuration(ResponseTime); } }
        public string SentLocal { get { return TbmText.FormatLocal(SentUtc); } }
        public string FirstReplyLocal { get { return TbmText.FormatLocal(FirstReplyUtc); } }
        public string RespondedText { get { return Responded ? "Yes" : "No"; } }
        public string Recipient
        {
            get { return string.IsNullOrEmpty(DisplayName) ? Email : DisplayName + " <" + Email + ">"; }
        }
    }

    public sealed class TbmBlastRecord
    {
        public TbmBlastRecord()
        {
            Id = Guid.NewGuid().ToString();
            CreatedUtc = DateTime.UtcNow;
            Recipients = new List<TbmRecipientResult>();
            Interactions = new List<TbmInteractionResult>();
        }

        public string Id { get; set; }
        public DateTime CreatedUtc { get; set; }
        public DateTime? CompletedUtc { get; set; }
        /// <summary>Individual or Group.</summary>
        public string Mode { get; set; }
        /// <summary>GraphSdk, GraphRest, DeepLinkAuto or DeepLinkManual.</summary>
        public string Transport { get; set; }
        public string Template { get; set; }
        /// <summary>Text or Html.</summary>
        public string Format { get; set; }
        public string SourceFile { get; set; }
        public string Topic { get; set; }
        public string SenderId { get; set; }
        public string SenderUpn { get; set; }
        public bool IsTest { get; set; }
        public DateTime? LastInteractionCheckUtc { get; set; }
        public List<TbmRecipientResult> Recipients { get; set; }
        public List<TbmInteractionResult> Interactions { get; set; }

        public int RecipientCount { get { return Recipients.Count; } }
        public int SentCount { get { return CountStatus(TbmRecipientStatus.Sent); } }
        public int FailedCount { get { return CountStatus(TbmRecipientStatus.Failed); } }
        public int RespondedCount
        {
            get
            {
                int count = 0;
                foreach (TbmInteractionResult item in Interactions) { if (item.Responded) { count++; } }
                return count;
            }
        }
        public string CreatedLocal { get { return TbmText.FormatLocal(CreatedUtc); } }
        public string TemplatePreview { get { return TbmText.Truncate(TbmText.SingleLine(Template), 80); } }
        public string ResponseSummary
        {
            get
            {
                if (!LastInteractionCheckUtc.HasValue) { return string.Empty; }
                return string.Format(CultureInfo.CurrentCulture, "{0}/{1}", RespondedCount, SentCount);
            }
        }

        private int CountStatus(TbmRecipientStatus status)
        {
            int count = 0;
            foreach (TbmRecipientResult item in Recipients) { if (item.Status == status) { count++; } }
            return count;
        }
    }

    /// <summary>Text helpers that are hot paths or must behave identically in every runspace.</summary>
    public static class TbmText
    {
        private static readonly Regex BreakTags = new Regex(@"<\s*(br|/p|/div|/li|/h[1-6])\s*/?\s*>", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex AnyTag = new Regex(@"<[^>]+>", RegexOptions.Compiled);
        private static readonly Regex BlankRuns = new Regex(@"[ \t]+", RegexOptions.Compiled);
        private static readonly Regex NewlineRuns = new Regex(@"(\r?\n\s*){3,}", RegexOptions.Compiled);
        private static readonly Regex Whitespace = new Regex(@"\s+", RegexOptions.Compiled);

        /// <summary>Converts Teams/Graph HTML message content to readable plain text.</summary>
        public static string StripHtml(string html)
        {
            if (string.IsNullOrEmpty(html)) { return string.Empty; }
            string text = html.Replace("\r", string.Empty).Replace("\n", string.Empty);
            text = BreakTags.Replace(text, "\n");
            text = AnyTag.Replace(text, string.Empty);
            text = WebUtility.HtmlDecode(text).Replace(' ', ' ');
            text = BlankRuns.Replace(text, " ");
            text = NewlineRuns.Replace(text, "\n\n");
            return text.Trim();
        }

        /// <summary>Plain text to Teams-safe HTML: encodes markup and preserves line breaks.</summary>
        public static string TextToHtml(string text)
        {
            if (string.IsNullOrEmpty(text)) { return string.Empty; }
            string encoded = WebUtility.HtmlEncode(text.Replace("\r\n", "\n").Replace('\r', '\n'));
            return encoded.Replace("\n", "<br>");
        }

        public static string SingleLine(string text)
        {
            if (string.IsNullOrEmpty(text)) { return string.Empty; }
            return Whitespace.Replace(text, " ").Trim();
        }

        public static string Truncate(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength) { return text ?? string.Empty; }
            return text.Substring(0, Math.Max(0, maxLength - 1)) + "…";
        }

        public static string FormatLocal(DateTime? utc)
        {
            if (!utc.HasValue) { return string.Empty; }
            DateTime value = utc.Value.Kind == DateTimeKind.Local ? utc.Value : DateTime.SpecifyKind(utc.Value, DateTimeKind.Utc).ToLocalTime();
            return value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture);
        }

        /// <summary>Formats a duration as "2d 03:14:05" / "03:14:05".</summary>
        public static string FormatDuration(TimeSpan? span)
        {
            if (!span.HasValue) { return string.Empty; }
            TimeSpan value = span.Value;
            if (value.TotalDays >= 1)
            {
                return string.Format(CultureInfo.InvariantCulture, "{0}d {1:D2}:{2:D2}:{3:D2}", (int)value.TotalDays, value.Hours, value.Minutes, value.Seconds);
            }
            return string.Format(CultureInfo.InvariantCulture, "{0:D2}:{1:D2}:{2:D2}", value.Hours, value.Minutes, value.Seconds);
        }

        public static TimeSpan? Median(List<TimeSpan> values)
        {
            if (values == null || values.Count == 0) { return null; }
            List<TimeSpan> sorted = new List<TimeSpan>(values);
            sorted.Sort();
            int middle = sorted.Count / 2;
            if (sorted.Count % 2 == 1) { return sorted[middle]; }
            return TimeSpan.FromTicks((sorted[middle - 1].Ticks + sorted[middle].Ticks) / 2);
        }

        public static TimeSpan? Average(List<TimeSpan> values)
        {
            if (values == null || values.Count == 0) { return null; }
            long total = 0;
            foreach (TimeSpan value in values) { total += value.Ticks; }
            return TimeSpan.FromTicks(total / values.Count);
        }

        /// <summary>RFC 4180 field quoting.</summary>
        public static string CsvField(string value)
        {
            if (value == null) { return string.Empty; }
            bool mustQuote = value.IndexOfAny(new char[] { ',', '"', '\r', '\n' }) >= 0 || value.StartsWith(" ") || value.EndsWith(" ");
            if (!mustQuote) { return value; }
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }

    /// <summary>Win32 helpers used by the deep-link auto-send mode.</summary>
    public static class TbmNative
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        private const int SW_RESTORE = 9;

        public static bool Activate(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) { return false; }
            if (IsIconic(hWnd)) { ShowWindow(hWnd, SW_RESTORE); }
            return SetForegroundWindow(hWnd);
        }

        public static int GetForegroundProcessId()
        {
            IntPtr hWnd = GetForegroundWindow();
            if (hWnd == IntPtr.Zero) { return 0; }
            uint processId;
            GetWindowThreadProcessId(hWnd, out processId);
            return (int)processId;
        }
    }
}
