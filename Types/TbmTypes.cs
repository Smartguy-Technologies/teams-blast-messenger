// Teams Blast Messenger - shared .NET types.
// Compiled once per AppDomain with Add-Type so every runspace (UI + background workers)
// sees the same type identity. Written for C# 5 (the compiler used by Windows PowerShell 5.1):
// no string interpolation, no null-conditional operators, no auto-property initialisers.

using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Management.Automation;
using System.Xml;

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
        private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);
        private static StreamWriter _writer;
        private static string _writerPath;
        private static bool _exitHooked;

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
                CloseWriter();
                System.IO.Directory.CreateDirectory(directory);
                _directory = directory;
                _minimumLevel = minimumLevel;
            }
            if (retentionDays > 0) { Purge(retentionDays); }
        }

        public static void Purge(int retentionDays)
        {
            if (_directory == null) { return; }
            lock (Sync) { CloseWriter(); }
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
                    if (_writer == null || !string.Equals(_writerPath, path, StringComparison.OrdinalIgnoreCase)) { OpenWriter(path); }
                    // Another app instance may share the file: always append at the current end.
                    _writer.BaseStream.Seek(0, SeekOrigin.End);
                    _writer.Write(entry.ToFileLine());
                    _writer.Write(Environment.NewLine);
                    _writer.Flush();
                }
                catch (IOException) { CloseWriter(); }
                catch (UnauthorizedAccessException) { CloseWriter(); }
            }
        }

        /// <summary>Flushes and closes the open log file (it reopens on the next write).</summary>
        public static void Close()
        {
            lock (Sync) { CloseWriter(); }
        }

        private static void OpenWriter(string path)
        {
            CloseWriter();
            FileStream stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);
            _writer = new StreamWriter(stream, Utf8NoBom);
            _writerPath = path;
            if (!_exitHooked)
            {
                AppDomain.CurrentDomain.ProcessExit += delegate { Close(); };
                _exitHooked = true;
            }
        }

        private static void CloseWriter()
        {
            if (_writer == null) { return; }
            try { _writer.Dispose(); }
            catch (IOException) { }
            catch (ObjectDisposedException) { }
            _writer = null;
            _writerPath = null;
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

    /// <summary>One typed/pasted recipient in the "Add recipients" dialog.</summary>
    public sealed class TbmManualEntry
    {
        public string Raw { get; set; }
        public string Name { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        /// <summary>OK, Invalid email, Duplicate, Not found, Ambiguous, Needs email.</summary>
        public string Status { get; set; }
        public string Detail { get; set; }
        public bool IsValid { get { return Status == "OK"; } }
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

        /// <summary>
        /// True for a record loaded from the history summary file: list columns and counts only, no
        /// recipients or interactions. Load the full record (CopyFrom) before using its lists.
        /// </summary>
        public bool IsSummaryOnly { get; private set; }
        /// <summary>The history file this record was loaded from (set by the history loader).</summary>
        public string FilePath { get; set; }
        private int _summaryRecipientCount;
        private int _summarySentCount;
        private int _summaryFailedCount;
        private int _summaryRespondedCount;

        public void SetSummaryCounts(int recipientCount, int sentCount, int failedCount, int respondedCount)
        {
            _summaryRecipientCount = recipientCount;
            _summarySentCount = sentCount;
            _summaryFailedCount = failedCount;
            _summaryRespondedCount = respondedCount;
            IsSummaryOnly = true;
        }

        /// <summary>Replaces this record's contents with a fully loaded copy, keeping object identity (grid selection).</summary>
        public void CopyFrom(TbmBlastRecord other)
        {
            if (other == null) { throw new ArgumentNullException("other"); }
            Id = other.Id;
            CreatedUtc = other.CreatedUtc;
            CompletedUtc = other.CompletedUtc;
            Mode = other.Mode;
            Transport = other.Transport;
            Template = other.Template;
            Format = other.Format;
            SourceFile = other.SourceFile;
            Topic = other.Topic;
            SenderId = other.SenderId;
            SenderUpn = other.SenderUpn;
            IsTest = other.IsTest;
            LastInteractionCheckUtc = other.LastInteractionCheckUtc;
            Recipients = other.Recipients;
            Interactions = other.Interactions;
            if (other.FilePath != null) { FilePath = other.FilePath; }
            IsSummaryOnly = false;
        }

        public int RecipientCount { get { return IsSummaryOnly ? _summaryRecipientCount : Recipients.Count; } }
        public int SentCount { get { return IsSummaryOnly ? _summarySentCount : CountStatus(TbmRecipientStatus.Sent); } }
        public int FailedCount { get { return IsSummaryOnly ? _summaryFailedCount : CountStatus(TbmRecipientStatus.Failed); } }
        public int RespondedCount
        {
            get
            {
                if (IsSummaryOnly) { return _summaryRespondedCount; }
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
        /// <summary>
        /// True for a single, plain email address (no display name, no list separators) with a dotted domain.
        /// </summary>
        public static bool IsEmailAddress(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) { return false; }
            string trimmed = value.Trim();
            if (trimmed.IndexOf(' ') >= 0 || trimmed.IndexOf(',') >= 0 || trimmed.IndexOf(';') >= 0) { return false; }
            try
            {
                System.Net.Mail.MailAddress address = new System.Net.Mail.MailAddress(trimmed);
                return string.Equals(address.Address, trimmed, StringComparison.OrdinalIgnoreCase) && address.Host.IndexOf('.') >= 0;
            }
            catch (FormatException) { return false; }
            catch (ArgumentException) { return false; }
        }

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
            text = WebUtility.HtmlDecode(text).Replace('\u00A0', ' ');
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
            if (maxLength <= 3) { return text.Substring(0, Math.Max(0, maxLength)); }
            return text.Substring(0, maxLength - 3) + "...";
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

    /// <summary>Sorts blasts newest first.</summary>
    public sealed class TbmBlastRecordByCreatedDesc : IComparer<TbmBlastRecord>
    {
        public int Compare(TbmBlastRecord x, TbmBlastRecord y)
        {
            if (ReferenceEquals(x, y)) { return 0; }
            if (x == null) { return 1; }
            if (y == null) { return -1; }
            return y.CreatedUtc.CompareTo(x.CreatedUtc);
        }
    }

    /// <summary>Sorts modules newest version first.</summary>
    public sealed class TbmModuleByVersionDesc : IComparer<PSModuleInfo>
    {
        public int Compare(PSModuleInfo x, PSModuleInfo y)
        {
            if (ReferenceEquals(x, y)) { return 0; }
            if (x == null) { return 1; }
            if (y == null) { return -1; }
            if (y.Version == null) { return x.Version == null ? 0 : -1; }
            return y.Version.CompareTo(x.Version);
        }
    }

    /// <summary>Sorts replies oldest first.</summary>
    public sealed class TbmReplyByCreated : IComparer<TbmReplyMessage>
    {
        public int Compare(TbmReplyMessage x, TbmReplyMessage y)
        {
            if (ReferenceEquals(x, y)) { return 0; }
            if (x == null) { return -1; }
            if (y == null) { return 1; }
            return x.CreatedUtc.CompareTo(y.CreatedUtc);
        }
    }

    /// <summary>
    /// Fast, StrictMode-safe readers for parsed JSON (PSCustomObject / hashtable) used in per-message loops,
    /// where a PowerShell function call per field would dominate the run time.
    /// </summary>
    public static class TbmObject
    {
        private static readonly char[] PathSeparator = new char[] { '.' };

        /// <summary>Reads a dot-separated property path; returns null when any step is missing.</summary>
        public static object GetValue(object input, string path)
        {
            if (string.IsNullOrEmpty(path)) { throw new ArgumentNullException("path"); }
            object current = input;
            foreach (string part in path.Split(PathSeparator))
            {
                if (current == null) { return null; }
                PSObject wrapper = current as PSObject;
                object baseObject = wrapper != null ? wrapper.BaseObject : current;
                IDictionary dictionary = baseObject as IDictionary;
                if (dictionary != null)
                {
                    if (!dictionary.Contains(part)) { return null; }
                    current = dictionary[part];
                    continue;
                }
                if (wrapper == null) { wrapper = PSObject.AsPSObject(current); }
                PSPropertyInfo property = wrapper.Properties[part];
                if (property == null) { return null; }
                current = property.Value;
            }
            return current;
        }

        /// <summary>GetValue converted to string; null stays null.</summary>
        public static string GetString(object input, string path)
        {
            object value = GetValue(input, path);
            if (value == null) { return null; }
            PSObject wrapper = value as PSObject;
            if (wrapper != null) { value = wrapper.BaseObject; }
            if (value is DateTime) { return ((DateTime)value).ToString("o", CultureInfo.InvariantCulture); }
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Converts a DateTime, DateTimeOffset or ISO-8601 string to UTC (Unspecified kind is taken as UTC).
        /// Returns null for null/blank input; throws FormatException for text that is not a date.
        /// </summary>
        public static DateTime? ToUtc(object value)
        {
            if (value == null) { return null; }
            PSObject wrapper = value as PSObject;
            if (wrapper != null) { value = wrapper.BaseObject; }
            if (value is DateTime)
            {
                DateTime date = (DateTime)value;
                if (date.Kind == DateTimeKind.Unspecified) { return DateTime.SpecifyKind(date, DateTimeKind.Utc); }
                return date.ToUniversalTime();
            }
            if (value is DateTimeOffset) { return ((DateTimeOffset)value).UtcDateTime; }
            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(text)) { return null; }
            return DateTimeOffset.Parse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal).UtcDateTime;
        }
    }

    /// <summary>
    /// Streaming reader for the parts of an .xlsx package that hold cell data. The worksheet is read with
    /// XmlReader (forward only, no DOM), so memory and time stay proportional to the sheet size.
    /// </summary>
    public static class TbmXlsx
    {
        public const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly Regex FormatLiterals = new Regex(@"""[^""]*""|\\.|_.|\*.|\[(?![hms]\])[^\]]*\]", RegexOptions.Compiled);

        public static XmlReaderSettings CreateReaderSettings()
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Prohibit;
            settings.XmlResolver = null;
            settings.IgnoreComments = true;
            settings.IgnoreProcessingInstructions = true;
            return settings;
        }

        /// <summary>Classifies an Excel number format as Date, Time, DateTime, Percent or Number.</summary>
        public static string GetNumberFormatKind(int formatId, string formatCode)
        {
            if ((formatId >= 14 && formatId <= 17) || (formatId >= 27 && formatId <= 31) || (formatId >= 34 && formatId <= 36) || (formatId >= 50 && formatId <= 58)) { return "Date"; }
            if (formatId == 22) { return "DateTime"; }
            if ((formatId >= 18 && formatId <= 21) || (formatId >= 45 && formatId <= 47) || formatId == 32 || formatId == 33) { return "Time"; }
            if (formatId == 9 || formatId == 10) { return "Percent"; }
            if (string.IsNullOrWhiteSpace(formatCode) || formatCode == "General" || formatCode == "@") { return "Number"; }

            // Strip literals ("text", \x, _x, *x) and [colour]/[locale] sections before looking for date/time letters.
            string code = FormatLiterals.Replace(formatCode.Split(';')[0], string.Empty).ToLowerInvariant();
            bool hasTime = code.Contains("h") || code.Contains("s");
            bool hasDate = code.Contains("y") || code.Contains("d") || (code.Contains("m") && !hasTime);
            if (hasDate && hasTime) { return "DateTime"; }
            if (hasDate) { return "Date"; }
            if (hasTime) { return "Time"; }
            if (code.Contains("%")) { return "Percent"; }
            return "Number";
        }

        /// <summary>Formats a numeric cell value for its number format kind; out-of-range dates fall back to the number.</summary>
        public static string FormatNumber(double value, string kind, bool date1904)
        {
            CultureInfo culture = CultureInfo.CurrentCulture;
            try
            {
                double serial = value;
                if (date1904 && kind != "Number" && kind != "Percent") { serial = value + 1462; }
                switch (kind)
                {
                    case "Date":
                        DateTime date = DateTime.FromOADate(serial);
                        if (date.TimeOfDay.Ticks != 0) { return date.ToString("g", culture); }
                        return date.ToString("d", culture);
                    case "DateTime":
                        return DateTime.FromOADate(serial).ToString("g", culture);
                    case "Time":
                        return DateTime.FromOADate(serial).ToString("t", culture);
                    case "Percent":
                        return (value * 100).ToString("0.##", culture) + "%";
                    default:
                        return value.ToString("G15", culture);
                }
            }
            catch (ArgumentException)
            {
                return value.ToString("G15", culture);
            }
        }

        /// <summary>Zero-based column index from the letters of an A1 reference ("C7" -> 2); -1 when there are none.</summary>
        public static int ColumnIndex(string reference)
        {
            int index = 0;
            if (reference == null) { return -1; }
            foreach (char character in reference)
            {
                if (character >= 'A' && character <= 'Z') { index = (index * 26) + (character - 64); }
                else if (character >= 'a' && character <= 'z') { index = (index * 26) + (character - 96); }
                else { break; }
            }
            return index - 1;
        }

        public static List<string> ReadSharedStrings(Stream stream)
        {
            List<string> strings = new List<string>();
            using (XmlReader reader = XmlReader.Create(stream, CreateReaderSettings()))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1 && reader.LocalName == "si") { strings.Add(ReadRichText(reader)); }
                }
            }
            return strings;
        }

        /// <summary>Maps each cell style index (cellXfs) to a number format kind.</summary>
        public static List<string> ReadStyleKinds(Stream stream)
        {
            XmlDocument document = new XmlDocument();
            document.XmlResolver = null;
            using (XmlReader reader = XmlReader.Create(stream, CreateReaderSettings())) { document.Load(reader); }
            XmlNamespaceManager ns = new XmlNamespaceManager(document.NameTable);
            ns.AddNamespace("m", SpreadsheetNs);

            Dictionary<int, string> customFormats = new Dictionary<int, string>();
            foreach (XmlNode node in document.SelectNodes("/m:styleSheet/m:numFmts/m:numFmt", ns))
            {
                XmlElement format = node as XmlElement;
                int id;
                if (format != null && int.TryParse(format.GetAttribute("numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out id)) { customFormats[id] = format.GetAttribute("formatCode"); }
            }
            List<string> kinds = new List<string>();
            foreach (XmlNode node in document.SelectNodes("/m:styleSheet/m:cellXfs/m:xf", ns))
            {
                XmlElement xf = node as XmlElement;
                int formatId = 0;
                if (xf != null) { int.TryParse(xf.GetAttribute("numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out formatId); }
                string code;
                customFormats.TryGetValue(formatId, out code);
                kinds.Add(GetNumberFormatKind(formatId, code));
            }
            return kinds;
        }

        /// <summary>Reads every row of /worksheet/sheetData as a string array indexed by column (null = empty cell).</summary>
        public static List<string[]> ReadSheet(Stream stream, IList<string> sharedStrings, IList<string> styleKinds, bool date1904)
        {
            List<string[]> rows = new List<string[]>();
            Dictionary<int, string> cells = new Dictionary<int, string>();
            using (XmlReader reader = XmlReader.Create(stream, CreateReaderSettings()))
            {
                string level1 = null;
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element) { continue; }
                    if (reader.Depth == 1) { level1 = reader.LocalName; continue; }
                    if (reader.Depth == 2 && level1 == "sheetData" && reader.LocalName == "row") { rows.Add(ReadRow(reader, cells, sharedStrings, styleKinds, date1904)); }
                }
            }
            return rows;
        }

        private static string[] ReadRow(XmlReader reader, Dictionary<int, string> cells, IList<string> sharedStrings, IList<string> styleKinds, bool date1904)
        {
            cells.Clear();
            int maxColumn = -1;
            int position = 0;
            if (!reader.IsEmptyElement)
            {
                int depth = reader.Depth;
                reader.Read();
                while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.Depth == depth + 1 && reader.LocalName == "c")
                    {
                        int column = position;
                        string reference = reader.GetAttribute("r");
                        if (!string.IsNullOrEmpty(reference))
                        {
                            int fromReference = ColumnIndex(reference);
                            if (fromReference >= 0) { column = fromReference; }
                        }
                        position = column + 1;
                        string text = ReadCell(reader, sharedStrings, styleKinds, date1904);
                        if (text != null)
                        {
                            cells[column] = text;
                            if (column > maxColumn) { maxColumn = column; }
                        }
                    }
                    reader.Read();
                }
            }
            string[] values = new string[maxColumn + 1];
            foreach (KeyValuePair<int, string> pair in cells) { values[pair.Key] = pair.Value; }
            return values;
        }

        /// <summary>Reads one &lt;c&gt; element; leaves the reader on its end (or empty) element.</summary>
        private static string ReadCell(XmlReader reader, IList<string> sharedStrings, IList<string> styleKinds, bool date1904)
        {
            string type = reader.GetAttribute("t") ?? string.Empty;
            string style = reader.GetAttribute("s");
            string raw = null;
            string inline = null;
            bool hasInline = false;
            if (!reader.IsEmptyElement)
            {
                int depth = reader.Depth;
                reader.Read();
                while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.Depth == depth + 1)
                    {
                        if (reader.LocalName == "v") { raw = ReadText(reader); continue; }
                        if (reader.LocalName == "is") { inline = ReadRichText(reader); hasInline = true; }
                    }
                    reader.Read();
                }
            }

            if (string.Equals(type, "s", StringComparison.OrdinalIgnoreCase))
            {
                int index;
                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out index) && index >= 0 && index < sharedStrings.Count) { return sharedStrings[index]; }
                return null;
            }
            if (string.Equals(type, "inlineStr", StringComparison.OrdinalIgnoreCase)) { return hasInline ? inline : null; }
            if (string.Equals(type, "str", StringComparison.OrdinalIgnoreCase)) { return raw; }
            if (string.Equals(type, "b", StringComparison.OrdinalIgnoreCase)) { return raw == "1" ? "TRUE" : "FALSE"; }
            if (string.Equals(type, "e", StringComparison.OrdinalIgnoreCase)) { return raw; }

            if (string.IsNullOrEmpty(raw)) { return null; }
            double number;
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) { return raw; }
            string kind = "Number";
            int styleIndex;
            if (!string.IsNullOrEmpty(style) && int.TryParse(style, NumberStyles.Integer, CultureInfo.InvariantCulture, out styleIndex) && styleIndex >= 0 && styleIndex < styleKinds.Count) { kind = styleKinds[styleIndex]; }
            return FormatNumber(number, kind, date1904);
        }

        /// <summary>
        /// Concatenates the &lt;t&gt; runs of a shared or inline string, ignoring phonetic (&lt;rPh&gt;) runs.
        /// Positioned on the si/is start element; leaves the reader on its end (or empty) element.
        /// </summary>
        private static string ReadRichText(XmlReader reader)
        {
            if (reader.IsEmptyElement) { return string.Empty; }
            StringBuilder builder = new StringBuilder();
            int depth = reader.Depth;
            string child = null;
            reader.Read();
            while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.Depth == depth + 1) { child = reader.LocalName; }
                    bool isText = reader.LocalName == "t" && (reader.Depth == depth + 1 || (reader.Depth == depth + 2 && child == "r"));
                    if (isText) { builder.Append(ReadText(reader)); continue; }
                }
                reader.Read();
            }
            return builder.ToString();
        }

        /// <summary>
        /// Text content of a simple element, moving past its end. Whitespace-only text counts only under
        /// xml:space="preserve", matching how the workbook's XML is normally interpreted.
        /// </summary>
        private static string ReadText(XmlReader reader)
        {
            bool preserve = reader.XmlSpace == XmlSpace.Preserve || reader.GetAttribute("xml:space") == "preserve";
            string text = reader.ReadElementContentAsString();
            if (!preserve && IsXmlWhitespace(text)) { return string.Empty; }
            return text;
        }

        private static bool IsXmlWhitespace(string text)
        {
            foreach (char character in text)
            {
                if (character != ' ' && character != '\t' && character != '\r' && character != '\n') { return false; }
            }
            return true;
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
