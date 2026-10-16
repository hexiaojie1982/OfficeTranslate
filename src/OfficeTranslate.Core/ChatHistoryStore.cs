using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace OfficeTranslate.Core
{
    public sealed class ChatHistoryTurn
    {
        public string Timestamp { get; set; } = "";
        public string Question { get; set; } = "";
        public string Answer { get; set; } = "";
    }
    public sealed class ChatHistoryRecord
    {
        public int SchemaVersion { get; set; } = 1;
        public string SessionId { get; set; } = "";
        public string CreatedAt { get; set; } = "";
        public string DocumentName { get; set; } = "";
        public List<ChatHistoryTurn> Turns { get; set; } = new List<ChatHistoryTurn>();
    }
    public sealed class ChatHistorySummary
    {
        public string SessionId { get; internal set; } = "";
        public DateTimeOffset CreatedAt { get; internal set; }
        public string DocumentName { get; internal set; } = "";
        public string Preview { get; internal set; } = "";
    }
    // Local history only. Never send these records to diagnostics or transport.
    public sealed class ChatHistoryStore
    {
        private static readonly object Sync = new object();
        private readonly string _directory;
        public ChatHistoryStore(string? directory = null) => _directory = directory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OfficeTranslate", "ChatHistory");
        private static JavaScriptSerializer Json() => new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private string FilePath(string id)
        {
            if (!Guid.TryParseExact(id, "N", out _)) throw new ArgumentException("Invalid session ID.");
            return Path.Combine(_directory, id + ".json");
        }
        public ChatHistoryRecord? Read(string id)
        {
            lock (Sync)
            {
                var path = FilePath(id);
                if (!File.Exists(path)) return null;
                var record = Json().Deserialize<ChatHistoryRecord>(File.ReadAllText(path, Encoding.UTF8));
                if (record == null || record.SchemaVersion != 1 || record.SessionId != id ||
                    !DateTimeOffset.TryParse(record.CreatedAt, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out _) ||
                    record.DocumentName == null || record.Turns == null || record.Turns.Count == 0 ||
                    record.Turns.Any(t => t == null || t.Question == null || t.Answer == null ||
                        !DateTimeOffset.TryParse(t.Timestamp, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out _)))
                    throw new InvalidDataException("Invalid chat history.");
                return record;
            }
        }
        public IReadOnlyList<ChatHistorySummary> List(out int unreadable)
        {
            lock (Sync)
            {
                unreadable = 0;
                var result = new List<ChatHistorySummary>();
                if (!Directory.Exists(_directory)) return result;
                foreach (var path in Directory.GetFiles(_directory, "*.json"))
                {
                    try
                    {
                        var record = Read(Path.GetFileNameWithoutExtension(path));
                        if (record == null) continue;
                        var preview = record.Turns[0].Question.Replace('\r', ' ').Replace('\n', ' ');
                        result.Add(new ChatHistorySummary { SessionId = record.SessionId, DocumentName = record.DocumentName,
                            CreatedAt = DateTimeOffset.Parse(record.CreatedAt, CultureInfo.InvariantCulture),
                            Preview = preview.Length > 100 ? preview.Substring(0, 100) + "…" : preview });
                    }
                    catch (Exception ex) when (IsStorageError(ex)) { unreadable++; }
                }
                return result.OrderByDescending(r => r.CreatedAt).ThenBy(r => r.SessionId, StringComparer.Ordinal).ToArray();
            }
        }
        internal void Append(string id, DateTimeOffset createdAt, string documentName, IReadOnlyList<ChatHistoryTurn> turns)
        {
            lock (Sync)
            {
                var record = Read(id) ?? new ChatHistoryRecord { SessionId = id, CreatedAt = createdAt.ToString("O", CultureInfo.InvariantCulture), DocumentName = documentName };
                record.Turns.AddRange(turns);
                Directory.CreateDirectory(_directory);
                var path = FilePath(id); var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    File.WriteAllText(temporary, Json().Serialize(record), new UTF8Encoding(false));
                    if (File.Exists(path)) File.Replace(temporary, path, null);
                    else File.Move(temporary, path);
                }
                finally { try { if (File.Exists(temporary)) File.Delete(temporary); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
            }
        }
        public void Delete(string id) { lock (Sync) File.Delete(FilePath(id)); }
        public void DeleteAll()
        {
            lock (Sync)
            {
                if (!Directory.Exists(_directory)) return;
                foreach (var path in Directory.GetFiles(_directory, "*.json")) File.Delete(path);
            }
        }
        internal static bool IsStorageError(Exception ex) => ex is IOException || ex is UnauthorizedAccessException ||
            ex is System.Security.SecurityException || ex is ArgumentException || ex is InvalidOperationException ||
            ex is FormatException || ex is OverflowException;
    }
}
