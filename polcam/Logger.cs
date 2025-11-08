// File: Logger.cs
// Project: polcam Observatory Control (Class Library .NET Framework 4.7.2)
// Purpose: Shared singleton logger that emits to ACP.Console and to C:\Wise\Logs\<date>\polcam.log

using System;
using System.IO;
using System.Linq;
using System.Text;

namespace polcam
{
    public sealed class Logger
    {
        // Public singleton instance for global use
        public static readonly Logger logger = new Logger();
        private readonly object _sync = new object();

        public Logger() { }

        // Public methods
        public void Debug(string message) => Log("[D]", message, null);
        public void info(string message) => Log("[I]", message, null);
        public void Warning(string message) => Log("[W]", message, null);
        public void Error(string message) => Log("[E]", message, null);

        //public void Debug(string message, Exception ex) => Log("[D]", message, ex);
        //public void Info(string message, Exception ex) => Log("[I]", message, ex);
        //public void Warning(string message, Exception ex) => Log("[W]", message, ex);
        //public void Error(string message, Exception ex) => Log("[E]", message, ex);

        private void Log(string prefix, string message, Exception ex)
        {
            try
            {
                string line = ComposeLine(prefix, message, ex);
                TryConsolePrint(line);

                string dateFolder = DateTime.Now.ToString("yyyy-MM-dd");
                string dir = Path.Combine("C:\\Wise\\Logs", dateFolder);
                string path = Path.Combine(dir, "polcam.log");
                Directory.CreateDirectory(dir);

                lock (_sync)
                {
                    File.AppendAllText(path, line + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch
            {
                // Never throw from logger
            }
        }

        private static string ComposeLine(string prefix, string message, Exception ex)
        {
            var sb = new StringBuilder();
            sb.Append(DateTime.UtcNow.ToString("HH:mm:ss'.'fff'Z'")).Append(' ').Append(prefix).Append(' ').Append(message ?? string.Empty);
            if (ex != null)
                sb.Append(" | EX: ").Append(ex.GetType().Name).Append(": ").Append(ex.Message);
            return sb.ToString();
        }

        private static void TryConsolePrint(string line)
        {
            try
            {
                var t = Type.GetTypeFromProgID("ACP.Console");
                if (t == null) return;
                dynamic console = Activator.CreateInstance(t);
                console.Print(line.EndsWith("\r\n") ? line : line + "\r\n");
            }
            catch { /* ignore console issues */ }
        }

        private static void Dispose() { }
    }
}
