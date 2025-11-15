// File: Logger.cs
// Project: H80 Observatory Control (Class Library .NET Framework 4.7.2)
// Purpose: Shared singleton logger that emits to ACP.Console and to C:\Wise\Logs\<date>\polcam.log

using ACP;
using System;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text;

namespace H80
{
    public sealed class Logger
    {
        // Public singleton instance for global use
        public static readonly Logger logger = new Logger();
        private readonly object _sync = new object();
        private readonly Util util = (Util) Marshal.GetActiveObject("ACP.Util");
        private const string LOG_DIR = @"C:\\Wise\\Logs"; // ensure writeable

        public Logger()
        {
            Directory.CreateDirectory(LOG_DIR);
        }

        // Public methods
        public void Debug(string message) => Log("[D]", message, null);
        public void info(string message) => Log("[I]", message, null);
        public void Warning(string message) => Log("[W]", message, null);
        public void Error(string message) => Log("[E]", message, null);

        private void Log(string prefix, string message, Exception ex)
        {
            try
            {
                string line = ComposeLine(prefix, message, ex);
                TryConsolePrint(line);

                string dateFolder = DateTime.Now.ToString("yyyy-MM-dd");
                string dir = Path.Combine("C:\\Wise\\Logs", dateFolder);
                string path = Path.Combine(dir, "h80.log");
                string timeStamp = DateTime.UtcNow.ToString("HH:mm:ss'.'fff'Z'") + " ";
                Directory.CreateDirectory(dir);

                lock (_sync)
                {
                    File.AppendAllText(path, timeStamp + line + Environment.NewLine, Encoding.UTF8);
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
            sb.Append(prefix).Append(' ').Append(message ?? string.Empty);
            if (ex != null)
                sb.Append(" | EX: ").Append(ex.GetType().Name).Append(": ").Append(ex.Message);
            return sb.ToString();
        }

        private void TryConsolePrint(string line)
        {
            if (line == null) return;

            try
            {
                util.Console.PrintLine(line.EndsWith("\r\n") ? line : line +"\r\n");
            }
            catch { /* ignore console issues */ }
        }

        private static void Dispose() { }
    }
}
