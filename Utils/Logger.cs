using System;
using System.IO;

namespace QRCodeRevitAddin.Utils
{
    public static class Logger
    {
        private static readonly string LogDirectory = Path.Combine(Path.GetTempPath(), "DEAXO_QRCodeAddin_Logs");
        private static readonly object LogLock = new object();

        static Logger()
        {
            try
            {
                if (!Directory.Exists(LogDirectory))
                {
                    Directory.CreateDirectory(LogDirectory);
                }
            }
            catch
            {
            }
        }

        public static void LogInfo(string message)
        {
            Log("INFO", message);
        }

        public static void LogWarning(string message)
        {
            Log("WARNING", message);
        }

        public static void LogError(string message, Exception ex = null)
        {
            string fullMessage = message;
            if (ex != null)
            {
                fullMessage += $"\nException: {ex.GetType().Name}\nMessage: {ex.Message}\nStackTrace: {ex.StackTrace}";
            }
            Log("ERROR", fullMessage);
        }

        private static void Log(string level, string message)
        {
            try
            {
                lock (LogLock)
                {
                    string logFileName = $"QRAddin_{DateTime.Now:yyyyMMdd}.log";
                    string logFilePath = Path.Combine(LogDirectory, logFileName);

                    string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}\n";

                    File.AppendAllText(logFilePath, logEntry);
                }
            }
            catch
            {
            }
        }

        public static string GetLogDirectory()
        {
            return LogDirectory;
        }
    }
}