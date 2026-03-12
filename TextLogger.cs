// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
using System;
using System.IO;

namespace MassMailingPaaSOnPremConnector
{
    /*
     * Provides a simple logging mechanism to write to a text file. The text fille will be stored on logLocation and named logName.
     * It is imperative that this log file is excluded from any monitoring or anti-virus software, as it will be written to constantly.
     */
    internal class TextLogger : IDisposable
    {
        private string _logPath = string.Empty;
        private StreamWriter _logStream = null;

        public TextLogger(string logLocation, string logName)
        {
            string logPath = Path.Combine(logLocation, logName);

            if (_logPath != logPath)
            {
                CloseStream();
            }

            try
            {
                _logPath = logPath;
                _logStream = new StreamWriter(_logPath, true);
            }
            catch (Exception)
            {
                // Swallow exceptions and continue without logging if the log stream cannot be created or accessed.
                // This is to avoid any disruption to the main functionality of the agent in case of issues with the log file (e.g. permission issues, file lock by other process, etc.).
            }
        }

        public TextLogger(string logPath)
        {
            if (_logPath != logPath)
            {
                CloseStream();
            }

            try
            {
                _logPath = logPath;
                _logStream = new StreamWriter(_logPath, true);
            }
            catch (Exception)
            {
                // Swallow exceptions and continue without logging if the log stream cannot be created or accessed.
                // This is to avoid any disruption to the main functionality of the agent in case of issues with the log file (e.g. permission issues, file lock by other process, etc.).
            }
        }

        ~TextLogger()
        {
            CloseStream();
        }

        void IDisposable.Dispose()
        {
            CloseStream();
        }
        private void CloseStream()
        {
            if (_logStream != null)
            {
                _logStream.Flush();
                _logStream.Dispose();
            }
        }

        public void WriteToText(string message)
        {
            _logStream.WriteLine(String.Format("{0:yyyy-MM-dd HH:mm:ss} | {1}", DateTime.Now, message));
            _logStream.Flush();
        }

    }
}
