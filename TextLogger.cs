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
