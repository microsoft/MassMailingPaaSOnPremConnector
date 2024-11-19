using System;
using System.Diagnostics;
using System.Security;
using System.Text;

namespace ACSOnPremConnector
{
    /*
     * Provides a simple logging mechanism to write to the Windows Event Log.
     * This will try to create a new Event Source (defined by eventSource) if it does not exist, if the operation fails, it will default to "Application".
     * For the creation of the Event Source, the user running the application must have the necessary permissions to create a new Event Source (requires Admin privileges).
     * The logging will be attempted on the custom Event Source (as defined on eventSource), if it fails, writing will be attempted on ACSOnPremConnector, if that fails again it will default to "Application".
     */
    internal class EventLogger : IDisposable
    {
        private string Source = String.Empty;
        private StringBuilder EventLogMessage = null;

        public EventLogger(string eventSource = "Application")
        {
            EventLogMessage = new StringBuilder();

            try
            {
                Source = eventSource;
                if (!EventLog.SourceExists(Source))
                {
                    EventLog.CreateEventSource(Source, "Application");
                }
            }
            catch (SecurityException)
            {
                try
                {
                    Source = "ACSOnPremConnector";
                    if (!!EventLog.SourceExists(Source))
                    {
                        EventLog.CreateEventSource(Source, "Application");
                    }
                }
                catch (SecurityException)
                {
                    Source = "Application";
                    if (!EventLog.SourceExists(Source))
                    {
                        EventLog.CreateEventSource(Source, "Application");
                    }
                }
            }
        }

        public void LogDebug(bool isDebugEnabled = true)
        {
            if (isDebugEnabled)
            {
                EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Information);
            }
            EventLogMessage.Clear();
        }

        public void LogInformation()
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Information);
            EventLogMessage.Clear();
        }

        public void LogWarning()
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Warning);
            EventLogMessage.Clear();
        }

        public void LogError()
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Error);
            EventLogMessage.Clear();
        }

        public void LogException()
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Error);
            EventLogMessage.Clear();
        }

        public void AppendLogEntry(string message)
        {
            EventLogMessage.AppendLine(message);
        }

        public void AppendLogEntry(Exception ex)
        {
            EventLogMessage.AppendLine("--------------------------------------------------------------------------------");
            EventLogMessage.AppendLine(String.Format("EXCEPTION MESSAGE: {0}", ex.Message));
            EventLogMessage.AppendLine(String.Format("EXCEPTION HRESULT: {0}", ex.HResult));
            EventLogMessage.AppendLine(String.Format("EXCEPTION SOURCE: {0}", ex.Source));
            EventLogMessage.AppendLine(String.Format("EXCEPTION INNER EXCEPTION: {0}", ex.InnerException));
            EventLogMessage.AppendLine(String.Format("EXCEPTION STRACK: {0}", ex.StackTrace));
            EventLogMessage.AppendLine("--------------------------------------------------------------------------------");
            EventLogMessage.AppendLine(ex.ToString());
            EventLogMessage.AppendLine("--------------------------------------------------------------------------------");
        }

        public void AppendLogEntry(object obj)
        {
            EventLogMessage.AppendLine(obj.ToString());
        }

        ~EventLogger()
        {
            WriteEventLogOnExit();
        }

        void IDisposable.Dispose()
        {
            WriteEventLogOnExit();
        }

        private void WriteEventLogOnExit()
        {
            if (!String.IsNullOrEmpty(EventLogMessage.ToString()))
            {
                EventLogMessage.AppendLine("Writing Event on Agent exit");
                EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Information);
                EventLogMessage.Clear();
            }
            EventLogMessage = null;
        }

    }
}
