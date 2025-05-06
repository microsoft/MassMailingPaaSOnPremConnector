// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
using System;
using System.Diagnostics;
using System.Security;
using System.Text;

namespace MassMailingPaaSOnPremConnector
{
    /*
     * Provides a simple logging mechanism to write to the Windows Event Log.
     * This will try to create a new Event Source (defined by eventSource) if it does not exist, if the operation fails, it will default to "Application".
     * For the creation of the Event Source, the user running the application must have the necessary permissions to create a new Event Source (requires Admin privileges).
     * The logging will be attempted on the custom Event Source (as defined on eventSource), if it fails, writing will be attempted on MassMailingPaaSOnPremConnector, if that fails again it will default to "Application".
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
                    Source = "MassMailingPaaSOnPremConnector";
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

        public void LogDebug(bool isDebugEnabled = true, int eventID = 1, short category = 1)
        {
            if (isDebugEnabled)
            {
                EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Information, eventID, category);
            }
            EventLogMessage.Clear();
        }

        public void LogInformation(int eventID = 1, short category = 1)
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Information, eventID, category);
            EventLogMessage.Clear();
        }

        public void LogWarning(int eventID = 3, short category = 1)
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Warning, eventID, category);
            EventLogMessage.Clear();
        }

        public void LogError(int eventID = 5, short category = 1)
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Error, eventID, category);
            EventLogMessage.Clear();
        }

        public void LogException(int eventID = 9, short category = 1)
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Error, eventID, category);
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

        public void ClearLogEntry()
        {
            EventLogMessage.Clear();
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
