using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Win32;
using System;
using System.Diagnostics;

namespace MassMailingPaaSOnPremConnector
{
    /*
     * This class provides debug functionality to log the properties of the message during processing.
     * This will print on the event log the key properites (P1, P2, Headers) on the Event Log.
     * As all the agents in this module executes either at OnResolvedMessage (2nd step in categorization) or OnRoutedMessage (3rd step in categorization), 
     * this agent will print the properties of the message when it's received OnSubmittedMessage (1st step in categorization) and post processing OnCategorizedMessage (4th step in categorization).
     * By doing so the event log will contain a snapshot of the message before processing, one event during processing with the changes performed, and one last event after processing.
     * This shall allow to investigate unexpected behaviour at message level, without needing to rely on Pipeline Tracing or SMTP Protocol Logs.
     * For this agent, opposite to all others, DebugEnabled is assumed "true" unless disabled via registry key. It can be disabled bia Disable-TransportAgent if necessary.
     * This agent is not intended to be left on at all time, but rather to be enabled for troubleshooting and then disabled.
     */
    public class MessageLevelInspector : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new MassMailingPaaSOnPremConnector_MessageLevelInspector();
        }
    }
    public class MassMailingPaaSOnPremConnector_MessageLevelInspector : RoutingAgent
    {
        static string EventLogName = "MessageLevelInspector";
        EventLogger EventLog = new EventLogger(EventLogName);

        static readonly string RegistryHive = @"Software\TransportAgents\MassMailingPaaSOnPremConnector\MessageLevelInspector";
        static readonly string RegistryKeyDebugEnabled = "DebugEnabled";
        static bool DebugEnabled = true;

        public MassMailingPaaSOnPremConnector_MessageLevelInspector()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(MessageLevelInspectorPreProcess);
            base.OnCategorizedMessage += new CategorizedMessageEventHandler(MessageLevelInspectorPostProcess);

            RegistryKey registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
            if (registryPath != null)
            {
                string registryKeyValue = null;
                bool valueConversionResult = false;

                registryKeyValue = registryPath.GetValue(RegistryKeyDebugEnabled, Boolean.FalseString).ToString();
                valueConversionResult = Boolean.TryParse(registryKeyValue, out DebugEnabled);
            }
        }

        void MessageLevelInspectorPreProcess(SubmittedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            PrintMessagePropertiesToLog("OnSubmittedMessage", evtMessage);
            return;
        }

        void MessageLevelInspectorPostProcess(CategorizedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            PrintMessagePropertiesToLog("OnCategorizedMessage", evtMessage);
            return;
        }

        void PrintMessagePropertiesToLog(string phase, QueuedMessageEventArgs evtMessage)
        {
            bool warningOccurred = false;
            Stopwatch stopwatch = Stopwatch.StartNew();

            EventLog.AppendLogEntry(String.Format("Processing message in MassMailingPaaSOnPremConnector:MessageLevelInspector:{0}", phase));

            EventLog.AppendLogEntry("==================== ENVELOPE - P1 ====================");
            EventLog.AppendLogEntry(String.Format("EnvelopeId: {0}", evtMessage.MailItem.EnvelopeId));
            EventLog.AppendLogEntry(String.Format("P1 Sender: {0}", evtMessage.MailItem.FromAddress.ToString().ToLower().Trim()));

            foreach (var recipient in evtMessage.MailItem.Recipients)
                EventLog.AppendLogEntry(String.Format("P1 Recipient: {0}", recipient.Address.ToString().ToLower()));

            EventLog.AppendLogEntry(String.Format("IsSystemMessage: {0}", evtMessage.MailItem.Message.IsSystemMessage));
            EventLog.AppendLogEntry(String.Format("IsInterpersonalMessage: {0}", evtMessage.MailItem.Message.IsInterpersonalMessage));
            EventLog.AppendLogEntry(String.Format("OriginatingDomain: {0}", evtMessage.MailItem.OriginatingDomain));
            EventLog.AppendLogEntry(String.Format("OriginatorOrganization: {0}", evtMessage.MailItem.OriginatorOrganization));
            EventLog.AppendLogEntry(String.Format("OriginalAuthenticator: {0}", evtMessage.MailItem.OriginalAuthenticator));

            foreach (var item in evtMessage.MailItem.Properties)
                EventLog.AppendLogEntry(String.Format("Property - {0}: {1}", item.Key.ToString(), item.Value.ToString()));

            EventLog.AppendLogEntry("==================== HEADERS ====================");
            foreach (var header in evtMessage.MailItem.Message.MimeDocument.RootPart.Headers)
                EventLog.AppendLogEntry(String.Format("{0}: {1}", header.Name, String.IsNullOrEmpty(header.Value) ? String.Empty : header.Value));

            EventLog.AppendLogEntry("==================== MESSAGE - P2 ====================");
            EventLog.AppendLogEntry(String.Format("MessageId: {0}", evtMessage.MailItem.Message.MessageId.ToString()));
            EventLog.AppendLogEntry(String.Format("Subject: {0}", evtMessage.MailItem.Message.Subject.Trim()));
            EventLog.AppendLogEntry(String.Format("P2 Sender: {0}", evtMessage.MailItem.Message.Sender.SmtpAddress.ToString().ToLower().Trim()));
            EventLog.AppendLogEntry(String.Format("P2 From: {0}", evtMessage.MailItem.Message.From.SmtpAddress.ToString().ToLower().Trim()));
            EventLog.AppendLogEntry(String.Format("MapiMessageClass: {0}", evtMessage.MailItem.Message.MapiMessageClass.ToString().Trim()));

            foreach (var recipient in evtMessage.MailItem.Message.To)
                EventLog.AppendLogEntry(String.Format("P2 To: {0}", recipient.SmtpAddress.ToString().ToLower().Trim()));

            foreach (var recipient in evtMessage.MailItem.Message.Cc)
                EventLog.AppendLogEntry(String.Format("P2 Cc: {0}", recipient.SmtpAddress.ToString().ToLower().Trim()));

            foreach (var recipient in evtMessage.MailItem.Message.Bcc)
                EventLog.AppendLogEntry(String.Format("P2 Bcc: {0}", recipient.SmtpAddress.ToString().ToLower().Trim()));

            foreach (var recipient in evtMessage.MailItem.Message.ReplyTo)
                EventLog.AppendLogEntry(String.Format("P2 ReplyTo: {0}", recipient.SmtpAddress.ToString().ToLower().Trim()));

            if ((evtMessage.MailItem.FromAddress.ToString().ToLower().Trim() != evtMessage.MailItem.Message.Sender.SmtpAddress.ToString().ToLower().Trim()) ||
                (evtMessage.MailItem.FromAddress.ToString().ToLower().Trim() != evtMessage.MailItem.Message.From.SmtpAddress.ToString().ToLower().Trim()))
            {
                EventLog.AppendLogEntry("==================== IMPORTANT ====================");
                EventLog.AppendLogEntry("Note that the P1 Sender and the P2 Sender mismatch. This can be source of problems");
                warningOccurred = true;
            }

            if (evtMessage.MailItem.Message.Sender.SmtpAddress.ToString().ToLower().Trim() != evtMessage.MailItem.Message.From.SmtpAddress.ToString().ToLower().Trim())
            {
                EventLog.AppendLogEntry("==================== IMPORTANT ====================");
                EventLog.AppendLogEntry("Note that the P2 Sender and the P2 From mismatch. This can be source of problems");
                warningOccurred = true;
            }

            if (evtMessage.MailItem.Message.Sender.SmtpAddress.ToString().ToLower().Trim().Contains(",") ||
                evtMessage.MailItem.Message.From.SmtpAddress.ToString().ToLower().Trim().Contains(","))
            {
                EventLog.AppendLogEntry("==================== IMPORTANT ====================");
                EventLog.AppendLogEntry("Note that the P2 Sender or From contains a comma ','. This might mean there are multiple From address set and can be source of problems");
                warningOccurred = true;
            }

            EventLog.AppendLogEntry(String.Format("MassMailingPaaSOnPremConnector:MessageLevelInspector:{0} took {1} ms to execute", phase, stopwatch.ElapsedMilliseconds));

            if (warningOccurred)
            {
                EventLog.LogWarning();
            }
            else
            {
                EventLog.LogDebug(DebugEnabled);
            }

            return;
        }
    }

}
