using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACSOnPremConnector
{
    public class MessageLevelInspector: RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new ACSOnPremConnector_MessageLevelInspector();
        }
    }
    public class ACSOnPremConnector_MessageLevelInspector : RoutingAgent
    {
        EventLogger EventLog = new EventLogger("MessageLevelInspector");

        static readonly string RegistryHive = @"Software\TransportAgents\ACSOnPremConnector\MessageLevelInspector";
        static readonly string RegistryKeyDebugEnabled = "DebugEnabled";
        static bool DebugEnabled = false;

        public ACSOnPremConnector_MessageLevelInspector()
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

        void PrintMessagePropertiesToLog (string phase, QueuedMessageEventArgs evtMessage)
        {
            bool warningOccurred = false;
            Stopwatch stopwatch = Stopwatch.StartNew();

            EventLog.AppendLogEntry(String.Format("Processing message in ACSOnPremConnector:MessageLevelInspector:{0}", phase));

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

            if (evtMessage.MailItem.Message.Sender.SmtpAddress.ToString().ToLower().Trim() !=  evtMessage.MailItem.Message.From.SmtpAddress.ToString().ToLower().Trim())
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

            EventLog.AppendLogEntry(String.Format("ACSOnPremConnector:MessageLevelInspector:{0} took {1} ms to execute", phase, stopwatch.ElapsedMilliseconds));

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
