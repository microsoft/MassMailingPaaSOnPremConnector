using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace ACSOnPremConnector
{
    /*
     * This address P1 Sender (Mail From:) and P2 Sender (From:) mismatch. Two functionalites are provided:
     *  - Overriding the P1 with the P2 value, or vice-versa by using the X-ACSOnPremConnector-P1P2MismatchAction header. Valid values are "UseP1" (overwrite P2), "UseP2" (Overwrite P1), "None" (do nothing).
     *  - Forcing P1 to a custom value, by using the X-ACSOnPremConnector-ForceP1 header. In this case the value of the header must be a valid SMTP address.
     * In case both headers are set, both operations will be executed, with the P1 override being executed last, ensuring that the provided value is used as P1 Sender (Mail From:).
     */
    public class ProcessSenderMismatch : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new ACSOnPremConnector_ProcessSenderMismatch();
        }
    }

    public class ACSOnPremConnector_ProcessSenderMismatch : RoutingAgent
    {
        static string EventLogName = "ProcessSenderMismatch";
        EventLogger EventLog = new EventLogger(EventLogName);

        static readonly string ACSOnPremConnectorP1P2MismatchActionName = "X-ACSOnPremConnector-P1P2MismatchAction";
        static string ACSOnPremConnectorP1P2MismatchActionValue = String.Empty;
        static readonly string ACSOnPremConnectorForceP1Name = "X-ACSOnPremConnector-ForceP1";
        static string ACSOnPremConnectorForceP1Value = String.Empty;

        static readonly string RegistryHive = @"Software\TransportAgents\ACSOnPremConnector\ProcessSenderMismatch";
        static readonly string RegistryKeyDebugEnabled = "DebugEnabled";
        static bool DebugEnabled = false;

        static readonly string ACSOnPremConnectorName = "X-ACSOnPremConnector-Name";
        static readonly string ACSOnPremConnectorNameValue = "ACSOnPremConnector-ProcessSenderMismatch";
        static readonly Dictionary<string, string> ACSOnPremConnectorHeaders = new Dictionary<string, string>
        {
            {ACSOnPremConnectorName, ACSOnPremConnectorNameValue},
            {"X-ACSOnPremConnector-Creator", "Tommaso Toniolo"},
            {"X-ACSOnPremConnector-Contact", "https://aka.ms/totoni"}
        };

        public ACSOnPremConnector_ProcessSenderMismatch()
        {
            base.OnRoutedMessage += new RoutedMessageEventHandler(ProcessSenderMismatch);

            RegistryKey registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
            if (registryPath != null)
            {
                string registryKeyValue = null;
                bool valueConversionResult = false;

                registryKeyValue = registryPath.GetValue(RegistryKeyDebugEnabled, Boolean.FalseString).ToString();
                valueConversionResult = Boolean.TryParse(registryKeyValue, out DebugEnabled);
            }
        }

        void ProcessSenderMismatch(RoutedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                bool warningOccurred = false;
                bool messageProcessed = false;
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                string P1Sender = evtMessage.MailItem.FromAddress.ToString();
                string P2Sender = evtMessage.MailItem.Message.Sender.SmtpAddress;
                HeaderList headers = evtMessage.MailItem.Message.MimeDocument.RootPart.Headers;
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in ACSOnPremConnector:ProcessSenderMismatch", messageId, sender, subject));

                Header ACSOnPremConnectorP1P2MismatchAction = headers.FindFirst(ACSOnPremConnectorP1P2MismatchActionName);

                if (ACSOnPremConnectorP1P2MismatchAction != null && evtMessage.MailItem.Message.IsSystemMessage == false)
                {
                    EventLog.AppendLogEntry(String.Format("Evaluating P1/P2 Sender Mismatch as the control header {0} is present", ACSOnPremConnectorP1P2MismatchActionName));
                    ACSOnPremConnectorP1P2MismatchActionValue = ACSOnPremConnectorP1P2MismatchAction.Value.Trim().ToUpper();

                    if (!String.IsNullOrEmpty(ACSOnPremConnectorP1P2MismatchActionValue) &&
                        (String.Equals(ACSOnPremConnectorP1P2MismatchActionValue, "UseP1", StringComparison.OrdinalIgnoreCase) ||
                         String.Equals(ACSOnPremConnectorP1P2MismatchActionValue, "UseP2", StringComparison.OrdinalIgnoreCase) ||
                         String.Equals(ACSOnPremConnectorP1P2MismatchActionValue, "None", StringComparison.OrdinalIgnoreCase))
                    )
                    {
                        EventLog.AppendLogEntry(String.Format("P1/P2 Mismatch Action is valid as the header {0} is set to {1}", ACSOnPremConnectorP1P2MismatchActionName, ACSOnPremConnectorP1P2MismatchActionValue));

                        EventLog.AppendLogEntry(String.Format("P1 Sender is set to: {0}", P1Sender));
                        EventLog.AppendLogEntry(String.Format("P2 Sender is set to: {0}", P2Sender));

                        switch (ACSOnPremConnectorP1P2MismatchActionValue)
                        {
                            case "USEP1":
                                evtMessage.MailItem.Message.Sender.SmtpAddress = P1Sender;
                                evtMessage.MailItem.Message.From.SmtpAddress = P1Sender;
                                EventLog.AppendLogEntry(String.Format("P2 Sender has been set to: {0}", P1Sender));
                                messageProcessed = true;
                                break;
                            case "USEP2":
                                evtMessage.MailItem.FromAddress = new RoutingAddress(P2Sender);
                                EventLog.AppendLogEntry(String.Format("P1 Sender has been set to: {0}", P2Sender));
                                messageProcessed = true;
                                break;
                            case "NONE":
                                EventLog.AppendLogEntry(String.Format("No action has been taken as the header is set to {0}", ACSOnPremConnectorP1P2MismatchActionValue));
                                messageProcessed = true;
                                break;
                            default:
                                EventLog.AppendLogEntry(String.Format("P1 and P2 have been left unmodified"));
                                break;
                        }
                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("There was a problem processing the {0} header value", ACSOnPremConnectorP1P2MismatchActionName));
                        EventLog.AppendLogEntry(String.Format("There value retrieved is: {0}; Valid (case insensitive) values are UseP1, UseP2, None", ACSOnPremConnectorP1P2MismatchActionValue));
                        warningOccurred = true;
                    }

                }
                else
                {
                    if (evtMessage.MailItem.Message.IsSystemMessage == true)
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as IsSystemMessage"));
                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is not set", ACSOnPremConnectorP1P2MismatchActionName));
                    }
                }

                Header ACSOnPremConnectorForceP1 = headers.FindFirst(ACSOnPremConnectorForceP1Name);

                if (ACSOnPremConnectorForceP1 != null && evtMessage.MailItem.Message.IsSystemMessage == false)
                {
                    EventLog.AppendLogEntry(String.Format("Overriding P1 Sender as the control header {0} is present", ACSOnPremConnectorForceP1Name));
                    ACSOnPremConnectorForceP1Value = ACSOnPremConnectorForceP1.Value.Trim().ToUpper();

                    RoutingAddress newP1 = new RoutingAddress(ACSOnPremConnectorForceP1Value);
                    EventLog.AppendLogEntry(String.Format("The new P1 Sender will be forced is {0}", newP1.ToString()));

                    EventLog.AppendLogEntry(String.Format("P1 Sender is currently set to: {0}", P1Sender));
                    EventLog.AppendLogEntry(String.Format("P2 Sender is currently set to: {0}", P2Sender));

                    if (newP1.IsValid == false)
                    {
                        EventLog.AppendLogEntry(String.Format("The provided P1 Sender {0} IS INVALID", newP1.ToString()));
                        warningOccurred = true;
                    }
                    else
                    {
                        evtMessage.MailItem.FromAddress = newP1;
                        messageProcessed = true;
                        EventLog.AppendLogEntry(String.Format("Forced P1 Sender to {0}", evtMessage.MailItem.FromAddress.ToString()));
                    }
                }
                else
                {
                    if (evtMessage.MailItem.Message.IsSystemMessage == true)
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as IsSystemMessage"));
                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is not set", ACSOnPremConnectorForceP1Name));
                    }
                }

                if(messageProcessed)
                {
                    foreach (var newHeader in ACSOnPremConnectorHeaders)
                    {
                        Header HeaderExists = headers.FindFirst(newHeader.Key);
                        if (HeaderExists == null || HeaderExists.Value != newHeader.Value)
                        {
                            evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader(newHeader.Key, newHeader.Value), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                            EventLog.AppendLogEntry(String.Format("ADDED header {0}: {1}", newHeader.Key, String.IsNullOrEmpty(newHeader.Value) ? String.Empty : newHeader.Value));
                        }
                    }
                }

                EventLog.AppendLogEntry(String.Format("ACSOnPremConnector:ProcessSenderMismatch took {0} ms to execute", stopwatch.ElapsedMilliseconds));

                if (warningOccurred)
                {
                    EventLog.LogWarning();
                }
                else
                {
                    EventLog.LogDebug(DebugEnabled);
                }

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in ACSOnPremConnector:ProcessSenderMismatch");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }
    }
}
