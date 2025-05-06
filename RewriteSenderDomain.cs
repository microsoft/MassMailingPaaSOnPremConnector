// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace MassMailingPaaSOnPremConnector
{

    public class RewriteSenderDomain : RoutingAgentFactory
    {
        /*
         * This class rewrites the sender domain for those senders mathicng the content of the header X-MassMailingPaaSOnPremConnector-SenderRewriteMap.
         * The header value is expected to be a semicolon (;) separated string mapping "<original>=<desired>" with the "=" separating the value fo each entry, such as "contoso.com=tailspintoys.com" or "contoso.com=tailspintoys.com;hotmail.it=hotmail.com".
         * As the X-MassMailingPaaSOnPremConnector-SenderRewriteMap will likely be set via Transport Rule, exclusions can be managed via the transport rules themselves (i.e. insert the X-MassMailingPaaSOnPremConnector-SenderRewriteMap header only if the recipient domain is not xyz).
         * In case multiple agents are active at the same time, be careful about the agent execution oder.
         */
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new MassMailingPaaSOnPremConnector_RewriteSenderDomain();
        }
    }

    public class MassMailingPaaSOnPremConnector_RewriteSenderDomain : RoutingAgent
    {
        static string EventLogName = "RewriteSenderDomain";
        EventLogger EventLog = new EventLogger(EventLogName);

        static readonly string MassMailingPaaSOnPremConnectorTargetName = "X-MassMailingPaaSOnPremConnector-SenderRewriteMap";
        static string MassMailingPaaSOnPremConnectorTargetValue = String.Empty;

        static readonly string RegistryHive = @"Software\TransportAgents\MassMailingPaaSOnPremConnector\RewriteSenderDomain";
        static readonly string RegistryKeyDebugEnabled = "DebugEnabled";
        static bool DebugEnabled = false;

        static readonly string MassMailingPaaSOnPremConnectorName = "X-MassMailingPaaSOnPremConnector-Name";
        static readonly string MassMailingPaaSOnPremConnectorNameValue = "MassMailingPaaSOnPremConnector-RewriteSenderDomain";
        static readonly Dictionary<string, string> MassMailingPaaSOnPremConnectorHeaders = new Dictionary<string, string>
        {
            {MassMailingPaaSOnPremConnectorName, MassMailingPaaSOnPremConnectorNameValue}
        };

        public MassMailingPaaSOnPremConnector_RewriteSenderDomain()
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(RewriteSenderDomain);

            RegistryKey registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
            if (registryPath != null)
            {
                string registryKeyValue = null;
                bool valueConversionResult = false;

                registryKeyValue = registryPath.GetValue(RegistryKeyDebugEnabled, Boolean.FalseString).ToString();
                valueConversionResult = Boolean.TryParse(registryKeyValue, out DebugEnabled);
            }
        }

        void RewriteSenderDomain(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                bool warningOccurred = false;  // controls whether there event log entry is a warning or informational; if anything is out of order log a warning instead of an information log entry. Warnings and Errors are logged regardless of the DebugEnabled setting.
                bool hasProcessedMessage = false; // will be set to true when the message is processed (header present) to only write debug logs when the agent processes the message, and avoiding to log information for messages that has no control header set
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                HeaderList headers = evtMessage.MailItem.Message.MimeDocument.RootPart.Headers;
                Stopwatch stopwatch = Stopwatch.StartNew();
                Dictionary<string, string> SenderRewriteMap = new Dictionary<string, string>();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in MassMailingPaaSOnPremConnector:RewriteSenderDomain", messageId, sender, subject));

                Header MassMailingPaaSOnPremConnectorTarget = headers.FindFirst(MassMailingPaaSOnPremConnectorTargetName);

                if (MassMailingPaaSOnPremConnectorTarget != null && evtMessage.MailItem.Message.IsSystemMessage == false)
                {
                    hasProcessedMessage = true;
                    EventLog.AppendLogEntry(String.Format("Rewriting applicable senders as the messages as the control header {0} is present", MassMailingPaaSOnPremConnectorTargetName));
                    MassMailingPaaSOnPremConnectorTargetValue = MassMailingPaaSOnPremConnectorTarget.Value.Trim();

                    if (!String.IsNullOrEmpty(MassMailingPaaSOnPremConnectorTargetValue))
                    {
                        SenderRewriteMap = MassMailingPaaSOnPremConnectorTargetValue.ToLower()
                                                                                       .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                                       .Select(part => part.Split('='))
                                                                                       .ToDictionary(split => split[0].Trim(), split => split[1].Trim());

                        EventLog.AppendLogEntry(String.Format("Sender domain rewite map start"));
                        foreach (var MapEntry in SenderRewriteMap)
                        {
                            EventLog.AppendLogEntry(String.Format("\t{0} : {1}", MapEntry.Key, MapEntry.Value));
                        }
                        EventLog.AppendLogEntry(String.Format("Sender domain rewite map end"));

                        // Rewriting P1 sender (MAIL FROM:)
                        RoutingAddress P1MsgSender = evtMessage.MailItem.FromAddress;
                        string P1FromLocal = P1MsgSender.LocalPart.ToLower();
                        string P1FromDomain = P1MsgSender.DomainPart.ToLower();
                        string P1FromNewDomain = string.Empty;

                        EventLog.AppendLogEntry(String.Format("Evaluating P1 MAIL FROM: {0}", P1MsgSender.ToString()));

                        if (SenderRewriteMap.ContainsKey(P1FromDomain))
                        {
                            P1FromNewDomain = SenderRewriteMap[P1FromDomain];
                            evtMessage.MailItem.FromAddress = new RoutingAddress(P1FromLocal, P1FromNewDomain);
                            EventLog.AppendLogEntry(String.Format("MAIL FROM {0}@{1} rewritten to {2}@{3}", P1FromLocal, P1FromDomain, P1FromLocal, P1FromNewDomain));
                        }

                        // Rewriting P2 sender (FROM:)
                        string P2MsgFrom = evtMessage.MailItem.Message.From.SmtpAddress;
                        int P2FromAtIndex = P2MsgFrom.IndexOf("@");
                        int P2FromRecLength = P2MsgFrom.Length;
                        string P2FromLocal = P2MsgFrom.Substring(0, P2FromAtIndex);
                        string P2FromDomain = P2MsgFrom.Substring(P2FromAtIndex + 1, P2FromRecLength - P2FromAtIndex - 1);
                        string P2FromNewDomain = string.Empty;

                        EventLog.AppendLogEntry(String.Format("Evaluating P2 FROM: {0}", P2MsgFrom));

                        if (SenderRewriteMap.ContainsKey(P2FromDomain))
                        {
                            P2FromNewDomain = SenderRewriteMap[P2FromDomain];
                            evtMessage.MailItem.Message.From.SmtpAddress = P2FromLocal + "@" + P2FromNewDomain;
                            EventLog.AppendLogEntry(String.Format("P2 FROM {0}@{1} rewritten to {2}@{3}", P2FromLocal, P2FromDomain, P2FromLocal, P2FromNewDomain));
                        }

                        // Rewriting P2 sender (SENDER:)
                        string P2MsgSender = evtMessage.MailItem.Message.Sender.SmtpAddress;
                        int P2SenderAtIndex = P2MsgSender.IndexOf("@");
                        int P2SenderRecLength = P2MsgSender.Length;
                        string P2SenderLocal = P2MsgSender.Substring(0, P2SenderAtIndex);
                        string P2SenderDomain = P2MsgSender.Substring(P2SenderAtIndex + 1, P2SenderRecLength - P2SenderAtIndex - 1);
                        string P2SenderNewDomain = string.Empty;

                        EventLog.AppendLogEntry(String.Format("Evaluating P2 SENDER: {0}", P2MsgFrom));

                        if (SenderRewriteMap.ContainsKey(P2SenderDomain))
                        {
                            P2SenderNewDomain = SenderRewriteMap[P2SenderDomain];
                            evtMessage.MailItem.Message.Sender.SmtpAddress = P2SenderLocal + "@" + P2SenderNewDomain;
                            EventLog.AppendLogEntry(String.Format("P2 SENDER {0}@{1} rewritten to {2}@{3}", P2SenderLocal, P2SenderDomain, P2SenderLocal, P2SenderNewDomain));
                        }

                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("There was a problem processing the {0} header value", MassMailingPaaSOnPremConnectorTargetName));
                        EventLog.AppendLogEntry(String.Format("There value retrieved is: {0}", MassMailingPaaSOnPremConnectorTargetValue));
                        warningOccurred = true;
                    }

                    foreach (var newHeader in MassMailingPaaSOnPremConnectorHeaders)
                    {
                        Header HeaderExists = headers.FindFirst(newHeader.Key);
                        if (HeaderExists == null || HeaderExists.Value != newHeader.Value)
                        {
                            evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader(newHeader.Key, newHeader.Value), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                            EventLog.AppendLogEntry(String.Format("ADDED header {0}: {1}", newHeader.Key, String.IsNullOrEmpty(newHeader.Value) ? String.Empty : newHeader.Value));
                        }
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
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is not set", MassMailingPaaSOnPremConnectorTargetName));
                    }
                }

                EventLog.AppendLogEntry(String.Format("MassMailingPaaSOnPremConnector:RewriteSenderDomain took {0} ms to execute", stopwatch.ElapsedMilliseconds));

                if (warningOccurred)
                {
                    EventLog.LogWarning();
                }
                else
                {
                    if (hasProcessedMessage)
                    {
                        EventLog.LogDebug(DebugEnabled);
                    }
                    else
                    {
                        EventLog.ClearLogEntry();
                    }
                }

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in MassMailingPaaSOnPremConnector:RewriteSenderDomain");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;


        }
    }
}