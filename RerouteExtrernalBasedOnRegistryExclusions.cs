// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace MassMailingPaaSOnPremConnector
{
    /*
     * This class reroutes messages to external recipients when the header X-MassMailingPaaSOnPremConnector-Target set to a domain.
     * The domain value doesn't need to be routable, but has to be avalid domain (i.e. something.value.tld).
     * This agent will reroute all the messages via the custom routing domain, only if the recipient email address or the recipient domain are not listed in the ExemptedRecipientDomains and ExemptedRecipientAddresses keys in local machine registry
     * As the X-MassMailingPaaSOnPremConnector-Target will likely be set via Transport Rule, further exclusions can still be managed via the transport rules themselves if necessary.
     * In case multiple agents are active at the same time, only the first one will trigger as the other will detect the presence of the X-MassMailingPaaSOnPremConnector-Target header which is used for loop protection. This is by design to protect mail loops.
     */
    public class RerouteExtrernalBasedOnRegistryExclusions : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new MassMailingPaaSOnPremConnector_RerouteExtrernalBasedOnRegistryExclusions();
        }
    }

    public class MassMailingPaaSOnPremConnector_RerouteExtrernalBasedOnRegistryExclusions : RoutingAgent
    {
        static string EventLogName = "RerouteExtrernalBasedOnRegistryExclusions";
        EventLogger EventLog = new EventLogger(EventLogName);

        static readonly string MassMailingPaaSOnPremConnectorTargetName = "X-MassMailingPaaSOnPremConnector-Target";
        static string MassMailingPaaSOnPremConnectorTargetValue = String.Empty;

        static readonly string RegistryHive = @"Software\TransportAgents\MassMailingPaaSOnPremConnector\RerouteExtrernalBasedOnRegistryExclusions";
        static readonly string RegistryKeyDebugEnabled = "DebugEnabled";
        static bool DebugEnabled = false;
        static readonly string RegistryKeyExemptedRecipientDomains = "ExemptedRecipientDomains";
        static List<string> ExemptedRecipientDomains = new List<string>();
        static readonly string RegistryKeyExemptedRecipientAddresses = "ExemptedRecipientAddresses";
        static List<string> ExemptedRecipientAddresses = new List<string>();

        static readonly string MassMailingPaaSOnPremConnectorName = "X-MassMailingPaaSOnPremConnector-Name";
        static readonly string MassMailingPaaSOnPremConnectorNameValue = "MassMailingPaaSOnPremConnector-RerouteExtrernalBasedOnRegistryExclusions";
        static readonly Dictionary<string, string> MassMailingPaaSOnPremConnectorHeaders = new Dictionary<string, string>
        {
            {MassMailingPaaSOnPremConnectorName, MassMailingPaaSOnPremConnectorNameValue}
        };

        public MassMailingPaaSOnPremConnector_RerouteExtrernalBasedOnRegistryExclusions()
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(RerouteExtrernalBasedOnRegistryExclusions);

            RegistryKey registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
            if (registryPath != null)
            {
                string registryKeyValue = null;
                bool valueConversionResult = false;

                registryKeyValue = registryPath.GetValue(RegistryKeyDebugEnabled, Boolean.FalseString).ToString();
                valueConversionResult = Boolean.TryParse(registryKeyValue, out DebugEnabled);

                string[] retrievedDomains = (string[])registryPath.GetValue(RegistryKeyExemptedRecipientDomains);
                if (retrievedDomains != null && retrievedDomains.Length > 0)
                {
                    foreach (string domain in retrievedDomains)
                        if (!ExemptedRecipientDomains.Contains(domain.ToLower()))
                            ExemptedRecipientDomains.Add(domain.ToLower());
                    ExemptedRecipientDomains.Sort();
                }

                string[] retrievedRecipients = (string[])registryPath.GetValue(RegistryKeyExemptedRecipientAddresses);
                if (retrievedRecipients != null && retrievedRecipients.Length > 0)
                {
                    foreach (string recipient in retrievedRecipients)
                        if (!ExemptedRecipientAddresses.Contains(recipient.ToLower()))
                            ExemptedRecipientAddresses.Add(recipient.ToLower());
                    ExemptedRecipientAddresses.Sort();
                }
            }

        }

        void RerouteExtrernalBasedOnRegistryExclusions(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
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

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in MassMailingPaaSOnPremConnector:RerouteExtrernalBasedOnRegistryExclusions", messageId, sender, subject));

                Header MassMailingPaaSOnPremConnectorTarget = headers.FindFirst(MassMailingPaaSOnPremConnectorTargetName);
                Header LoopPreventionHeader = headers.FindFirst(MassMailingPaaSOnPremConnectorName);

                if (MassMailingPaaSOnPremConnectorTarget != null && evtMessage.MailItem.Message.IsSystemMessage == false && LoopPreventionHeader == null)
                {
                    hasProcessedMessage = true;
                    EventLog.AppendLogEntry(String.Format("Rerouting messages as the control header {0} is present", MassMailingPaaSOnPremConnectorTargetName));
                    MassMailingPaaSOnPremConnectorTargetValue = MassMailingPaaSOnPremConnectorTarget.Value.Trim();

                    if (!String.IsNullOrEmpty(MassMailingPaaSOnPremConnectorTargetValue) && (Uri.CheckHostName(MassMailingPaaSOnPremConnectorTargetValue) == UriHostNameType.Dns))
                    {
                        EventLog.AppendLogEntry(String.Format("Rerouting domain is valid as the header {0} is set to {1}", MassMailingPaaSOnPremConnectorTargetName, MassMailingPaaSOnPremConnectorTargetValue));

                        foreach (EnvelopeRecipient recipient in evtMessage.MailItem.Recipients)
                        {
                            if (ExemptedRecipientDomains.Contains(recipient.Address.DomainPart.ToLower()))
                            {
                                EventLog.AppendLogEntry(String.Format("Recipient {0} not overridden as the recipient domain {1} is present in the registry key {2}", recipient.Address.ToString(), recipient.Address.DomainPart.ToString(), RegistryKeyExemptedRecipientDomains));
                            }
                            else if (ExemptedRecipientAddresses.Contains(recipient.Address.ToString().ToLower()))
                            {
                                EventLog.AppendLogEntry(String.Format("Recipient {0} not overridden as the recipient address {1} is present in the registry key {2}", recipient.Address.ToString(), recipient.Address.ToString(), RegistryKeyExemptedRecipientAddresses));
                            }
                            else
                            {
                                RoutingDomain customRoutingDomain = new RoutingDomain(MassMailingPaaSOnPremConnectorTargetValue);
                                RoutingOverride destinationOverride = new RoutingOverride(customRoutingDomain, DeliveryQueueDomain.UseOverrideDomain);
                                source.SetRoutingOverride(recipient, destinationOverride);
                                EventLog.AppendLogEntry(String.Format("Recipient {0} overridden to {1}", recipient.Address.ToString(), MassMailingPaaSOnPremConnectorTargetValue));
                            }
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
                    else if (LoopPreventionHeader != null)
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is already present", LoopPreventionHeader.Name));
                        EventLog.AppendLogEntry(String.Format("This might mean there is a mail LOOP. Trace the message carefully."));
                        warningOccurred = true;
                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is not set", MassMailingPaaSOnPremConnectorTargetName));
                    }
                }

                EventLog.AppendLogEntry(String.Format("MassMailingPaaSOnPremConnector:RerouteExtrernalBasedOnRegistryExclusions took {0} ms to execute", stopwatch.ElapsedMilliseconds));

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
                EventLog.AppendLogEntry("Exception in MassMailingPaaSOnPremConnector:RerouteExtrernalBasedOnRegistryExclusions");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }
    }
}