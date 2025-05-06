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
     * This agent will reroute all the messages via the custom routing domain, only if the target domain is not an accepted domain (i.e. hotmail.com).
     * As the X-MassMailingPaaSOnPremConnector-Target will likely be set via Transport Rule, exclusions can still be managed via the transport rules themselves if necessary.
     * In case multiple agents are active at the same time, only the first one will trigger as the other will detect the presence of the X-MassMailingPaaSOnPremConnector-Target header which is used for loop protection. This is by design to protect mail loops.
     */
    public class RerouteExternalBasedOnAcceptedDomains : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new MassMailingPaaSOnPremConnector_RerouteExternalBasedOnAcceptedDomains(server.AcceptedDomains);
        }
    }

    public class MassMailingPaaSOnPremConnector_RerouteExternalBasedOnAcceptedDomains : RoutingAgent
    {
        static string EventLogName = "RerouteExternalBasedOnAcceptedDomains";
        EventLogger EventLog = new EventLogger(EventLogName);

        static readonly string MassMailingPaaSOnPremConnectorTargetName = "X-MassMailingPaaSOnPremConnector-Target";
        static string MassMailingPaaSOnPremConnectorTargetValue = String.Empty;

        static readonly string RegistryHive = @"Software\TransportAgents\MassMailingPaaSOnPremConnector\RerouteExternalBasedOnAcceptedDomains";
        static readonly string RegistryKeyDebugEnabled = "DebugEnabled";
        static bool DebugEnabled = false;

        static readonly string MassMailingPaaSOnPremConnectorName = "X-MassMailingPaaSOnPremConnector-Name";
        static readonly string MassMailingPaaSOnPremConnectorNameValue = "MassMailingPaaSOnPremConnector-RerouteExternalBasedOnAcceptedDomains";
        static readonly Dictionary<string, string> MassMailingPaaSOnPremConnectorHeaders = new Dictionary<string, string>
        {
            {MassMailingPaaSOnPremConnectorName, MassMailingPaaSOnPremConnectorNameValue}
        };

        static AcceptedDomainCollection acceptedDomains;

        public MassMailingPaaSOnPremConnector_RerouteExternalBasedOnAcceptedDomains(AcceptedDomainCollection serverAcceptedDomains)
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(RerouteExternalBasedOnAcceptedDomains);

            RegistryKey registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
            if (registryPath != null)
            {
                string registryKeyValue = null;
                bool valueConversionResult = false;

                registryKeyValue = registryPath.GetValue(RegistryKeyDebugEnabled, Boolean.FalseString).ToString();
                valueConversionResult = Boolean.TryParse(registryKeyValue, out DebugEnabled);
            }

            acceptedDomains = serverAcceptedDomains;

        }

        void RerouteExternalBasedOnAcceptedDomains(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
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

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in MassMailingPaaSOnPremConnector:RerouteExternalBasedOnAcceptedDomains", messageId, sender, subject));

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
                            AcceptedDomain resolvedDomain = acceptedDomains.Find(recipient.Address.DomainPart.ToString());
                            EventLog.AppendLogEntry(String.Format("The check of whether the recipient domain {0} is an Accepted Domain has returned {1}", recipient.Address.DomainPart.ToString(), resolvedDomain == null ? "NULL" : resolvedDomain.IsInCorporation.ToString()));

                            if (resolvedDomain != null)
                            {
                                EventLog.AppendLogEntry(String.Format("Recipient {0} not overridden as the recipient domain IS AN ACCEPTED DOMAIN", recipient.Address.ToString()));
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

                EventLog.AppendLogEntry(String.Format("MassMailingPaaSOnPremConnector:RerouteExternalBasedOnAcceptedDomains took {0} ms to execute", stopwatch.ElapsedMilliseconds));

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
                EventLog.AppendLogEntry("Exception in MassMailingPaaSOnPremConnector:RerouteExternalBasedOnAcceptedDomains");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }
    }
}