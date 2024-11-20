using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ACSOnPremConnector
{
    /*
     * This class reroutes messages to external recipients when the header X-ACSOnPremConnector-Target set to a domain.
     * The domain value doesn't need to be routable, but has to be avalid domain (i.e. something.value.tld).
     * This agent will reroute all the messages via the custom routing domain, only if the target domain is not an accepted domain (i.e. hotmail.com).
     * As the X-ACSOnPremConnector-Target will likely be set via Transport Rule, exclusions can still be managed via the transport rules themselves if necessary.
     * In case multiple agents are active at the same time, only the first one will trigger as the other will detect the presence of the X-ACSOnPremConnector-Target header which is used for loop protection. This is by design to protect mail loops.
     */
    public class RerouteExternalBasedOnAcceptedDomains : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new ACSOnPremConnector_RerouteExternalBasedOnAcceptedDomains(server.AcceptedDomains);
        }
    }

    public class ACSOnPremConnector_RerouteExternalBasedOnAcceptedDomains : RoutingAgent
    {
        EventLogger EventLog = new EventLogger("RerouteExternalBasedOnAcceptedDomains");
        static readonly string ACSOnPremConnectorTargetName = "X-ACSOnPremConnector-Target";
        static string ACSOnPremConnectorTargetValue = String.Empty;

        static readonly string RegistryHive = @"Software\TransportAgents\ACSOnPremConnector\RerouteExternalBasedOnAcceptedDomains";
        static readonly string RegistryKeyDebugEnabled = "DebugEnabled";
        static bool DebugEnabled = false;

        static readonly string ACSOnPremConnectorName = "X-ACSOnPremConnector-Name";
        static readonly string ACSOnPremConnectorNameValue = "ACSOnPremConnector-RerouteExternalBasedOnAcceptedDomains";
        static readonly Dictionary<string, string> ACSOnPremConnectorHeaders = new Dictionary<string, string>
        {
            {ACSOnPremConnectorName, ACSOnPremConnectorNameValue},
            {"X-ACSOnPremConnector-Creator", "Tommaso Toniolo"},
            {"X-ACSOnPremConnector-Contact", "https://aka.ms/totoni"}
        };

        static AcceptedDomainCollection acceptedDomains;

        public ACSOnPremConnector_RerouteExternalBasedOnAcceptedDomains(AcceptedDomainCollection serverAcceptedDomains)
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
                bool warningOccurred = false;
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                HeaderList headers = evtMessage.MailItem.Message.MimeDocument.RootPart.Headers;
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in ACSOnPremConnector:RerouteExternalBasedOnAcceptedDomains", messageId, sender, subject));

                Header ACSOnPremConnectorTarget = headers.FindFirst(ACSOnPremConnectorTargetName);
                Header LoopPreventionHeader = headers.FindFirst(ACSOnPremConnectorName);

                if (ACSOnPremConnectorTarget != null && evtMessage.MailItem.Message.IsSystemMessage == false && LoopPreventionHeader == null)
                {
                    EventLog.AppendLogEntry(String.Format("Rerouting messages as the control header {0} is present", ACSOnPremConnectorTargetName));
                    ACSOnPremConnectorTargetValue = ACSOnPremConnectorTarget.Value.Trim();

                    if (!String.IsNullOrEmpty(ACSOnPremConnectorTargetValue) && (Uri.CheckHostName(ACSOnPremConnectorTargetValue) == UriHostNameType.Dns))
                    {
                        EventLog.AppendLogEntry(String.Format("Rerouting domain is valid as the header {0} is set to {1}", ACSOnPremConnectorTargetName, ACSOnPremConnectorTargetValue));

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
                                RoutingDomain customRoutingDomain = new RoutingDomain(ACSOnPremConnectorTargetValue);
                                RoutingOverride destinationOverride = new RoutingOverride(customRoutingDomain, DeliveryQueueDomain.UseOverrideDomain);
                                source.SetRoutingOverride(recipient, destinationOverride);
                                EventLog.AppendLogEntry(String.Format("Recipient {0} overridden to {1}", recipient.Address.ToString(), ACSOnPremConnectorTargetValue));
                            }
                        }
                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("There was a problem processing the {0} header value", ACSOnPremConnectorTargetName));
                        EventLog.AppendLogEntry(String.Format("There value retrieved is: {0}", ACSOnPremConnectorTargetValue));
                        warningOccurred = true;
                    }

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
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is not set", ACSOnPremConnectorTargetName));
                    }
                }

                EventLog.AppendLogEntry(String.Format("ACSOnPremConnector:RerouteExternalBasedOnAcceptedDomains took {0} ms to execute", stopwatch.ElapsedMilliseconds));

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
                EventLog.AppendLogEntry("Exception in ACSOnPremConnector:RerouteExternalBasedOnAcceptedDomains");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }
    }
}