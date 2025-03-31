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

    public class RewriteRecipientDomain : RoutingAgentFactory
    {
        /*
         * This class rewrites the recipeint domain for those recipients mathicng the content of the header X-MassMailingPaaSOnPremConnector-RecipientRewriteMap.
         * The header value is expected to be a semicolon (;) separated string mapping "<original>=<desired>" with the "=" separating the value fo each entry, such as "contoso.com=tailspintoys.com" or "contoso.com=tailspintoys.com;hotmail.it=hotmail.com".
         * The domain value doesn't need to be routable, but has to be avalid domain (i.e. something.value.tld).
         * As the X-MassMailingPaaSOnPremConnector-RecipientRewriteMap will likely be set via Transport Rule, exclusions can be managed via the transport rules themselves (i.e. insert the X-MassMailingPaaSOnPremConnector-RecipientRewriteMap header only if the recipient domain is not xyz).
         * In case multiple agents are active at the same time, be careful about the agent execution oder.
         */
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new MassMailingPaaSOnPremConnector_RewriteRecipientDomain();
        }
    }

    public class MassMailingPaaSOnPremConnector_RewriteRecipientDomain : RoutingAgent
    {
        static string EventLogName = "RewriteRecipientDomain";
        EventLogger EventLog = new EventLogger(EventLogName);

        static readonly string MassMailingPaaSOnPremConnectorTargetName = "X-MassMailingPaaSOnPremConnector-RecipientRewriteMap";
        static string MassMailingPaaSOnPremConnectorTargetValue = String.Empty;

        static readonly string RegistryHive = @"Software\TransportAgents\MassMailingPaaSOnPremConnector\RewriteRecipientDomain";
        static readonly string RegistryKeyDebugEnabled = "DebugEnabled";
        static bool DebugEnabled = false;

        static readonly string MassMailingPaaSOnPremConnectorName = "X-MassMailingPaaSOnPremConnector-Name";
        static readonly string MassMailingPaaSOnPremConnectorNameValue = "MassMailingPaaSOnPremConnector-RewriteRecipientDomain";
        static readonly Dictionary<string, string> MassMailingPaaSOnPremConnectorHeaders = new Dictionary<string, string>
        {
            {MassMailingPaaSOnPremConnectorName, MassMailingPaaSOnPremConnectorNameValue}
        };

        public MassMailingPaaSOnPremConnector_RewriteRecipientDomain()
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(RewriteRecipientDomain);

            RegistryKey registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
            if (registryPath != null)
            {
                string registryKeyValue = null;
                bool valueConversionResult = false;

                registryKeyValue = registryPath.GetValue(RegistryKeyDebugEnabled, Boolean.FalseString).ToString();
                valueConversionResult = Boolean.TryParse(registryKeyValue, out DebugEnabled);
            }
        }

        void RewriteRecipientDomain(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                bool warningOccurred = false;
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                HeaderList headers = evtMessage.MailItem.Message.MimeDocument.RootPart.Headers;
                Stopwatch stopwatch = Stopwatch.StartNew();
                Dictionary<string, string> RecipientRewriteMap = new Dictionary<string, string>();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in MassMailingPaaSOnPremConnector:RewriteRecipientDomain", messageId, sender, subject));

                Header MassMailingPaaSOnPremConnectorTarget = headers.FindFirst(MassMailingPaaSOnPremConnectorTargetName);

                if (MassMailingPaaSOnPremConnectorTarget != null && evtMessage.MailItem.Message.IsSystemMessage == false)
                {
                    EventLog.AppendLogEntry(String.Format("Rewriting applicable recipients as the messages as the control header {0} is present", MassMailingPaaSOnPremConnectorTargetName));
                    MassMailingPaaSOnPremConnectorTargetValue = MassMailingPaaSOnPremConnectorTarget.Value.Trim();

                    if (!String.IsNullOrEmpty(MassMailingPaaSOnPremConnectorTargetValue))
                    {
                        RecipientRewriteMap = MassMailingPaaSOnPremConnectorTargetValue.ToLower()
                                                                                       .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                                                       .Select(part => part.Split('='))
                                                                                       .ToDictionary(split => split[0].Trim(), split => split[1].Trim());

                        EventLog.AppendLogEntry(String.Format("Recipient domain rewite map start"));
                        foreach (var MapEntry in RecipientRewriteMap)
                        {
                            EventLog.AppendLogEntry(String.Format("\t{0} : {1}", MapEntry.Key, MapEntry.Value));
                        }
                        EventLog.AppendLogEntry(String.Format("Recipient domain rewite map end"));

                        // Rewriting P1 recipients (RPCT TO:)
                        for (int i = 0; i < evtMessage.MailItem.Recipients.Count; i++)
                        {
                            RoutingAddress msgRecipient = evtMessage.MailItem.Recipients[i].Address;
                            string rcptLocal = msgRecipient.LocalPart.ToLower();
                            string rcptDomain = msgRecipient.DomainPart.ToLower();
                            string rcptNewDomain = string.Empty;

                            EventLog.AppendLogEntry(String.Format("Evaluating P1 recipient: {0}", msgRecipient.ToString()));

                            if (RecipientRewriteMap.ContainsKey(rcptDomain))
                            {
                                rcptNewDomain = RecipientRewriteMap[rcptDomain];
                                evtMessage.MailItem.Recipients[i].Address = new RoutingAddress(rcptLocal, rcptNewDomain);
                                EventLog.AppendLogEntry(String.Format("Recipient {0}@{1} rewritten to {2}@{3}", rcptLocal, rcptDomain, rcptLocal, rcptNewDomain));
                            }
                        }

                        // Rewriting P2 TO recipients (TO:)
                        // If the DisplayName is not resolved to "Name Surname" but is an email address, the DisplayName will be set to the new domain as well.
                        for (int i = 0; i < evtMessage.MailItem.Message.To.Count; i++)
                        {
                            string msgRecipient = evtMessage.MailItem.Message.To[i].SmtpAddress;
                            int atIndex = msgRecipient.IndexOf("@");
                            int recLength = msgRecipient.Length;
                            string rcptLocal = msgRecipient.Substring(0, atIndex);
                            string rcptDomain = msgRecipient.Substring(atIndex + 1, recLength - atIndex - 1);
                            string rcptNewDomain = string.Empty;

                            EventLog.AppendLogEntry(String.Format("Evaluating P2 TO recipient: {0}", msgRecipient));

                            if (RecipientRewriteMap.ContainsKey(rcptDomain))
                            {
                                rcptNewDomain = RecipientRewriteMap[rcptDomain];
                                evtMessage.MailItem.Message.To[i].SmtpAddress = rcptLocal + "@" + rcptNewDomain;
                                EventLog.AppendLogEntry(String.Format("P2 TO Recipient {0}@{1} rewritten to {2}@{3}", rcptLocal, rcptDomain, rcptLocal, rcptNewDomain));
                                if (evtMessage.MailItem.Message.To[i].DisplayName.Contains("@"))
                                {
                                    evtMessage.MailItem.Message.To[i].DisplayName = rcptLocal + "@" + rcptNewDomain;
                                    EventLog.AppendLogEntry(String.Format("P2 TO DisplayName {0}@{1} rewritten to {2}@{3} as it contains just the Smtp Address", rcptLocal, rcptDomain, rcptLocal, rcptNewDomain));
                                }
                            }
                        }

                        // Rewriting P2 CC recipients (CC:)
                        // If the DisplayName is not resolved to "Name Surname" but is an email address, the DisplayName will be set to the new domain as well.
                        for (int i = 0; i < evtMessage.MailItem.Message.Cc.Count; i++)
                        {
                            string msgRecipient = evtMessage.MailItem.Message.Cc[i].SmtpAddress;
                            int atIndex = msgRecipient.IndexOf("@");
                            int recLength = msgRecipient.Length;
                            string rcptLocal = msgRecipient.Substring(0, atIndex);
                            string rcptDomain = msgRecipient.Substring(atIndex + 1, recLength - atIndex - 1);
                            string rcptNewDomain = string.Empty;

                            EventLog.AppendLogEntry(String.Format("Evaluating P2 CC recipient: {0}", msgRecipient));

                            if (RecipientRewriteMap.ContainsKey(rcptDomain))
                            {
                                rcptNewDomain = RecipientRewriteMap[rcptDomain];
                                evtMessage.MailItem.Message.Cc[i].SmtpAddress = rcptLocal + "@" + rcptNewDomain;
                                EventLog.AppendLogEntry(String.Format("P2 CC Recipient {0}@{1} rewritten to {2}@{3}", rcptLocal, rcptDomain, rcptLocal, rcptNewDomain));
                                if (evtMessage.MailItem.Message.Cc[i].DisplayName.Contains("@"))
                                {
                                    evtMessage.MailItem.Message.Cc[i].DisplayName = rcptLocal + "@" + rcptNewDomain;
                                    EventLog.AppendLogEntry(String.Format("P2 CC DisplayName {0}@{1} rewritten to {2}@{3} as it contains just the Smtp Address", rcptLocal, rcptDomain, rcptLocal, rcptNewDomain));
                                }
                            }
                        }

                        // Rewriting P2 BCC recipients (BCC:)
                        // If the DisplayName is not resolved to "Name Surname" but is an email address, the DisplayName will be set to the new domain as well.
                        for (int i = 0; i < evtMessage.MailItem.Message.Bcc.Count; i++)
                        {
                            string msgRecipient = evtMessage.MailItem.Message.Bcc[i].SmtpAddress;
                            int atIndex = msgRecipient.IndexOf("@");
                            int recLength = msgRecipient.Length;
                            string rcptLocal = msgRecipient.Substring(0, atIndex);
                            string rcptDomain = msgRecipient.Substring(atIndex + 1, recLength - atIndex - 1);
                            string rcptNewDomain = string.Empty;

                            EventLog.AppendLogEntry(String.Format("Evaluating P2 BCC recipient: {0}", msgRecipient));

                            if (RecipientRewriteMap.ContainsKey(rcptDomain))
                            {
                                rcptNewDomain = RecipientRewriteMap[rcptDomain];
                                evtMessage.MailItem.Message.Bcc[i].SmtpAddress = rcptLocal + "@" + rcptNewDomain;
                                EventLog.AppendLogEntry(String.Format("P2 BCC Recipient {0}@{1} rewritten to {2}@{3}", rcptLocal, rcptDomain, rcptLocal, rcptNewDomain));
                                if (evtMessage.MailItem.Message.Bcc[i].DisplayName.Contains("@"))
                                {
                                    evtMessage.MailItem.Message.Bcc[i].DisplayName = rcptLocal + "@" + rcptNewDomain;
                                    EventLog.AppendLogEntry(String.Format("P2 BCC DisplayName {0}@{1} rewritten to {2}@{3} as it contains just the Smtp Address", rcptLocal, rcptDomain, rcptLocal, rcptNewDomain));
                                }
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
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is not set", MassMailingPaaSOnPremConnectorTargetName));
                    }
                }

                EventLog.AppendLogEntry(String.Format("MassMailingPaaSOnPremConnector:RewriteRecipientDomain took {0} ms to execute", stopwatch.ElapsedMilliseconds));

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
                EventLog.AppendLogEntry("Exception in MassMailingPaaSOnPremConnector:RewriteRecipientDomain");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;


        }
    }
}
