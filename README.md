# Mass-Mailing PaaS On-Prem Connector

This is a Mass-Mailing PaaS On-Premises Connector which through a Transport Agent for Microsoft Exchange Server allows incompatible devices such as the ones not capable of authenticating via Username and Password, those relying on IP-based authentication, those reliant on Certificate Authentication, or other authentication mechanisms to leverage mass mailing services.

To provide this functionality, Microsoft Exchange will act as a bridge between the devices/applications and chosen Mass-Mailing PaaS, intercepting messages during transport and re-routing the same to chosen endpoint. This, practically, is an implementation of what is commonly referred as "Conditional Routing".

If Conditional Routing is not required, and all the traffic traversing the messaging infrastructure has to be relayed to Azure Communication Services Email, then the usage of the Mass-Mailing PaaS On-Prem Connector might not be necessary and the solution can be implemented directly with Microsoft Exchange (or any other Mail Transport Server) by setting the chosen Mass-Mailing PaaS service as the downstream smart-host on the MTA directly.

## Documentation

For detailed information refer to the avaialble [Wiki](https://github.com/kavejo/ACSOnPremConnector/wiki)

## Important Note

Whilst this solution has been adopted by organization leveraging Azure Communication Services Email as their chosen Mass-Mailing PaaS, this is not part of Azure Communication Services.

## Disclaimer

This code is provided "as is", as a sample without warranty of any kind.
Microsoft and myself further disclaims all implied warranties including without limitation any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the samples remains with you. In no event shall myself, Microsoft or its suppliers be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the samples, even if Microsoft has been advised of the possibility of such damages.
