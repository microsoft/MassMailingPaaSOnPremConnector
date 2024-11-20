# ACSOnPremConnector

This repository implements the Azure Communication Services (ACS) On-Prem Connector functionality through a Transport Agent for Microsoft Exchange Server 2019. 
This allows ACS-incompatible devices sucha as the ones not capable of authenticating via Username and Password, as those relying on IP-based authentication, Certiicate Authentication, or other authentication mechanisms common to mail servers to leverage ACS.
To achieve this goal, Exchange will act as a bridge between the ACS and the devices, intercepting messages during transport and re-routing the same to ACS.

## Documentation

For detailed information refer to the avaialble [Wiki](https://github.com/kavejo/ACSOnPremConnector/wiki)

## Disclaimer

This code is provided "as is", as a sample without warranty of any kind.
Microsoft and myself further disclaims all implied warranties including without limitation any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the samples remains with you. In no event shall myself, Microsoft or its suppliers be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the samples, even if Microsoft has been advised of the possibility of such damages.