# Name: Report Certificate and Secret Expiry for Azure Tenant

Description: Searches one or more Azure Tenants for Certificates and Client Secrets. The script reports on all credentials discovered, with their expiry date and a 'hint' that identifies the secret.
If expired or soon-to-expire credentials are discovered, an event log is written - this can be used as a trigger to generate an alert
The Application specified in the credential set must have the following permissions:
Application.Read.All (mandatory) - to read the secret metadata attached to the application
User.Read.All (mandatory) - to report the owner name and contact details
Directory.Read.All (optional) - to report the tenant name


Version: 2.0.20

Creator: Bill Powell

Date Created: 11/09/2023 20:14:35

Date Modified: 03/24/2024 14:09:57

Scripting Language: ps1

