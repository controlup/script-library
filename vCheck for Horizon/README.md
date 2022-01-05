# Name: vCheck for Horizon

Description: The vCheck for Horizon is based on the regular vCheck for vSphere and uses the Horizon SOAP api's to gather data. It outputs the results to an html file to file and optionally sends it by email.
Required:
-	Connection server fqdn
-	Output Path
-	Send Email: true or false ( default)
-	Use SSL for Email: true or false (default)
Optional:
-	SMTP server
-	From Email Address
-	To Email Address
-	Email Subject
The vCheck for Horizon SBA downloads the regular vCheck for Horizon from Github if no previous download can be found. To get the newest version of the vCheck for Horizon just remove the %programdata%\Controlup\ScriptSupport \vCheck-HorizonView-master folder.
By default the html file will be saved to the supplied output path. Set the enable email to true and configure all parameters to receive an email as well.

Requirements: Configured Horizon Credentials object for the shared account that will be used. See the Horizon Credentials SBA to create this.
Requirements : PowerCLI 12
Source: https://github.com/vCheckReport/vCheck-HorizonView


Version: 1.6.10

Creator: wouter.kursten

Date Created: 05/07/2021 14:48:56

Date Modified: 05/11/2021 18:24:19

Scripting Language: ps1

