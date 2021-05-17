# Name: Clear a Computer OrgID

Description: When a Computer was previously managed by another organization
it will not accept connections form another (new) organization.
To resolve this, the script will stop a running agent 
and clear the registry settings locking a running agent to an Organization. 
This script will also clear any previous ControlUp Agent customizations
The script will restart the ControlUp Service at the end
Documentation:
https://support.controlup.com/hc/en-us/articles/207234285-Computer-already-belongs-to-another-ControlUp-organization
Limitation: Target computers need to be resolved by hostname and the user executing the script must be a local admin in that computer
Due to multiple remote actions, this script might take up to 2 minutes to execute

Version: 2.6.15

Creator: christian.zorn

Date Created: 06/05/2019 14:06:29

Date Modified: 07/31/2019 17:02:01

Scripting Language: BAT

