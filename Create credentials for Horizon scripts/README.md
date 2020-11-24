# Name: Create credentials for Horizon scripts

Description: Connecting to a Horizon Connection server is required for running Horizon scripts. The server does not allow passthrough (Active Directory) authentication. In order to allow scripts to run without asking for a password each time (such as in Automated Actions) a PSCredential object needs to be stored on each target device (i.e. each machine that will be used for running Horizon scripts). This script can create this PSCredential object on the targets.
PSCREDENTIIAL OBJECTS CAN ONLY BE USED BY THE USER THAT CREATED THE OBJECT AND ON THE MACHINE THE OBJECT WAS CREATED.
    - The User that creates the file is required to have a local profile when creating the file. This is a limitation from Powershell
    
    Modification history:   20/08/2019 - Anthonie de Vreede - First version
                            03/06/2020 - Wouter Kursten - Second version

    Changelog ;
        Second Version
            - Added check for local profile
            - changed error message when failing to create the xml file

Version: 4.6.13

Creator: Ton de Vreede

Date Created: 12/02/2019 14:12:29

Date Modified: 11/17/2020 15:01:05

Scripting Language: ps1

