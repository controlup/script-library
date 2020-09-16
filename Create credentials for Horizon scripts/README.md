# Name: Create credentials for Horizon scripts

Description: Connecting to a Horizon Connection server is required for running Horizon scripts. The server does not allow passthrough (Active Directory) authentication. In order to allow scripts to run without asking for a password each time (such as in Automated Actions) a PSCredential object needs to be stored on each target device (i.e. each machine that will be used for running Horizon scripts). This script can create this PSCredential object on the targets.
PSCREDENTI    Connecting to a Horizon Connection server is required for running Horizon View scripts. The server does not allow passthrough (Active Directory) authentication. In order to allow scripts to run without asking for a password each time (such as in Automated Actions) a PSCredential
    object needs to be stored on each target device (ie. each machine that will be used for running Horizon View scripts). This script can create this PSCredential object on the targets.
    PSCREDENTIAL OBJECTS CAN ONLY BE USED BY THE USER THAT CREATED THE OBJECT AND ON THE MACHINE THE OBJECT WAS CREATED.
    - The User that creates the file is required to have a local profile when creating the file. This is a limitation from Powershell
    
    Modification history:   20/08/2019 - Anthonie de Vreede - First version
                            03/06/2020 - Wouter Kursten - Second version

    Changelog ;
        Second Version
            - Added check for local profile
            - changed error message when failing to create the xml fileAL OBJECTS CAN ONLY BE USED BY THE USER THAT CREATED THE OBJECT AND ON THE MACHINE THE OBJECT WAS CREATED.

Version: 3.6.12

Creator: Ton de Vreede

Date Created: 12/02/2019 14:12:29

Date Modified: 09/11/2020 08:11:51

Scripting Language: ps1

