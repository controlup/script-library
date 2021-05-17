# Name: Clean Windows system drive

Description: This script is used to clean up a disk by deleting the content of directories which are known to accumulate large amounts of useless data. By default, the following folders are emptied of files (if they exist):
%systemroot%\Downloaded Program Files
%systemroot%\Temp
%systemdrive%\Windows.old
%systemdrive%\Temp
%systemdrive%\MSOCache\All Users
%allusersprofile%\Adobe\Setup
%allusersprofile%\Microsoft\Windows Defender\Definition Updates
%allusersprofile%\Microsoft\Windows Defender\Scans
%allusersprofile%\Microsoft\Windows\WER
%SystemRoot%memory.dmp
%SystemRoot%Minidump.dmp

Extra folders to be cleaned and specific files to be removed can be added when running the script.
Option: Run CLEANMGR with all options set, and delete Volume Shadow Copies
Option: Run DISM if CLEANMGR was supposed to run but is not available
Note: The CLEANMGR option and DISM fallback remove installers of updates and old Windows versions. If either of these has been run, IT IS NOT POSSIBLE TO UNINSTALL UPDATES OR REVERT TO AN OLDER WINDOWS BUILD!

Version: 2.6.34

Creator: Ton de Vreede

Date Created: 12/02/2018 12:27:05

Date Modified: 12/09/2020 11:32:33

Scripting Language: ps1

