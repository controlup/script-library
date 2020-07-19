@ECHO OFF
cd %1

set argC=0
for %%x in (%*) do Set /A argC+=1


IF [%argC%] EQU [3]  (
       ECHO Password not provided - recommended to use an authorized key in the default settings for Putty
       start putty.exe %2@%3
       exit
       )

       ECHO 4th parameter detected, using password as %3 and IP as %4
       start putty.exe %2@%4  -pw %3
exit





