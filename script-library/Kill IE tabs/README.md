# Name: Kill IE tabs

Description: This script finds Internet Explorer processes corresponding to tabs in which the URL matches the pattern provided by the user in the "Tab Title Pattern" argument, and kills these processes.

Internet Explorer may recover the killed tab automatically, which is a default behavior controlled by the  the "Enable automatic crash recovery" setting. When set to "yes", the "Disable Tab Recovery" setting of this script will prevent the killed tab from getting reopened. Tab recovery will then remain disabled at the user level.

If the Force parameter is set to "yes", the script will terminate the processes it found, even if their count does not correlate to the number of tabs matching the provided pattern. Be advised that this option may result in closing more tabs than intended, and should be used with caution.

Version: 1.5.28

Creator: Guy Leech

Date Created: 04/01/2019 15:47:15

Date Modified: 04/10/2019 11:51:43

Scripting Language: ps1

