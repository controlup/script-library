# Name: Check Profile Sizes

Description: Check Profile Sizes examines user profiles for all or selected user accounts on the target machine, grouping the results by file type, using the extension.

For each group of files, if the size of the group exceeds a threshold (default 15% of the total profile size) the individual files are listed, sorted by path or by size (descending) and showing the actual file size in bytes.

To keep the output reasonably short, a threshold is set on the number of files shown individually per are listed by path (default 6) - beyond this, files are summarized by folder (order by count of files, descending).

Arguments:
ThresholdPercentToExpand (default: 15) - the threshold percent of the total profile size at which a file-extension group is listed.
SamAccountNameList (default: All) - the list of account names to be reported (comma-separated, any leading or trailing spaces will be trimmed). If set to All, the script will include local user and Active Directory user accounts.
SortBy (default: Size) - must be set to Size (individual files are listed by size, descending) or Path (individual files are listed by full path, ascending).
PreSummarySize (default: 6) - the number of files that will be listed individually (by group, according to the configured sort order) before the script switches to reporting files grouped by folder.

Version: 1.2.17

Creator: Bill Powell

Date Created: 03/08/2023 17:37:42

Date Modified: 03/23/2023 15:13:00

Scripting Language: ps1

