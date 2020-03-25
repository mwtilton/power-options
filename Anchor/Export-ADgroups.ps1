#Enables the script to be run
set-executionpolicy -scope CurrentUser remotesigned  -force

#Pulls Memberof attribute for all users
get-adprincipalgroupmembership user1 | select name | Export-Csv C:\style.csv -nti

#Re-restrict scripts from running
set-executionpolicy -scope CurrentUser restricted  -force
exit