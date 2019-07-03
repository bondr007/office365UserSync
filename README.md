# office365UserSync
Powershell Sync users script from CSV 

# Requires MSOnline powershell module
run Install-Module MSOnline

# what it do
script does a delta sync between csv. It is desinged to be passed a csv with new information on a schedule, it compares it to the last csv it process and syncs the difference to Office 365. It also querys active directory for some user information. 

# Disclaimer
This is provided for example purposes only. 
