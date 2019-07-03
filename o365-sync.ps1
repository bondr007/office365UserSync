Param(
    [Parameter(Mandatory = $True)]
    [string]$Path

)

#############################################################################
#If Powershell is running the 32-bit version on a 64-bit machine, we 
#need to force powershell to run in 64-bit mode .
#############################################################################
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    write-warning "Y'arg Matey, we're off to 64-bit land....."
    if ($myInvocation.Line) {
        &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile $myInvocation.Line
    }else{
        &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile -file "$($myInvocation.InvocationName)" $args
    }
exit $lastexitcode
}


#############################################################################
# Change the below for you enviroment
#############################################################################

$o365AdminUser = "adminuser@example.onmicrosoft.com"
$o365AdminPass = ""

$LicName = "example:STANDARDWOFFPACK_IW_STUDENT"
$LO = New-MsolLicenseOptions -AccountSkuId $LicName -DisabledPlans "BPOS_S_TODO_2", "AAD_BASIC_EDU", "SCHOOL_DATA_SYNC_P1", "STREAM_O365_E3", "TEAMS1", "INTUNE_O365", "Deskless", "FLOW_O365_P2", "POWERAPPS_O365_P2", "RMS_S_ENTERPRISE", "OFFICE_FORMS_PLAN_2", "PROJECTWORKMANAGEMENT", "SWAY", "YAMMER_EDU", "SHAREPOINTWAC_EDU", "SHAREPOINTSTANDARD_EDU", "EXCHANGE_S_STANDARD", "MCOSTANDARD"

# Used for Write-Log function
. .\logger.ps1

# import AD to check if user exists and if they are disbled
Import-module ActiveDirectory  

# import and connect to MS office 365 stuff
import-module MSOnline;
$msolcredential = New-Object System.Management.Automation.PsCredential($o365AdminUser, (ConvertTo-SecureString $o365AdminPass -AsPlainText -Force));
connect-msolservice -credential $msolcredential;


Function New-RandomComplexPassword ($length = 15) {
    $Assembly = Add-Type -AssemblyName System.Web
    $password = [System.Web.Security.Membership]::GeneratePassword($length, 2)
    return $password + '!'
}

Write-Log "Lets Get This Party Started..." "INFO"
#Counts/ Stats
$ModCounts = @{
    "Names" = 0;
    "ImmutableId" = 0;
    "BlockCredential" = 0;
    "NewUsers" = 0;
    "NotInAD" = 0;
    "TotalUsersProcessed" = 0;
    "UpdatedLicense" = 0
  }

###################################################
## Time to Get Synkie
###################################################

$Banner_csv = Import-Csv $Path
forEach ($CSV_item in $Banner_csv) {
    $ModCounts.TotalUsersProcessed++
    $userID = $CSV_item.USERID
    $userID = $userID.ToLower()
    
    $ADUser = Get-ADUser -Filter {sAMAccountName -eq $userID}
    $ADuserID = $null

    if ($ADUser){
        $ADuserID = $ADUser.SamAccountName
        $ADuserID = $ADuserID.ToLower()
    }

    If ($ADuserID -eq $userID) {
        Write-Log "$CSV_item.USERID Exists in AD. Whoo" "INFO"

        $o365User = Get-MsolUser -UserPrincipalName $CSV_item.EMAIL
        If ($o365User -ne $Null) {
            Write-Log "$CSV_item.EMAIL Exists in o365. Next checking if needs any update." "INFO"
            # checks for update properties:: 
            #DisplayName/ NAMES 
            if ($o365User.DisplayName -ne $CSV_item.DisplayName) {
                Write-Log "$CSV_item.EMAIL:: Updating DisplayName from ($o365User.DisplayName) to ($CSV_item.DisplayName) " "INFO"
                Set-MsolUser -UserPrincipalName $CSV_item.EMAIL -DisplayName $CSV_item.DisplayName -LastName $CSV_item.LastName -FirstName $CSV_item.FirstName
                $ModCounts.Names++
            }
            #ImmutableId
            if ($o365User.ImmutableId -ne $CSV_item.ImmutableId) {
                Write-Log "$CSV_item.EMAIL:: Updating ImmutableId from ($o365User.ImmutableId) to ($CSV_item.ImmutableId) " "INFO"
                Set-MsolUser -UserPrincipalName $CSV_item.EMAIL -ImmutableId $CSV_item.ImmutableId
                $ModCounts.ImmutableId++
            }
            #BlockCredential -- if disbled in AD
            if ($ADUser.Enabled -ne -Not($o365User.BlockCredential)) {
                $newStatus = -Not($ADUser.Enabled )
                Write-Log "$CSV_item.EMAIL:: Updating BlockCredential from ($o365User.BlockCredential) to ($newStatus) " "INFO"
                Set-MsolUser -UserPrincipalName $CSV_item.EMAIL -BlockCredential $newStatus
                $ModCounts.BlockCredential++
            }

            #Licenses
            if($o365User.IsLicensed){
                $NeedsLicense = $true
                foreach ($element in $o365User.Licenses) {
                    if ($element.AccountSkuId -eq $LicName){
                        $NeedsLicense = $false
                    }
                }

                if ($NeedsLicense){
                    Set-MsolUserLicense -UserPrincipalName $CSV_item.EMAIL -AddLicenses $LicName -LicenseOptions $LO
                    $ModCounts.UpdatedLicense++
                }
            }
            else {
                Set-MsolUserLicense -UserPrincipalName $CSV_item.EMAIL -AddLicenses $LicName -LicenseOptions $LO
                $ModCounts.UpdatedLicense++
            }

        }
        else {
            Write-Log "$CSV_item.EMAIL does not Exists in o365. creating user." "INFO"
            $notusedPass = New-RandomComplexPassword
            new-msoluser -UserPrincipalName $CSV_item.EMAIL -AlternateEmailAddresses $CSV_item.EMAIL -ImmutableId $CSV_item.ImmutableId -Displayname $CSV_item.DisplayName -FirstName $CSV_item.FirstName -LastName $CSV_item.LastName -Password $notusedPass -ForceChangePassword 0 -UsageLocation "US" -LicenseAssignment $LicName -LicenseOptions $LO
            $ModCounts.NewUsers++
        }
    }
    Else {
        Write-Log "$CSV_item.USERID does not Exist in AD. Will not process user." "WARN"
        $ModCounts.NotInAD++
    }
       
}  


Write-Log "Its Been real fun y'all, but thats all. Create and Sync for O365 is DONE!!!" "INFO"
#$outputCounts = $ModCounts.ToString()
$str = $ModCounts | Out-String
Write-Log "Stats:: $str " "INFO"





