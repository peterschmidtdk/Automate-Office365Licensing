# Author: Peter Schmidt (psc@globeteam.com) - Blog: www.msdigest.net
# The Azure AD bits, are based on the original Azure AD script made by Johan Dahlbom (365lab.net)
# Last updated: 2017.01.13
# Version: 0.2
# Requirements: Microsoft Online and Azure AD Preview v2 modules for PowerShell 

#Import Modules
Import-Module AzureADPreview
Import-Module MSOnline
    
#Office 365 Admin Credentials
$CloudUsername = 'admin@fabrikam.onmicrosoft.com'
$CloudPassword = ConvertTo-SecureString 'password' -AsPlainText -Force
$CloudCred = New-Object System.Management.Automation.PSCredential $CloudUsername, $CloudPassword
    
#Connect to Azure AD 
#Connect-AzureAD -Credential $CloudCred

#Connect to Office 365 
Connect-MsolService -Credential $CloudCred

#Get members from Azure AD Group
$GroupLicensing = Get-AzureADGroup -SearchString O365_E3_Skype | Get-AzureADGroupMember
Write-Host $GroupLicensing.UserPrincipalName

Write-host -ForegroundColor Cyan "---------------------------------------------------------------------------"
Write-host -ForegroundColor Cyan "Checks if User are member of group O365_E3_Skype and assigns licensed plans"
Write-host -ForegroundColor Cyan "---------------------------------------------------------------------------"

foreach ($member in $GroupLicensing) {
    #write-host $member 
    $UserToLicense = $member 

    #Define the plans that will be enabled (Exchange Online, Skype for Business and Office 365 ProPlus )
    #$EnabledPlans = 'MCOSTANDARD','MCOSTANDARD','OFFICESUBSCRIPTION'
    $EnabledPlans = 'MCOSTANDARD'
    #Get the LicenseSKU and create the Disabled ServicePlans object
    $LicenseSku = Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -eq 'ENTERPRISEPACK'} 
    #Loop through all the individual plans and disable all plans except the one in $EnabledPlans
    $DisabledPlans = $LicenseSku.ServicePlans | ForEach-Object -Process { 
      $_ | Where-Object -FilterScript {$_.ServicePlanName -notin $EnabledPlans }
    }
 
    #Create the AssignedLicense object with the License and DisabledPlans earlier created
    $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $License.SkuId = $LicenseSku.SkuId
    $License.DisabledPlans = $DisabledPlans.ServicePlanId
 
    #Create the AssignedLicenses Object 
    $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $AssignedLicenses.AddLicenses = $License
    $AssignedLicenses.RemoveLicenses = @()

    #Assign the license to the user
    Set-AzureADUserLicense -ObjectId $UserToLicense.ObjectId -AssignedLicenses $AssignedLicenses
    write-host "Set license to" $UserToLicense.UserPrincipalName
}

Write-host -ForegroundColor Cyan "------------------------------------------------------------------------------------------------"
Write-host -ForegroundColor Cyan "Checks if User is license and member of the licensing Groups, if not - then licenses are removed"
Write-host -ForegroundColor Cyan "------------------------------------------------------------------------------------------------"

$RemoveLicensing = Get-MsolUser -All | Where-Object { $_.isLicensed -eq "TRUE" }
Write-Host "Users with a license: " $RemoveLicensing.UserPrincipalName

foreach ($members in $RemoveLicensing) {
$License = (Get-MsolUser -UserPrincipalName $members.UserPrincipalName).Licenses.AccountSkuId

write-host "Each Licensed User:" $members.UserPrincipalName
write-host "License attached:" $OldLicense
$UserToRemove = $members.UserPrincipalName 

if (Get-MsolGroupMember -All -GroupObjectId (Get-MsolGroup -All | Where-Object {$_.DisplayName -eq "O365_E3_Skype"}).ObjectId | Where-Object {$_.EmailAddress -eq $UserToRemove}) {

write-host $UserToRemove "is member of O365_E3_Skype group"

} Elseif (Get-MsolGroupMember -All -GroupObjectId (Get-MsolGroup -All | Where-Object {$_.DisplayName -eq "o365_SharePoint"}).ObjectId | Where-Object {$_.EmailAddress -eq $UserToRemove}) {  

write-host $UserToRemove "is member of O365_SharePoint group"

} Elseif (Get-MsolGroupMember -All -GroupObjectId (Get-MsolGroup -All | Where-Object {$_.DisplayName -eq "o365_Yammer"}).ObjectId | Where-Object {$_.EmailAddress -eq $UserToRemove}) {  

write-host $UserToRemove "is member of O365_Yammer group"

}  Else {  

        #The user is no longer a member of any license group, remove license
        Write-Warning "$UserToRemove is not a member of any group, license will be removed... "
            try {  
                Set-MsolUserLicense -UserPrincipalName $UserToRemove -RemoveLicenses $License -ErrorAction Stop -WarningAction Stop

                Write-Output "SUCCESS: Removed $OldLicense for $UserToRemove" -ForegroundColor Green
                } catch {  
                Write-Warning "Error when removing license on user`r`n$_"

        }
 }
}

