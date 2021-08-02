######################################################
## -- Script Variables
######################################################

$scriptOffice365ExportRules = "Office 365 - Export Inbox Rules"
$scriptOffice365ExportMailboxes = "Office 365 - Export Mailboxes"
$scriptOffice365ExportAdministrators = "Office 365 - Export Administrators"
$scriptOffice365ExportMailboxAudit = "Office 365 - Export Mailbox Audit"
$scriptOffice365ExportAccountMFAState = "Office 365 - Export Account MFA State"
$scriptOffice365ExportAccountMailForwarding = "Office 365 - Export Mail Forwarding Rules"
$scriptExit = "Exit"

$defaultExportPath = "C:\scriptExports\"

$menuList = @(
    $scriptOffice365ExportRules,
    $scriptOffice365ExportMailboxes,
    $scriptOffice365ExportAdministrators,
    $scriptOffice365ExportMailboxAudit,
    $scriptOffice365ExportAccountMFAState,
    $scriptOffice365ExportAccountMailForwarding,
    $scriptExit
)

$xmin = 3
$ymin = 5

######################################################
## -- Modules
######################################################

# Check to see if the required modules are already available and import them if they are
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement ) {
    Import-Module ExchangeOnlineManagement 
} 
else {
    Install-Module -Name ExchangeOnlineManagement -scope CurrentUser
}
        
# Check to see if the export path already exists
if (!(Test-Path -path $defaultExportPath)) {  
    New-Item -ItemType directory -Path $defaultExportPath
}

######################################################
## -- Menu
######################################################
function ShowMenu {

    while ($true) {
        
        #Write Menu
        Clear-Host
        Write-Host ""
        Write-Host "  Use the up / down arrow to navigate and Enter to make a selection"
        Write-Host ""
        Write-Host "  Default Export Location: ""$($defaultExportPath)"""
        Write-Host ""
        [Console]::SetCursorPosition(0, $ymin)
        foreach ($listItem in $menuList) {
            for ($i = 0; $i -lt $xmin; $i++) {
                Write-Host " " -NoNewline
            }
            Write-Host "   " + $listItem
        }
        
        #Highlight first item by default
        $cursorY = 0
        Write-Highlighted
        
        $selection = ""
        $menu_active = $true
        

        
        while ($menu_active) {
            if ([console]::KeyAvailable) {
                $x = $Host.UI.RawUI.ReadKey()
                [Console]::SetCursorPosition(1, $cursorY)
                Write-Normal
                switch ($x.VirtualKeyCode) { 
                    38 {
                        #down key
                        if ($cursorY -gt 0) {
                            $cursorY = $cursorY - 1
                        }
                    }
                    
                    40 {
                        #up key
                        if ($cursorY -lt $menuList.Length - 1) {
                            $cursorY = $cursorY + 1
                        }
                    }
                    13 {
                        #enter key
                        $selection = $menuList[$cursorY]
                        $menu_active = $false
                    }
                }
                Write-Highlighted
            }
            Start-Sleep -Milliseconds 5
        }
    
        Clear-Host

        switch ( $selection ) {
            $scriptOffice365ExportRules { ExportAllInboxRules -ExportPath $defaultExportPath }
            $scriptOffice365ExportMailboxes { ExportAllMailboxes -ExportPath $defaultExportPath }
            $scriptOffice365ExportAdministrators { ExportAllAdministrators -ExportPath $defaultExportPath }
            $scriptOffice365ExportMailboxAudit { ExportAllMailboxAudit -ExportPath $defaultExportPath }
            $scriptOffice365ExportAccountMFAState { ExportAccountMFAState -ExportPath $defaultExportPath }
            $scriptOffice365ExportAccountMailForwarding { ExportForwardingRules -ExportPath $defaultExportPath }
            $scriptExitSelected { ExitScript }
            default { 'Invalid menu selection.' }
        }
    }
}


######################################################
## -- Script Functions
######################################################
function ExitScript {
    exit 0 
}

function ExportAllInboxRules {
    Param($ExportPath)

    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop 6>$null
    try {
        $users = (Get-EXOMailbox -resultsize unlimited).UserPrincipalName
        foreach ($user in $users) {
            Get-InboxRule -Mailbox $user | Select-Object MailboxOwnerID, Name, Description, Enabled, RedirectTo, MoveToFolder, ForwardTo | Export-CSV "$($ExportPath)InboxRule.csv" -NoTypeInformation -Append
        }
    }
    catch {
        Out-LogFile "Unable to write output to disk" -warning;
        Write-Error $_.Exception.Message;
        Write-Host -ForegroundColor Red "[!] Unable to write output to disk"
        Read-Host
    }
    Read-Host

    Disconnect-ExchangeOnline
}

function ExportAllMailboxes {
    Param($ExportPath)

    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop 6>$null
    try {
        Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName, UserPrincipalName, RecipientTypeDetails | Export-CSV "$($ExportPath)Mailboxes.csv" -NoTypeInformation   
    }
    catch {
        Out-LogFile "Unable to write output to disk" -warning;
        Write-Error $_.Exception.Message;
        Write-Host -ForegroundColor Red "[!] Unable to write output to disk"
        Read-Host
    }
    Read-Host

    Disconnect-ExchangeOnline -Confirm:$false
}

function ExportForwardingRules {
    Param($ExportPath)

    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop 6>$null

    $SMTPForward = Get-Mailbox -ResultSize Unlimited | Where-Object { ($null -ne $_.ForwardingAddress -or $null -ne $_.ForwardingSMTPAddress) } | Select-Object Name, ForwardingAddress, ForwardingSMTPAddress, DeliverToMailboxAndForward;
    
    try {
        $SMTPForward | Export-CSV "$($ExportPath)\MailForwardingRules.csv" -NoTypeInformation;   
    }
    catch {
        Out-LogFile "Unable to write output to disk" -warning;
        Write-Error $_.Exception.Message;
        Write-Host -ForegroundColor Red "[!] Unable to write output to disk"
        Read-Host
    }
    Disconnect-ExchangeOnline -Confirm:$false
    Read-Host
}

function ExportAllAdministrators {
    Param($ExportPath)

    Connect-MsolService -ShowBanner:$false -ErrorAction Stop 6>$null
    try {
        $roles = Get-MsolRole | Where-Object { $_.name -notlike "*Directory Readers*" }

        foreach ($role in $roles) {
            Get-MsolRoleMember -RoleObjectId $role.ObjectID | Select-Object * | Export-CSV "$($ExportPath)Administrators.csv" -NoTypeInformation -Append
        }
    }
    catch {
        Out-LogFile "Unable to write output to disk" -warning;
        Write-Error $_.Exception.Message;
        Write-Host -ForegroundColor Red "[!] Unable to write output to disk"
        Read-Host
    }
    Read-Host
    
    Disconnect-MsolService -Confirm:$false
}

function ExportAllMailboxAudit {
    Param($ExportPath)

    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop 6>$null
    try {
        Get-Mailbox -ResultSize Unlimited -Filter { RecipientTypeDetails -eq "UserMailbox" } | Select-Object DisplayName, WindowsEmailAddress, IsMailboxEnabled, AuditEnabled | Export-CSV "$($ExportPath)MailboxAudit.csv" -NoTypeInformation   
    }
    catch {
        Out-LogFile "Unable to write output to disk" -warning;
        Write-Error $_.Exception.Message;
        Write-Host -ForegroundColor Red "[!] Unable to write output to disk"
        Read-Host
    }
    Read-Host

    Disconnect-ExchangeOnline -Confirm:$false
}

function ExportAccountMFAState {
    Param($ExportPath)

    Connect-MsolService -ShowBanner:$false -ErrorAction Stop 6>$null
    try {
        $Report = [System.Collections.Generic.List[Object]]::new()

        $Users = Get-MsolUser -All | Where-Object { $_.UserType -ne "Guest" }

        ForEach ($User in $Users) {
            $MFAMethods = $User.StrongAuthenticationMethods.MethodType
            $MFAEnforced = $User.StrongAuthenticationRequirements.State
            $MFAPhone = $User.StrongAuthenticationUserDetails.PhoneNumber
            $DefaultMFAMethod = ($User.StrongAuthenticationMethods | Where-Object { $_.IsDefault -eq "True" }).MethodType
            If (($MFAEnforced -eq "Enforced") -or ($MFAEnforced -eq "Enabled")) {
                Switch ($DefaultMFAMethod) {
                    "OneWaySMS" { $MethodUsed = "One-way SMS" }
                    "TwoWayVoiceMobile" { $MethodUsed = "Phone call verification" }
                    "PhoneAppOTP" { $MethodUsed = "Hardware token or authenticator app" }
                    "PhoneAppNotification" { $MethodUsed = "Authenticator app" }
                }
            }
            Else {
                $MFAEnforced = "Not Enabled"
                $MethodUsed = "MFA Not Used" 
            }
  
            $ReportLine = [PSCustomObject] @{
                User        = $User.UserPrincipalName
                Name        = $User.DisplayName
                MFAUsed     = $MFAEnforced
                MFAMethod   = $MethodUsed 
                PhoneNumber = $MFAPhone
            }
                 
            $Report.Add($ReportLine) 
        }

        $Report | Sort-Object Name | Export-CSV "$($ExportPath)MFAStatus.csv" -NoTypeInformation
    }
    catch {
        Out-LogFile "Unable to write output to disk" -warning;
        Write-Error $_.Exception.Message;
        Write-Host -ForegroundColor Red "[!] Unable to write output to disk"
        Read-Host
    }
    Read-Host

    Disconnect-MsolService -Confirm:$false
}

function Write-Highlighted {
     
    [Console]::SetCursorPosition(1 + $xmin, $cursorY + $ymin)
    Write-Host ">" -BackgroundColor Yellow -ForegroundColor Black -NoNewline
    Write-Host " " + $menuList[$cursorY] -BackgroundColor Yellow -ForegroundColor Black
    [Console]::SetCursorPosition(0, $cursorY + $ymin)     
}

function Write-Normal {
    [Console]::SetCursorPosition(1 + $xmin, $cursorY + $ymin)
    Write-Host "  " + $menuList[$cursorY]  
}

ShowMenu
