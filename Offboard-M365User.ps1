#region Description
<#     
       .NOTES
       ==============================================================================
       Created on:         2025/02/20 
       Created by:         Drago Petrovic
       Organization:       MSB365.blog
       Filename:           Offboard-M365User.ps1
       Current version:    V1.0     

       Find us on:
             * Website:         https://www.msb365.blog
             * Technet:         https://social.technet.microsoft.com/Profile/MSB365
             * LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
             * MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
       ==============================================================================

       .DESCRIPTION
       Automate Microsoft 365 User Offboarding with PowerShell           
       

       .NOTES
       This script is based on the script of the AdminDroid team.
       The referencing link is: https://github.com/admindroid-community/powershell-scripts/tree/master/Automate%20M365%20User%20Offboarding
       The script has been extended and rewritten with several functions.
       An important new feature is the HTML report for the documentation.





       .EXAMPLE
       .\Offboard-M365User.ps1
             

       .COPYRIGHT
       Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
       to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
       and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
       WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
       ===========================================================================
       .CHANGE LOG
             V1.00, 2025/02/20 - DrPe - Initial version

             
			 




--- keep it simple, but significant ---


--- by MSB365 Blog ---

#>
#endregion
##############################################################################################################
[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "Offboard Microsoft 365 User"
$RKEY = "MSB365_Offboard-M365User"
###############################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

function ShowEULAPopup($mode)
{
    $EULA = New-Object -TypeName System.Windows.Forms.Form
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $btnAcknowledge = New-Object System.Windows.Forms.Button
    $btnCancel = New-Object System.Windows.Forms.Button

    $EULA.SuspendLayout()
    $EULA.Name = "MIT"
    $EULA.Text = "$Scriptname - License Agreement"

    $richTextBox1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $richTextBox1.Location = New-Object System.Drawing.Point(12,12)
    $richTextBox1.Name = "richTextBox1"
    $richTextBox1.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $richTextBox1.Size = New-Object System.Drawing.Size(776, 397)
    $richTextBox1.TabIndex = 0
    $richTextBox1.ReadOnly=$True
    $richTextBox1.Add_LinkClicked({Start-Process -FilePath $_.LinkText})
    $richTextBox1.Rtf = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fswiss\fprq2\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Microsoft Sans Serif;}}
{\colortbl ;\red0\green0\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mdispDef1\mwrapIndent1440 }\viewkind4\uc1
\pard\widctlpar\f0\fs19\lang1033 MSB365 SOFTWARE MIT LICENSE\par
Copyright (c) 2025 Drago Petrovic\par
$Scriptname \par
\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}These license terms are an agreement between you and MSB365 (or one of its affiliates). IF YOU COMPLY WITH THESE LICENSE TERMS, YOU HAVE THE RIGHTS BELOW. BY USING THE SOFTWARE, YOU ACCEPT THESE TERMS.\par
\par
MIT License\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}\par
\pard
{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}
\fi-360\li360 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \par
\pard\widctlpar\par
\pard\widctlpar\li360 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
\pard\widctlpar\fi-360\li360 2.\tab THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 3.\tab IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 4.\tab DISCLAIMER OF WARRANTY. THE SOFTWARE IS PROVIDED \ldblquote AS IS,\rdblquote  WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL MSB365 OR ITS LICENSORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THE SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 5.\tab LIMITATION ON AND EXCLUSION OF DAMAGES. IF YOU HAVE ANY BASIS FOR RECOVERING DAMAGES DESPITE THE PRECEDING DISCLAIMER OF WARRANTY, YOU CAN RECOVER FROM MICROSOFT AND ITS SUPPLIERS ONLY DIRECT DAMAGES UP TO U.S. $1.00. YOU CANNOT RECOVER ANY OTHER DAMAGES, INCLUDING CONSEQUENTIAL, LOST PROFITS, SPECIAL, INDIRECT, OR INCIDENTAL DAMAGES. This limitation applies to (i) anything related to the Software, services, content (including code) on third party Internet sites, or third party applications; and (ii) claims for breach of contract, warranty, guarantee, or condition; strict liability, negligence, or other tort; or any other claim; in each case to the extent permitted by applicable law. It also applies even if MSB365 knew or should have known about the possibility of the damages. The above limitation or exclusion may not apply to you because your state, province, or country may not allow the exclusion or limitation of incidental, consequential, or other damages.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 6.\tab ENTIRE AGREEMENT. This agreement, and any other terms MSB365 may provide for supplements, updates, or third-party applications, is the entire agreement for the software.\par
\pard\widctlpar\qj\par
\pard\widctlpar\fi-360\li360\qj 7.\tab A complete script documentation can be found on the website https://www.msb365.blog.\par
\pard\widctlpar\par
\pard\sa200\sl276\slmult1\f1\fs22\lang9\par
\pard\f2\fs17\lang2057\par
}
"@
    $richTextBox1.BackColor = [System.Drawing.Color]::White
    $btnAcknowledge.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAcknowledge.Location = New-Object System.Drawing.Point(544, 415)
    $btnAcknowledge.Name = "btnAcknowledge";
    $btnAcknowledge.Size = New-Object System.Drawing.Size(119, 23)
    $btnAcknowledge.TabIndex = 1
    $btnAcknowledge.Text = "Accept"
    $btnAcknowledge.UseVisualStyleBackColor = $True
    $btnAcknowledge.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::Yes})

    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Location = New-Object System.Drawing.Point(669, 415)
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = New-Object System.Drawing.Size(119, 23)
    $btnCancel.TabIndex = 2
    if($mode -ne 0)
    {
   $btnCancel.Text = "Close"
    }
    else
    {
   $btnCancel.Text = "Decline"
    }
    $btnCancel.UseVisualStyleBackColor = $True
    $btnCancel.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::No})

    $EULA.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
    $EULA.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $EULA.ClientSize = New-Object System.Drawing.Size(800, 450)
    $EULA.Controls.Add($btnCancel)
    $EULA.Controls.Add($richTextBox1)
    if($mode -ne 0)
    {
   $EULA.AcceptButton=$btnCancel
    }
    else
    {
        $EULA.Controls.Add($btnAcknowledge)
   $EULA.AcceptButton=$btnAcknowledge
        $EULA.CancelButton=$btnCancel
    }
    $EULA.ResumeLayout($false)
    $EULA.Size = New-Object System.Drawing.Size(800, 650)

    Return ($EULA.ShowDialog())
}

function ShowEULAIfNeeded($toolName, $mode)
{
$eulaRegPath = "HKCU:Software\Microsoft\$RKEY"
$eulaAccepted = "No"
$eulaValue = $toolName + " EULA Accepted"
if(Test-Path $eulaRegPath)
{
$eulaRegKey = Get-Item $eulaRegPath
$eulaAccepted = $eulaRegKey.GetValue($eulaValue, "No")
}
else
{
$eulaRegKey = New-Item $eulaRegPath
}
if($mode -eq 2) # silent accept
{
$eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
else
{
if($eulaAccepted -eq "No")
{
$eulaAccepted = ShowEULAPopup($mode)
if($eulaAccepted -eq [System.Windows.Forms.DialogResult]::Yes)
{
        $eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
}
}
return $eulaAccepted
}

if ($accepteula)
    {
         ShowEULAIfNeeded "DS Authentication Scripts:" 2
         "EULA Accepted"
    }
else
    {
        $eulaAccepted = ShowEULAIfNeeded "DS Authentication Scripts:" 0
        if($eulaAccepted -ne "Yes")
            {
                "EULA Declined"
                exit
            }
         "EULA Accepted"
    }
###############################################################################
write-host "  _           __  __ ___ ___   ____  __ ___  " -ForegroundColor Yellow
write-host " | |__ _  _  |  \/  / __| _ ) |__ / / /| __| " -ForegroundColor Yellow
write-host " | '_ \ || | | |\/| \__ \ _ \  |_ \/ _ \__ \ " -ForegroundColor Yellow
write-host " |_.__/\_, | |_|  |_|___/___/ |___/\___/___/ " -ForegroundColor Yellow
write-host "       |__/                                  " -ForegroundColor Yellow
Start-Sleep -s 2
write-host ""                                                                                   
write-host ""
write-host ""
write-host ""
###############################################################################


#----------------------------------------------------------------------------------------
Function Get-UPNInput {
    $upn = Read-Host "Enter the UPN of the user to be offboarded"
    return $upn.Trim()
}

Function Show-Progress {
    param (
        [string]$Step,
        [string]$Status
    )
    Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] $Step : $Status" -ForegroundColor Cyan
}

Function Create-HTMLReport {
    param (
        [string]$UPN,
        [hashtable]$Tasks
    )
    $reportDate = Get-Date -Format "yyyy-MM-dd"
    $reportName = "Offboarding_${UPN}_${reportDate}.html"
    $reportPath = Join-Path "C:\Temp" $reportName
    $logoUrl = "https://msb365.abstergo.ch/wp-content/uploads/2023/07/Logo_Long_Plus_B2.png"
    $adminUPN = $env:USERNAME + "@" + (Get-WmiObject -Class Win32_ComputerSystem).Domain

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Offboarding Report - $UPN</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 0; padding: 20px; }
        .header { text-align: center; margin-bottom: 20px; }
        .header img { max-width: 300px; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .success { color: green; }
        .failure { color: red; }
    </style>
</head>
<body>
    <div class="header">
        <img src="$logoUrl" alt="Company Logo">
        <h1>Offboarding Report</h1>
    </div>
    <p><strong>Offboarded User:</strong> $UPN</p>
    <p><strong>Offboarding Performed By:</strong> $adminUPN</p>
    <table>
        <tr><th>Task</th><th>Status</th></tr>
"@

    foreach ($task in $Tasks.Keys) {
        $statusClass = if ($Tasks[$task] -eq "Success") { "success" } else { "failure" }
        $html += "<tr><td>$task</td><td class=`"$statusClass`">$($Tasks[$task])</td></tr>"
    }

    $html += @"
    </table>
</body>
</html>
"@

    $html | Out-File -FilePath $reportPath -Encoding utf8
    Write-Host "HTML report saved to $reportPath" -ForegroundColor Green
}

Function ConnectModules {
    $MsGraphBetaModule = Get-Module Microsoft.Graph.Beta -ListAvailable
    if ($null -eq $MsGraphBetaModule) { 
        Write-Host "Important: Microsoft Graph Beta module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host "Are you sure you want to install Microsoft Graph Beta module? [Y] Yes [N] No"
        if ($confirm -match "[yY]") { 
            Write-Host "Installing Microsoft Graph Beta module..."
            Install-Module Microsoft.Graph.Beta -Scope CurrentUser -AllowClobber -Force
            Write-Host "Microsoft Graph Beta module is installed in the machine successfully" -ForegroundColor Magenta 
        } 
        else { 
            Write-Host "Exiting. `nNote: Microsoft Graph Beta module must be available in your system to run the script" -ForegroundColor Red
            Exit 
        } 
    }
    $ExchangeOnlineModule = Get-Module ExchangeOnlineManagement -ListAvailable
    if ($null -eq $ExchangeOnlineModule) { 
        Write-Host "Important: Exchange Online module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host "Are you sure you want to install Exchange Online module? [Y] Yes [N] No"
        if ($confirm -match "[yY]") { 
            Write-Host "Installing Exchange Online module..."
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
            Write-Host "Exchange Online Module is installed in the machine successfully" -ForegroundColor Magenta 
        } 
        else { 
            Write-Host "Exiting. `nNote: Exchange Online module must be available in your system to run the script" 
            Exit 
        } 
    }
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Connecting modules (Microsoft Graph and Exchange Online module)...`n"
    try {
        Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All","RoleManagement.ReadWrite.Directory"
        Connect-ExchangeOnline
    }
    catch {
        Write-Host $_.Exception.message -ForegroundColor Red
        Exit
    }
    Write-Host "Microsoft Graph Beta PowerShell module is connected successfully" -ForegroundColor Cyan
    Write-Host "Exchange Online module is connected successfully" -ForegroundColor Cyan
}

Function DisableUser {
    try {
        Update-MgUser -UserId $UPN -AccountEnabled:$false -ErrorAction Stop
        $Script:DisableUserAction = "Success"
    }
    catch {
        $Script:DisableUserAction = "Failed"
        $ErrorLog = "$($UPN) - Disable User Action - Action could not be executed"
        $ErrorLog >> $ErrorsLogFile
    }
}

Function ResetPasswordToRandom {
    $Password = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 8 | ForEach-Object {[char]$_})
    $log = "$UPN - $Password"
    $Pwd = ConvertTo-SecureString $Password -AsPlainText -Force
    try {
        $Passwordprofile = @{
            forceChangePasswordNextSignIn = $true
            password = $Pwd
        }
        Update-MgUser -UserId $UPN -PasswordProfile $Passwordprofile -ErrorAction Stop
        $log >> $PasswordLogFile
        $Script:ResetPasswordToRandomAction = "Success"
    }
    catch {
        $Script:ResetPasswordToRandomAction = "Failed"
        $ErrorLog = "$($UPN) - Reset Password To Random Action - Action could not be executed"
        $ErrorLog >> $ErrorsLogFile
    }
}

Function ResetOfficeName {
    try {
        Update-MgUser -UserId $UPN -OfficeLocation "EXD" -ErrorAction Stop
        $Script:ResetOfficeNameAction = "Success"
    }
    catch {
        $Script:ResetOfficeNameAction = "Failed"
        $ErrorLog = "$($UPN) - Reset Office Name Action - Action could not be executed"
        $ErrorLog >> $ErrorsLogFile
    }
}

Function RemoveMobileNumber {
    try {
        Update-MgUser -UserId $UPN -MobilePhone $null -ErrorAction Stop
        $Script:RemoveMobileNumberAction = "Success"
    }
    catch {
        $Script:RemoveMobileNumberAction = "Failed"
        $ErrorLog = "$($UPN) - Remove Mobile Number Action - Action could not be executed"
        $ErrorLog >> $ErrorsLogFile
    }
}

Function RemoveGroupMemberships {
    $groupMemberships = Get-MgUserMemberOf -UserId $UPN | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
    foreach ($Membership in $groupMemberships) {
        try { 
            Remove-MgGroupMemberByRef -GroupId $Membership.Id -DirectoryObjectId $UserId -ErrorAction Stop
        }
        catch {
            try {
                Remove-DistributionGroupMember -Identity $Membership.Id -Member $UserId -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
            }
            catch {
                $ErrorLog = "$($UPN) - GroupId($($Membership.Id)) - Remove Group Memberships Action - " + $_.Exception.Message
                $ErrorLog >> $ErrorsLogFile
            }
        }
    }
    $UserOwnerships = Get-MgUserOwnedObject -UserId $UPN | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
    foreach ($UserOwnership in $UserOwnerships) {
        try {
            Remove-MgGroupOwnerByRef -GroupId $UserOwnership.Id -DirectoryObjectId $UserId -ErrorAction Stop
        }
        catch {
            $ErrorLog = "$($UPN) - GroupId($($UserOwnership.Id)) - Remove Group Ownerships Action - " + $_.Exception.Message
            $ErrorLog >> $ErrorsLogFile
        }
    }
    $DistributionGroupOwnerships = Get-DistributionGroup | Where-Object { $_.ManagedBy -contains $UserId }
    foreach ($DistributionGroupOwnership in $DistributionGroupOwnerships) {
        try {
            Set-DistributionGroup -Identity $DistributionGroupOwnership.Identity -BypassSecurityGroupManagerCheck -ManagedBy @{Remove=$UPN} -ErrorAction Stop
        }
        catch {
            $ErrorLog = "$($UPN) - GroupId($($DistributionGroupOwnership.ExternalDirectoryObjectId)) - Remove Distribution Group Ownerships Action - " + $_.Exception.Message
            $ErrorLog >> $ErrorsLogFile
        }
    }
    if ($null -eq $ErrorLog) {
        $Script:RemoveGroupMembershipsAction = "Success"
    }
    elseif ($null -eq $groupMemberships -and $null -eq $UserOwnerships -and $null -eq $DistributionGroupOwnerships) {
        $Script:RemoveGroupMembershipsAction = "No group memberships"
    }
    else {
        $Script:RemoveGroupMembershipsAction = "Partial Success"
    }
}

Function RemoveAdminRoles {
    $AdminRoles = Get-MgUserMemberOf -UserId $UPN | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.directoryRole' }
    if ($null -eq $AdminRoles) {
        $Script:RemoveAdminRolesAction = "No admin roles"
    }
    else {
        foreach ($AdminRole in $AdminRoles) {
            try {
                Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $AdminRole.Id -DirectoryObjectId $UserId -ErrorAction Stop
            }
            catch {
                $ErrorLog = "$($UPN) - Role Id($($AdminRole.DisplayName)) Remove Admin Roles Action - " + $_.Exception.Message
                $ErrorLog >> $ErrorsLogFile
            }
        }
        if ($null -eq $ErrorLog) {
            $Script:RemoveAdminRolesAction = "Success"
        }
        else {
            $Script:RemoveAdminRolesAction = "Partial Success"
        }
    }
}

Function RemoveAppRoleAssignments {
    $AppRoleAssignments = Get-MgUserAppRoleAssignment -UserId $UPN
    if ($null -ne $AppRoleAssignments) {
        foreach ($assignment in $AppRoleAssignments) {
            try {
                Remove-MgUserAppRoleAssignment -AppRoleAssignmentId $assignment.Id -UserId $UPN -ErrorAction Stop
            }
            catch {
                $ErrorLog = "$($UPN) - Remove App Role Assignments Action - Action could not be executed"
                $ErrorLog >> $ErrorsLogFile
            }
        }
        if ($null -eq $ErrorLog) {
            $Script:RemoveAppRoleAssignmentsAction = "Success"
        }
        else {
            $Script:RemoveAppRoleAssignmentsAction = "Partial Success"
        }
    }
    else {
        $Script:RemoveAppRoleAssignmentsAction = "No app role assignments"
    }
}

Function HideFromAddressList {
    if ($MailBoxAvailability -eq 'No') {
        $Script:HideFromAddressListAction = "No Exchange license assigned to user"
        return
    }
    try {
        Set-Mailbox -Identity $UPN -HiddenFromAddressListsEnabled $true 
        $Script:HideFromAddressListAction = "Success"
    }
    catch {
        $Script:HideFromAddressListAction = "Failed"
        $ErrorLog = "$($UPN) - Hide From Address List Action - " + $_.Exception.Message
        $ErrorLog >> $ErrorsLogFile
    }
}

Function RemoveEmailAlias {
    if ($MailBoxAvailability -eq 'No') {
        $Script:RemoveEmailAliasAction = "No Exchange license assigned to user"
        return
    }
    try {
        $EmailAliases = Get-Mailbox $UPN | Select-Object -ExpandProperty EmailAddresses | Where-Object { $_ -clike "smtp:*" }
        if ($null -eq $EmailAliases) {
            $Script:RemoveEmailAliasAction = "No alias"
        }
        else {
            Set-Mailbox $UPN -EmailAddresses @{Remove=$EmailAliases} -WarningAction SilentlyContinue
            $Script:RemoveEmailAliasAction = "Success"
        }
    }
    catch {
        $Script:RemoveEmailAliasAction = "Failed"
        $ErrorLog = "$($UPN) - Remove Email Alias Action - " + $_.Exception.Message
        $ErrorLog >> $ErrorsLogFile
    }
}

Function WipingMobileDevice {
    if ($MailBoxAvailability -eq 'No') {
        $Script:MobileDeviceAction = "No Exchange license assigned to user"
        return
    }
    try {
        $MobileDevices = Get-MobileDevice -Mailbox $UPN 
        if ($null -eq $MobileDevices) {
            $Script:MobileDeviceAction = "No mobile devices"
        }
        else {
            $MobileDevices | Clear-MobileDevice -AccountOnly -Confirm:$false
            $Script:MobileDeviceAction = "Success"
        }
    }
    catch {
        $Script:MobileDeviceAction = "Failed"
        $ErrorLog = "$($UPN) - Wiping Mobile Device Action - " + $_.Exception.Message
        $ErrorLog >> $ErrorsLogFile
    }
}

Function DeleteInboxRule {
    if ($MailBoxAvailability -eq 'No') {
        $Script:DeleteInboxRuleAction = "No Exchange license assigned to user"
        return
    }
    try {
        $MailboxRules = Get-InboxRule -Mailbox $UPN 
        if ($null -eq $MailboxRules) {
            $Script:DeleteInboxRuleAction = "No inbox rules"
        }
        else {
            $MailboxRules | Remove-InboxRule -Confirm:$false
            $Script:DeleteInboxRuleAction = "Success"
        }
    }
    catch {
        $Script:DeleteInboxRuleAction = "Failed"
        $ErrorLog = "$($UPN) - Delete Inbox Rule Action - " + $_.Exception.Message
        $ErrorLog >> $ErrorsLogFile
    }
}

Function ConvertToSharedMailbox {
    if ($MailBoxAvailability -eq 'No') {
        $Script:ConvertToSharedMailboxAction = "No Exchange license assigned to user"
        return
    }
    try {
        Set-Mailbox -Identity $UPN -Type Shared -WarningAction SilentlyContinue
        $Script:ConvertToSharedMailboxAction = "Success"
    }
    catch {
        $Script:ConvertToSharedMailboxAction = "Failed"
        $ErrorLog = "$($UPN) - Convert To Shared Mailbox Action - " + $_.Exception.Message
        $ErrorLog >> $ErrorsLogFile
    }
}

Function RemoveLicense {
    $Licenses = Get-MgUserLicenseDetail -UserId $UPN
    if ($null -ne $Licenses) {
        try {
            Set-MgUserLicense -UserId $UPN -RemoveLicenses @($Licenses.SkuId) -AddLicenses @() -ErrorAction Stop
            $Script:RemoveLicenseAction = "Removed licenses - $($Licenses.SkuPartNumber -join ',')"
        }
        catch {
            $Script:RemoveLicenseAction = "Failed"
            $ErrorLog = "$($UPN) - Remove License Action - Action could not be executed" 
            $ErrorLog >> $ErrorsLogFile
        }
    }
    else {
        $Script:RemoveLicenseAction = "No license"
    }
}

Function SignOutFromAllSessions {
    try {
        Revoke-MgUserSignInSession -UserId $UPN -ErrorAction Stop
        $Script:SignOutFromAllSessionsAction = "Success"
    }
    catch {
        $Script:SignOutFromAllSessionsAction = "Failed"
        $ErrorLog = "$($UPN) - Sign Out From All Sessions Action - Action could not be executed"
        $ErrorLog >> $ErrorsLogFile
    }
}

Function Disconnect_Modules {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Exit
}

Function Create-FinalHTMLReport {
    param (
        [hashtable]$AllTasks
    )
    $reportDate = Get-Date -Format "yyyy-MM-dd"
    $reportName = "Final_Offboarding_Report_${reportDate}.html"
    $reportPath = Join-Path "C:\Temp" $reportName
    $logoUrl = "https://msb365.abstergo.ch/wp-content/uploads/2023/07/Logo_Long_Plus_B2.png"
    $adminUPN = $env:USERNAME + "@" + (Get-WmiObject -Class Win32_ComputerSystem).Domain

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Final Offboarding Report - $reportDate</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 0; padding: 20px; }
        .header { text-align: center; margin-bottom: 20px; }
        .header img { max-width: 300px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .success { color: green; }
        .failure { color: red; }
        h2 { margin-top: 30px; }
    </style>
</head>
<body>
    <div class="header">
        <img src="$logoUrl" alt="Company Logo">
        <h1>Final Offboarding Report</h1>
    </div>
    <p><strong>Offboarding Performed By:</strong> $adminUPN</p>
"@

    foreach ($UPN in $AllTasks.Keys) {
        $html += @"
    <h2>User: $UPN</h2>
    <table>
        <tr><th>Task</th><th>Status</th></tr>
"@
        foreach ($task in $AllTasks[$UPN].Keys) {
            $status = $AllTasks[$UPN][$task]
            $statusClass = if ($status -eq "Success") { "success" } else { "failure" }
            $html += "<tr><td>$task</td><td class=`"$statusClass`">$status</td></tr>"
        }
        $html += "</table>"
    }

    $html += @"
</body>
</html>
"@

    $html | Out-File -FilePath $reportPath -Encoding utf8
    Write-Host "Final HTML report saved to $reportPath" -ForegroundColor Green
}

Function main {
    ConnectModules

    $allTasks = @{}  # Hashtable to store all tasks for all users
    $Location = "C:\Temp"
    $ExportCSV = Join-Path $Location "M365UserOffBoarding_StatusFile_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
    $PasswordLogFile = Join-Path $Location "PasswordLogFile_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).txt"
    $ErrorsLogFile = Join-Path $Location "ErrorsLogFile$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).txt"

    do {
        $UPN = Get-UPNInput

        Write-Host "`nWe can perform the following operations:`n" -ForegroundColor Cyan
        Write-Host "           1.  Disable user" -ForegroundColor Yellow
        Write-Host "           2.  Reset password to random" -ForegroundColor Yellow 
        Write-Host "           3.  Reset Office name" -ForegroundColor Yellow 
        Write-Host "           4.  Remove Mobile number" -ForegroundColor Yellow
        Write-Host "           5.  Remove group memberships" -ForegroundColor Yellow
        Write-Host "           6.  Remove admin roles" -ForegroundColor Yellow
        Write-Host "           7.  Remove app role assignments" -ForegroundColor Yellow
        Write-Host "           8.  Hide from address list" -ForegroundColor Yellow
        Write-Host "           9.  Remove email alias" -ForegroundColor Yellow
        Write-Host "           10. Wiping mobile device" -ForegroundColor Yellow
        Write-Host "           11. Delete inbox rule" -ForegroundColor Yellow
        Write-Host "           12. Convert to shared mailbox" -ForegroundColor Yellow
        Write-Host "           13. Remove license" -ForegroundColor Yellow
        Write-Host "           14. Sign-out from all sessions" -ForegroundColor Yellow
        Write-Host "           15. All the above operations" -ForegroundColor Yellow
        Write-Host "           16. Exit" -ForegroundColor Yellow
        $Actions = Read-Host "`nPlease choose the action(s) to continue (comma-separated numbers)"
        
        if ($Actions -eq "") {
            Write-Host "`nPlease choose at least one action from the above." -ForegroundColor Red
            continue
        }
        $Actions = $Actions.Trim() -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        $CheckActions = $Actions | Where-Object { $_ -notin 1..16 }
        if ($CheckActions) {
            Write-Host "`nPlease choose correct action number(s) from the above actions." -ForegroundColor Red
            continue
        }

        if ($Actions -contains 16) {
            Write-Host "Creating final HTML report..." -ForegroundColor Cyan
            Create-FinalHTMLReport -AllTasks $allTasks
            Write-Host "Final HTML report created successfully." -ForegroundColor Green
            Write-Host "Exiting the script." -ForegroundColor Green
            Disconnect_Modules
            break
        }

        Show-Progress "Processing" $UPN
        $Script:Status = "$UPN - "
        $User = Get-MgUser -UserId $UPN -ErrorAction SilentlyContinue 
        $UserId = $User.Id
        if ($null -eq $User) {
            Write-Host "Invalid user: $UPN" -ForegroundColor Red
            Continue
        }
        $MailBox = Get-Mailbox -Identity $UPN -RecipientTypeDetails UserMailbox -ErrorAction SilentlyContinue
        $MailBoxAvailability = if ($null -ne $MailBox) { "Yes" } else { "No" }

        if ($Actions -contains 15) {
            $Actions = 1..14
        }

        $Tasks = @{}  # Hashtable to store task results for HTML report

        foreach ($Action in $Actions) {
            switch ($Action) {
                1  { Show-Progress "Disabling User" "In Progress"; DisableUser; $Tasks["Disable User"] = $DisableUserAction; break }
                2  { Show-Progress "Resetting Password" "In Progress"; ResetPasswordToRandom; $Tasks["Reset Password"] = $ResetPasswordToRandomAction; break }
                3  { Show-Progress "Resetting Office Name" "In Progress"; ResetOfficeName; $Tasks["Reset Office Name"] = $ResetOfficeNameAction; break }
                4  { Show-Progress "Removing Mobile Number" "In Progress"; RemoveMobileNumber; $Tasks["Remove Mobile Number"] = $RemoveMobileNumberAction; break }
                5  { Show-Progress "Removing Group Memberships" "In Progress"; RemoveGroupMemberships; $Tasks["Remove Group Memberships"] = $RemoveGroupMembershipsAction; break }
                6  { Show-Progress "Removing Admin Roles" "In Progress"; RemoveAdminRoles; $Tasks["Remove Admin Roles"] = $RemoveAdminRolesAction; break }
                7  { Show-Progress "Removing App Role Assignments" "In Progress"; RemoveAppRoleAssignments; $Tasks["Remove App Role Assignments"] = $RemoveAppRoleAssignmentsAction; break }
                8  { Show-Progress "Hiding From Address List" "In Progress"; HideFromAddressList; $Tasks["Hide From Address List"] = $HideFromAddressListAction; break }
                9  { Show-Progress "Removing Email Alias" "In Progress"; RemoveEmailAlias; $Tasks["Remove Email Alias"] = $RemoveEmailAliasAction; break }
                10 { Show-Progress "Wiping Mobile Device" "In Progress"; WipingMobileDevice; $Tasks["Wipe Mobile Device"] = $MobileDeviceAction; break }
                11 { Show-Progress "Deleting Inbox Rule" "In Progress"; DeleteInboxRule; $Tasks["Delete Inbox Rule"] = $DeleteInboxRuleAction; break }
                12 { Show-Progress "Converting To Shared Mailbox" "In Progress"; ConvertToSharedMailbox; $Tasks["Convert To Shared Mailbox"] = $ConvertToSharedMailboxAction; break }
                13 { Show-Progress "Removing License" "In Progress"; RemoveLicense; $Tasks["Remove License"] = $RemoveLicenseAction; break }
                14 { Show-Progress "Signing Out From All Sessions" "In Progress"; SignOutFromAllSessions; $Tasks["Sign Out From All Sessions"] = $SignOutFromAllSessionsAction; break }
            }
            Show-Progress $Tasks.Keys[-1] $Tasks.Values[-1]
        }

        $allTasks[$UPN] = $Tasks

        Create-HTMLReport -UPN $UPN -Tasks $Tasks

        $Result = [PSCustomObject]@{
            'UPN' = $UPN
            'Disable User' = $DisableUserAction
            'Reset Password To Random' = $ResetPasswordToRandomAction
            'Reset OfficeName' = $ResetOfficeNameAction
            'Remove Mobile Number' = $RemoveMobileNumberAction
            'Remove Group Memberships' = $RemoveGroupMembershipsAction
            'Remove Admin Roles' = $RemoveAdminRolesAction
            'Remove AppRole Assignments' = $RemoveAppRoleAssignmentsAction
            'Exchange User' = $MailBoxAvailability
            'Hide From Address List' = $HideFromAddressListAction
            'Remove Email Alias' = $RemoveEmailAliasAction
            'Wiping Mobile Device' = $MobileDeviceAction
            'Delete Inbox Rule' = $DeleteInboxRuleAction
            'ConvertToSharedMailbox' = $ConvertToSharedMailboxAction
            'Remove License' = $RemoveLicenseAction
            'SignOut From All Sessions' = $SignOutFromAllSessionsAction
        } 
        $Result | Export-Csv -Path $ExportCSV -Append -NoTypeInformation

        Write-Host "`nScript executed successfully" -ForegroundColor Green
        Write-Host "Status file available in: $ExportCSV" -ForegroundColor Yellow
        if (Test-Path -Path $ErrorsLogFile) {
            Write-Host "Errors log file available in: $ErrorsLogFile" -ForegroundColor Yellow
        }
        if ($Actions -contains 15 -or $Actions -contains 2) {
            Write-Host "Password log file available in: $PasswordLogFile" -ForegroundColor Yellow
        }
        Write-Host "`n~~ Script prepared by MSB365 ~~" -ForegroundColor Green
        

        $continue = Read-Host "`nDo you want to perform another action? (Y/N)"
        if ($continue -notmatch "[yY]") {
            Write-Host "Creating final HTML report..." -ForegroundColor Cyan
            Create-FinalHTMLReport -AllTasks $allTasks
            Write-Host "Final HTML report created successfully." -ForegroundColor Green
            Write-Host "Exiting the script." -ForegroundColor Green
            Disconnect_Modules
            break
        }
    } while ($true)
}

. main

