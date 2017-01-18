#Requires -Modules PowerShellApplication, TervisMailMessage, PasswordStatePowerShell, StringPowerShell, TervisMES, TervisPaylocity

function Install-PaylocityToActiveDirectory {
    param (
        $PathToScriptForScheduledTask = $PSScriptRoot,
        [Parameter(Mandatory)]$ScheduledTaskUserPassword,
        [Parameter(Mandatory)]$PathToPaylocityDataExport,
        [Parameter(Mandatory)]$PaylocityDepartmentsWithNiceNamesJsonPath
    )
    
    Install-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask `
        -ScheduledTaskUserPassword $ScheduledTaskUserPassword `
        -ScheduledTaskFunctionName "Send-EmailRequestingPaylocityReportBeRun" `
        -RepetitionInterval OnceAWeekMondayMorning

    Install-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask `
        -ScheduledTaskUserPassword $ScheduledTaskUserPassword `
        -ScheduledTaskFunctionName "Invoke-PaylocityToActiveDirectory" `
        -RepetitionInterval OnceAWeekTuesdayMorning

    Install-TervisPaylocity -PathToPaylocityDataExport $PathToPaylocityDataExport -PaylocityDepartmentsWithNiceNamesJsonPath $PaylocityDepartmentsWithNiceNamesJsonPath
}

function Uninstall-PaylocityToActiveDirectory {
    param (
        $PathToScriptForScheduledTask = $PSScriptRoot
    )
    Uninstall-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask -ScheduledTaskFunctionName "Invoke-PaylocityToActiveDirectory"
}

Function Invoke-DeployPaylocityToActiveDirectory {
    param (
        $ComputerName,
        [Parameter(Mandatory)]$PathToPaylocityDataExport,
        [Parameter(Mandatory)]$PaylocityDepartmentsWithNiceNamesJsonPath
    )

    $Credential = Get-PasswordstateCredential -PasswordID 259
    $ScheduledTaskUserPassword = $Credential.GetNetworkCredential().password

    $Session = New-PSSession -ComputerName $ComputerName -Credential $Credential
    Invoke-Command -Session $Session -ScriptBlock {
        Set-Location ($ENV:PSModulepath -split ";")[0]
        
        if ((Get-Command "git.exe" -ErrorAction SilentlyContinue) -eq $null) { 
            choco install git -y
        }
        
        "PaylocityToActiveDirectory","PowerShellApplication", "TervisMailMessage", "PasswordStatePowerShell", "StringPowerShell", "TervisMES" | % {
            Git clone "https://github.com/Tervis-Tumbler/$_"
        }

        "PaylocityToActiveDirectory","PowerShellApplication", "TervisMailMessage", "PasswordStatePowerShell", "StringPowerShell", "TervisMES" | % {
            Write-host $_
            Push-Location -Path ".\$_"
            git pull
            Pop-Location
        }

        Install-PaylocityToActiveDirectory -ScheduledTaskUserPassword $ScheduledTaskUserPassword
    }
    
}

Function Get-PaylocityADUser {
    param(
        [ValidateSet("A","T")]$Status
    )
    $PaylocityRecords = Get-PaylocityEmployees @PSBoundParameters
    $ADUsers = Get-ADUser -Properties EmployeeID,MemberOf -Filter *
    $PaylocityADUsers = $ADUsers | where EmployeeID -In $PaylocityRecords.EmployeeID
    
    $PaylocityADUsers | % {
        $_ |
        Add-Member -Name PaylocityDepartmentCode -MemberType NoteProperty -PassThru -Force -Value (
            $PaylocityRecords |
            where EmployeeID -eq $_.EmployeeID |
            select -ExpandProperty DepartmentCode
        ) |
        Add-Member -Name PaylocityDepartmentName -MemberType NoteProperty -PassThru -Force -Value (
            $PaylocityRecords |
            where EmployeeID -eq $_.EmployeeID |
            select -ExpandProperty DepartmentName
        ) |
        Add-Member -Name PaylocityDepartmentNiceName -MemberType ScriptProperty -PassThru -Force -Value {
            Get-DepartmentNiceName -PaylocityDepartmentName $this.PaylocityDepartmentName 
        } |
        Add-Member -Name PaylocityDepartmentRoleSAMAccountName -MemberType ScriptProperty -PassThru -Force -Value {
            "Role_Paylocity$($this.PaylocityDepartmentCode)"
        } |
        Add-Member -Name PaylocityDepartmentRoleName -MemberType ScriptProperty -PassThru -Force -Value {
            "Role_Paylocity$($this.PaylocityDepartmentName)"
        }
    }
}

function Get-PaylocityEmployeesWithoutADAccount {
    param(
        [ValidateSet("A","T")]$Status
    )

    $PaylocityRecords = Get-PaylocityEmployees @PSBoundParameters
    $ADUsers = Get-ADUser -Properties employeeid -Filter *
    $PaylocityRecordsWithoutADUserAccount = $PaylocityRecords | where EmployeeID -NotIn $ADUsers.employeeid
    $PaylocityRecordsWithoutADUserAccount
}

function Get-ADUserWithPaylocityEmployeeRecord {
    param(
        [ValidateSet("A","T")]$Status
    )

    $PaylocityRecords = Get-PaylocityEmployees @PSBoundParameters
    $ADUsers = Get-ADUser -Properties employeeid -Filter *
    $ADUsers | where EmployeeID -In $PaylocityRecords.EmployeeID
}

Function Get-PaylocityEmployeesWithoutADAccountThatShouldHaveAnAccount {
    param(
        [ValidateSet("A","T")]$Status
    )

    $PaylocityRecordsWithoutADUserAccount = Get-PaylocityEmployeesWithoutADAccount @PSBoundParameters
    $StoreEmployeesWhoDontGetADAccounts = Get-StoreEmployeesWhoDontGetADAccounts
    
    $PaylocityEmployeesWithoutADAccountThatShouldHaveAnAccount = $PaylocityRecordsWithoutADUserAccount | 
    where status -EQ A |
    where EmployeeID -NotIn $StoreEmployeesWhoDontGetADAccounts.EmployeeId |
    Where { "Ann Donelly" -ne "$($_.GivenName) $($_.Surname)"}

    $PaylocityEmployeesWithoutADAccountThatShouldHaveAnAccount 
}

Function Get-StoreEmployeesWhoDontGetADAccounts {
    Get-PaylocityEmployees |
    Where DepartmentName -EQ "Stores" |
    Where JobTitle -in "Sales Associate","Key Holder","Assistant Store Manager I","Stock Clerk"
}

Function Get-PaylocityTerminatedEmployeeStillEnabledInActiveDirectory {
    $PaylocityTerminatedEmployee = Get-PaylocityEmployees -Status T
    $ADUsers = Get-ADUser -Properties Employeeid, Title, MemberOf -Filter {Enabled -eq $true} | where {-not ($_.MemberOf -Match "CN=Contractor,")}
    $PaylocityTerminatedEmployeeStillEnabledInActiveDirectory = $PaylocityTerminatedEmployee | where EmployeeID -In $ADUsers.employeeid
    $PaylocityTerminatedEmployeeStillEnabledInActiveDirectory
}

Function Get-ActiveDirectoryUsersWithoutEmployeeIDThatShouldHaveEmployeeID {
    $DepartmentsOU = Get-ADOrganizationalUnit -Filter * | where name -Match "Departments"
    $ADUsersWithoutEmployeeID = Get-ADUser -SearchBase $DepartmentsOU.DistinguishedName -Filter * -Properties EmployeeID, Manager, Department, LastLogonDate, MemberOf, Created | 
    where {-not $_.EmployeeId} |
    where DistinguishedName -NotMatch "OU=Store Accounts,OU=Users,OU=Stores,OU=Departments" |
    where {-not ($_.MemberOf -Match "CN=Contractor,")} |
    where {-not ($_.MemberOf -Match "CN=SharedAccountsThatNeedToBeAddressed,")} |
    where {-not ($_.MemberOf -Match "CN=Test Users,")}
    $ADUsersWithoutEmployeeID | sort name | select -Property * -ExcludeProperty MemberOf, PropertyNames
}

Function Invoke-ReviewActiveDirectoryUsersWithoutEmployeeIDThatShouldHaveEmployeeID {
    Get-ActiveDirectoryUsersWithoutEmployeeIDThatShouldHaveEmployeeID | 
    select -Property * -ExcludeProperty DistinguishedName,ObjectClass,ObjectGUID,EmployeeID,PSShowComputerName,SID | ft
}

function Invoke-PaylocityToActiveDirectory {
    Import-PaylocityOrganizationStructureIntoActiveDirectory
    Set-ADUserDepartmentBasedOnPaylocityDepartment
    Invoke-EnsurePaylocityDepartmentsHaveRole
    Invoke-PaylocityDepartmentMemberShipToRoleSync
}

function Import-PaylocityOrganizationStructureIntoActiveDirectory {
    [CmdletBinding()]
    param ()
    $PaylocityRecords = Get-PaylocityEmployees -Status A
    $ADUsers = Get-ADUser -Properties EmployeeID,Manager -Filter *
    $PaylocityRecordsWithADUserAccount = $PaylocityRecords | where EmployeeID -In $ADUsers.employeeid

    foreach ($PaylocityRecord in $PaylocityRecordsWithADUserAccount) {
        $EmployeeADUser = $ADUsers | where employeeid -EQ $PaylocityRecord.EmployeeID
        if ($PaylocityRecord.ManagerEmployeeID) {
            $ManagerADUser = $ADUsers | where employeeid -EQ $PaylocityRecord.ManagerEmployeeID
            if ($ManagerADUser -and $EmployeeADUser.Manager -ne $ManagerADUser.DistinguishedName) {
                Write-Verbose "Employee $($EmployeeADUser.samaccountname)"
                Write-Verbose "Manager $($ManagerADUser.SamAccountName)"
                set-aduser $EmployeeADUser.samaccountname -Manager $ManagerADUser
            }
        }
    }
}

Function Get-ADUsersWithGivenNamesThatDontMatchPaylocity {
    $PaylocityRecords = Get-PaylocityEmployees
    $ADUsers = Get-ADUser -Properties employeeid -Filter *
    $PaylocityRecordsWithADUserAccount = $PaylocityRecords | where EmployeeID -In $ADUsers.employeeid

    foreach ($PaylocityRecord in $PaylocityRecordsWithADUserAccount) {
        $EmployeeADUser = $ADUsers | where employeeid -EQ $PaylocityRecord.EmployeeID

        if ($PaylocityRecord.EmployeeGivenName -ne $EmployeeADUser.GivenName) {
            "$($PaylocityRecord.EmployeeGivenName) $($PaylocityRecord.EmployeeSurname) $($EmployeeADUser.name)"         
        }
    }
}

Function Invoke-MatchPaylocityEmployeeWithADUser {
    [CmdletBinding()]
    param(
        [Switch]$IncludeMatchesOnOnlySurname,
        [Switch]$OnylActiveEmployees
    )
    
    if ($OnylActiveEmployees) {
        $PaylocityEmployeesWithoutADAccountThatShouldHaveAnAccount = Get-PaylocityEmployeesWithoutADAccountThatShouldHaveAnAccount -Status A
    } else {
        $PaylocityEmployeesWithoutADAccountThatShouldHaveAnAccount = Get-PaylocityEmployeesWithoutADAccount
    }

    foreach ($Employee in $PaylocityEmployeesWithoutADAccountThatShouldHaveAnAccount) {
        [string]$Surname = $Employee.Surname
        [string]$GivenName = $Employee.GivenName
        $ADUser = Get-aduser -Filter {Surname -eq $Surname -and GivenName -eq $GivenName -and Employeeid -notlike "*"} -Properties employeeid, Title, Department, Manager
        If ($ADUser -and -not $ADUser.count) {
            $Employee | Write-VerboseAdvanced -Verbose:$true
            $Aduser | Write-VerboseAdvanced -Verbose:$true
            $ADUser | Set-Aduser -EmployeeID $Employee.EmployeeID -Confirm
        } elseif ($ADUser -and $ADUser.count)  {
            $Employee | Write-VerboseAdvanced -Verbose:$true
            $SelectedADUser = $Aduser | Out-GridView -PassThru
            if ($SelectedADUser) { $SelectedADUser | Set-Aduser -EmployeeID $Employee.EmployeeID -Confirm }
        } elseif ($IncludeMatchesOnOnlySurname) {
            $ADUserMatchingSurname = Get-aduser -Filter {Surname -eq $Surname -and Employeeid -notlike "*"} -Properties employeeid, Title, Department, Manager
            if ($ADUserMatchingSurname)  {
                $Employee | Write-VerboseAdvanced -Verbose:$true
                $SelectedADUser = if($ADUserMatchingSurname.count) {
                    $ADUserMatchingSurname | Out-GridView -PassThru
                } else {
                    $ADUserMatchingSurname | Write-VerboseAdvanced -PassThrough -Verbose
                }
                if ($SelectedADUser) { $SelectedADUser | Set-Aduser -EmployeeID $Employee.EmployeeID -Confirm }
            }
        }
    }
}

function Backup-ActiveDirectoryUserData {
    $ActiveDirectoryUsersExport = Get-ADUser -Filter * -Properties Employeeid, Manager, Title
    $ActiveDirectoryUsersExport | ConvertTo-Json | Out-File ~\ActiveDirectoryBackup.json
}

function Import-PaylocityEmployeeTitleIntoActiveDirectory {
    $PaylocityRecords = Get-PaylocityEmployees
    $ADUsers = Get-ADUser -Properties Employeeid, Title -Filter *
    $PaylocityRecordsWithADUserAccount = $PaylocityRecords | where EmployeeID -In $ADUsers.employeeid

    foreach ($PaylocityRecord in $PaylocityRecordsWithADUserAccount) {
        $EmployeeADUser = $ADUsers | where employeeid -EQ $PaylocityRecord.EmployeeID
        #Unfinished
    }    
}

Function Remove-PaylocityTerminatedProductionEmployeeStillInActiveDirectory {
    $PaylocityTerminatedEmployee = Get-PaylocityEmployees -Status T

    foreach ($Employee in $PaylocityTerminatedEmployee) {
        [string]$Surname = $Employee.Surname
        [string]$GivenName = $Employee.GivenName

        $ADUser = Get-ADUser -Filter {Enabled -eq $false -and Surname -eq $Surname -and GivenName -eq $GivenName -and Employeeid -notlike "*"} -Properties employeeid, Title, Department, Manager -SearchBase "OU=Users,OU=Production Floor,OU=Operations,OU=Departments,DC=tervis,DC=prv"
        if ($ADUser -and -not $ADUser.count) {
            $Employee | Write-VerboseAdvanced -Verbose:$true
            $Aduser | Write-VerboseAdvanced -Verbose:$true
            $ADUser | Remove-ADUser -Confirm
        }
    }
}

function Set-ADUserDepartmentBasedOnPaylocityDepartment {
    $ADUsersWithEmployeeIDs = Get-ADUser -Filter {Employeeid -like "*"} -Properties Employeeid, Department, Division
    $PaylocityRecords = Get-PaylocityEmployees
    foreach ($ADUser in $ADUsersWithEmployeeIDs) {
        $PaylocityRecord = $PaylocityRecords | where EmployeeID -EQ $ADUser.EmployeeID
        $DepartmentNiceName = Get-DepartmentNiceName -PaylocityDepartmentName $PaylocityRecord.DepartmentName
        if (-not $DepartmentNiceName) { Throw "No DepartmentNiceName returned by Get-DepartmentNiceName" }
        if ($ADUser.Department -ne $DepartmentNiceName) {
            $ADUser | Set-ADUser -Department $DepartmentNiceName -Division $ADUser.Department
        }
    }
}

Function New-WorkOrderToTerminatePaylocityEmployeeInTerminatedStatusButActiveInActiveDirectory {
    $PaylocityTerminatedEmployeeStillEnabledInActiveDirectory = Get-PaylocityTerminatedEmployeeStillEnabledInActiveDirectory
    foreach ($Employee in $PaylocityTerminatedEmployeeStillEnabledInActiveDirectory) {
        $Employee | Get-ADUserByEmployeeID | Disable-ADAccount
        New-KanbanizeTask -Title "EmployeeID $($Employee.EmployeeID) Name $($Employee.GivenName) $($Employee.SurName), terminated in Paylocity but not AD" -BoardID 29 -Column "Requested"
    }
}

Function Invoke-EnsurePaylocityDepartmentsHaveRole {
    $PaylocityDepartmentNamesAndCodes = Get-PaylocityDepartmentNamesAndCodes

    ForEach ($PaylocityDepartmentNameAndCode in $PaylocityDepartmentNamesAndCodes) {
        $DepartmentNiceName = Get-DepartmentNiceName -PaylocityDepartmentName $PaylocityDepartmentNameAndCode.DepartmentName
        $ExistingADGroupForPaylocityDepartment = Get-ADGroup -Identity "Role_Paylocity$($PaylocityDepartmentNameAndCode.DepartmentCode)" -Properties Description
        
        if ($ExistingADGroupForPaylocityDepartment) {
            if ($ExistingADGroupForPaylocityDepartment.Name -ne "Role_Paylocity$($PaylocityDepartmentNameAndCode.DepartmentName)" ) {
               $ExistingADGroupForPaylocityDepartment | Rename-ADObject -NewName "Role_Paylocity$($PaylocityDepartmentNameAndCode.DepartmentName)"
            }

            if ($ExistingADGroupForPaylocityDepartment.Description -ne "Role Paylocity $DepartmentNiceName") {
                $ExistingADGroupForPaylocityDepartment | Set-ADGroup -Description "Role Paylocity $DepartmentNiceName"
            }
        } else {
            New-ADGroup -Path "OU=Paylocity,OU=Company - Security Groups,DC=tervis,DC=prv" -Name "Role_Paylocity$($PaylocityDepartmentNameAndCode.DepartmentName)" -Description "Role Paylocity $DepartmentNiceName" -GroupCategory Security -GroupScope Universal -SamAccountName "Role_Paylocity$($PaylocityDepartmentNameAndCode.DepartmentCode)"
        }
    }
}

Function Invoke-PaylocityDepartmentMemberShipToRoleSync {
    $ADUsers = Get-PaylocityADUser -Status A | 
    where {-not ($_.MemberOf -Match $_.PaylocityDepartmentRoleName) }

    foreach ($ADUser in $ADUsers) {
        Add-ADGroupMember -Identity $ADUser.PaylocityDepartmentRoleSAMAccountName -Members $ADUser
    }
}

Function Send-EmailRequestingPaylocityReportBeRun {
    $HTMLBody = @"
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
w\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style><![endif]--><style><!--
/* Font Definitions */
@font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
	{font-family:Verdana;
	panose-1:2 11 6 4 3 5 4 4 2 4;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0in;
	margin-bottom:.0001pt;
	font-size:10.0pt;
	font-family:"Verdana",sans-serif;}
a:link, span.MsoHyperlink
	{mso-style-priority:99;
	color:#0563C1;
	text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
	{mso-style-priority:99;
	color:#954F72;
	text-decoration:underline;}
span.EmailStyle17
	{mso-style-type:personal-compose;
	font-family:"Verdana",sans-serif;
	color:windowtext;
	font-weight:normal;
	font-style:normal;
	text-decoration:none none;}
.MsoChpDefault
	{mso-style-type:export-only;
	font-size:10.0pt;
	font-family:"Verdana",sans-serif;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection1
	{page:WordSection1;}
--></style><!--[if gte mso 9]><xml>
<o:shapedefaults v:ext="edit" spidmax="1026" />
</xml><![endif]--><!--[if gte mso 9]><xml>
<o:shapelayout v:ext="edit">
<o:idmap v:ext="edit" data="1" />
</o:shapelayout></xml><![endif]--></head><body lang=EN-US link="#0563C1" vlink="#954F72"><div class=WordSection1><p class=MsoNormal>Alicia,<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Can you please run the paylocity report we used before and save the results as xml into <a href="file://tervis.prv/departments/HR/HR/Paylocity%20Data%20Export">this folder</a>?<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Thanks,<o:p></o:p></p><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='margin-left:5.4pt;border-collapse:collapse'><tr style='height:102.65pt'><td width=205 valign=top style='width:153.9pt;padding:0in 5.4pt 0in 5.4pt;height:102.65pt'><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=485 style='width:363.65pt;border-collapse:collapse'><tr style='height:29.95pt'><td width=447 valign=top style='width:335.45pt;padding:0in 0in 0in 0in;height:29.95pt'><p class=MsoNormal style='line-height:115%'><span style='font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal style='line-height:115%'>Chris Magnuson<o:p></o:p></p><p class=MsoNormal style='line-height:115%'>Technical Services Manager<o:p></o:p></p><p class=MsoNormal style='line-height:115%'>d: 941.441.4491<o:p></o:p></p><p class=MsoNormal style='line-height:115%'><span style='font-family:"Calibri",sans-serif'><img border=0 id="Picture_x0020_25" src="https://sharepoint.tervis.com/SiteCollectionImages/NEW_Logo.jpg" alt="Tervis_Color_Logo_URL"><o:p></o:p></span></p><p class=MsoNormal style='margin-left:4.5pt;line-height:115%'><span style='font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p></td><td width=38 valign=top style='width:28.2pt;padding:0in 5.4pt 0in 5.4pt;height:29.95pt'><p class=MsoNormal align=center style='margin-left:-23.4pt;text-align:center;line-height:115%'><span style='font-size:11.0pt;line-height:115%;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p></td></tr></table></td><td width=240 valign=top style='width:2.5in;padding:0in 5.4pt 0in 5.4pt;height:102.65pt'><p class=MsoNormal align=center style='margin-left:-23.4pt;text-align:center;line-height:115%'><span style='font-size:11.0pt;line-height:115%'><o:p>&nbsp;</o:p></span></p></td></tr></table><p class=MsoNormal><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>
"@

    Send-TervisMailMessage -To alarkins@tervis.com -Bcc cmagnuson@tervis.com -Subject "Export of data from Paylocity" -Body $HTMLBody -From "Chris Magnuson <cmagnuson@tervis.com>" -BodyAsHtml
}

function Test-ADUsersWithDuplicateEmployeeIDs {
    $ADUsersWithEmployeeIDs = Get-ADUser -Filter {Employeeid -like "*"} -Properties Employeeid
    $ADUsersWithEmployeeIDs | group employeeid | where count -GT 1
}

function Test-DuplicateEmployeeID {
    $PaylocityRecords = Get-PaylocityEmployees
    $ADUsers = Get-ADUser -Properties employeeid -Filter *

    $PaylocityRecords | Group-Object EmployeeID | where count -gt 1
    $ADUsers | Group-Object employeeid | where count -gt 1 | where name -NE "" | select -ExpandProperty group
}
