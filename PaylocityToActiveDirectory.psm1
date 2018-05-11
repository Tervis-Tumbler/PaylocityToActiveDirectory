function Install-PaylocityToActiveDirectory {
    param (
        $PathToScriptForScheduledTask = $PSScriptRoot,
        [Parameter(Mandatory)]$PathToPaylocityDataExport,
        [Parameter(Mandatory)]$PaylocityDepartmentsWithNiceNamesJsonPath,
        [Parameter(Mandatory)]$ComputerName
    )
    $ScheduledTasksCredential = Get-PasswordstateCredential -PasswordID 259
    
    $InstallPowerShellApplicationParameters = @{
        ModuleName = "PaylocityToActiveDirectory"
        DependentTervisModuleNames = "TervisPaylocity", "TervisActiveDirectory"
        ScheduledScriptCommandsString = ""
    }

    Install-PowerShellApplication -ComputerName $ComputerName @InstallPowerShellApplicationParameters

    Install-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask `
        -Credential $ScheduledTasksCredential `
        -ScheduledTaskFunctionName "Invoke-PaylocityToActiveDirectory" `
        -RepetitionInterval OnceAWeekTuesdayMorning `
        -ComputerName $ComputerName

    Install-TervisPaylocity -PathToPaylocityDataExport $PathToPaylocityDataExport -PaylocityDepartmentsWithNiceNamesJsonPath $PaylocityDepartmentsWithNiceNamesJsonPath
}

function Uninstall-PaylocityToActiveDirectory {
    param (
        $PathToScriptForScheduledTask = $PSScriptRoot,
        [Parameter(Mandatory)]$ComputerName
    )
    Uninstall-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask -ComputerName $ComputerName -ScheduledTaskFunctionName "Invoke-PaylocityToActiveDirectory"
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
        
        "PaylocityToActiveDirectory","PowerShellApplication", "TervisPasswordStatePowerShell", "StringPowerShell", "TervisMES" | % {
            Git clone "https://github.com/Tervis-Tumbler/$_"
        }

        "PaylocityToActiveDirectory","PowerShellApplication",  "TervisPasswordStatePowerShell", "StringPowerShell", "TervisMES" | % {
            Write-host $_
            Push-Location -Path ".\$_"
            git pull
            Pop-Location
        }

        Install-PaylocityToActiveDirectory -ScheduledTaskUserPassword $ScheduledTaskUserPassword
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

Function New-WorkOrderToTerminatePaylocityEmployeeInTerminatedStatusButActiveInActiveDirectory {
    $PaylocityTerminatedEmployeeStillEnabledInActiveDirectory = Get-PaylocityTerminatedEmployeeStillEnabledInActiveDirectory
    foreach ($Employee in $PaylocityTerminatedEmployeeStillEnabledInActiveDirectory) {
        $Employee | Get-ADUserByEmployeeID | Disable-ADAccount
        New-KanbanizeTask -Title "EmployeeID $($Employee.EmployeeID) Name $($Employee.GivenName) $($Employee.SurName), terminated in Paylocity but not AD" -BoardID 29 -Column "Requested"
    }
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


function Sync-PaylocityPropertiesToActiveDirectory {
    [CmdletBinding()]
    param ()
    $ADUsers = Get-TervisADUser -Filter {Employeeid -like "*"} -IncludePaylocityEmployee -Properties Department,Division,Manager,MemberOf |
    Where-Object { $_.PaylocityEmployee }

    $ADUsers | Set-ADUserTitleBasedOnPaylocityEmployeeJobTitle
    $ADUsers | Set-ADUserDepartmentBasedOnPaylocityDepartment
    $ADUsers | Set-ADUserManagerBasedOnPaylocityManager -ADUsers $ADUsers

    Invoke-EnsurePaylocityDepartmentsHaveRole
    $ADUsers | Add-ADUserToPaylocityDepartmentRole
}

function Set-ADUserTitleBasedOnPaylocityEmployeeJobTitle {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ADUser
    )
    process {
        $PaylocityJobtitle = $ADUser.PaylocityEmployee.JobTitle
        if ($ADUser.Title -ne $PaylocityJobtitle) {
            Write-Verbose "Changing $($ADUser.Name) current title $($ADUser.Title) to $PaylocityJobtitle"
            $ADUser | Set-ADUser -Title $PaylocityJobtitle
        }
    }
}

function Set-ADUserDepartmentBasedOnPaylocityDepartment {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ADUser
    )
    process {
        $DepartmentNiceName = $ADUser.PaylocityEmployee.DepartmentNiceName
        if (-not $DepartmentNiceName) { Throw "No DepartmentNiceName returned by Get-DepartmentNiceName" }
        if ($ADUser.Department -ne $DepartmentNiceName) {
            Write-Verbose "Changing $($ADUser.Name) current department $($ADUser.Department) to $DepartmentNiceName"
            $ADUser | Set-ADUser -Department $DepartmentNiceName -Division $ADUser.Department
        }
    }
}

function Set-ADUserManagerBasedOnPaylocityManager {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ADUser,
        [Parameter(Mandatory)]$ADUsers
    )
    begin {
        $ADUsersEmployeeIDHash = @{}
        $ADUsers |
        ForEach-Object -Process {
            $ADUsersEmployeeIDHash.Add($_.EmployeeID, $_)
        }
    }
    process {
        $EmployeeADUser = $ADUser
        $ManagerEmployeeID = $EmployeeADUser.PaylocityEmployee.ManagerEmployeeID
        if ($ManagerEmployeeID) {            
            $ManagerADUser = $ADUsersEmployeeIDHash[$ManagerEmployeeID]
            if ($ManagerADUser -and $EmployeeADUser.Manager -ne $ManagerADUser.DistinguishedName) {
                Write-Verbose "$($EmployeeADUser.samaccountname) manager being set to $($ManagerADUser.SamAccountName)"
                $EmployeeADUser | Set-ADUser -Manager $ManagerADUser
            }
        }
    }
}

Function Invoke-EnsurePaylocityDepartmentsHaveRole {
    $PaylocityDepartments = Get-PaylocityDepartment

    ForEach ($PaylocityDepartment in $PaylocityDepartments) {
        $RoleDescription = "Role Paylocity $($PaylocityDepartment.NiceName)"

        $ADGroup = Try {
            Get-ADGroup -Identity $PaylocityDepartment.RoleSAMAccountName -Properties Description
        } catch {
            $ADOrganizationalUnit = Get-ADOrganizationalUnit -Filter { Name -eq "Paylocity" }
            New-ADGroup -Path $ADOrganizationalUnit -Name $PaylocityDepartment.RoleName -Description $RoleDescription -GroupCategory Security -GroupScope Universal -SamAccountName $PaylocityDepartment.RoleSAMAccountName
        }
        
        if ($ADGroup.Name -ne $PaylocityDepartment.RoleName ) {
            $ADGroup | Rename-ADObject -NewName $PaylocityDepartment.RoleName
        }
        
        if ($ADGroup.Description -ne $RoleDescription) {
            $ADGroup | Set-ADGroup -Description $RoleDescription
        }
    }
}

function Add-ADUserToPaylocityDepartmentRole {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ADUser
    )
    process {
        if (-not ($ADUser.MemberOf -Match $ADUser.PaylocityEmployee.DepartmentRoleName)) {
            Add-ADGroupMember -Identity $ADUser.PaylocityEmployee.DepartmentRoleSAMAccountName -Members $ADUser
        }
    }
}

