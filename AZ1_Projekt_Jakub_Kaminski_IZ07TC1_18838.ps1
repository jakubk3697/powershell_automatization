<#----- Functions -----#>

# Create a csv file with headers if it doesn't exist
# First argument - file name with extension, second argument - header values
function createFileWithHeader($fileName, $headers) {
  $currPath = "$($dirPath)\$($fileName)"
  if(-not(Test-Path -Path $currPath)) {
    Set-Content $currPath -Value $headers
    Write-Host "New file created: $($currPath)" -ForegroundColor Green
  }
}

# Adds content to an existing csv file
# First argument - file name with extension, second argument - header values
function addToFile($fileName, $headers) {
   $headers | Add-Content -Path "$($dirPath)\$($fileName)"
   $line = $_.InvocationInfo.ScriptLineNumber
   Write-Host "Data added to the file: $($fileName)" $line -ForegroundColor Green 
}

# Get the domain name
function getDomainName {
    return Get-WMIObject Win32_ComputerSystem | Select-Object -ExpandProperty Domain
}

# Generates a random password
function randomPass {
    $newPass = ''
    1..12 | ForEach-Object {
        $newPass += [char](Get-Random -Minimum 48 -Maximum 122)
}
    return $newPass
}

# Create directory path if it doesn't exist
function verifyAndCreateDirPath($path) {
    if(Test-Path -Path $path) {
        Write-Host "The specified directory path exists" -ForegroundColor Red
    } else {
        New-Item $path -ItemType Directory -Force
        Write-Host "A new directory path has been created:  $($path)" -ForegroundColor Green
    }
}

# Fetches the necessary user information to create a user account in the createNewUser function
# and sends logs to file create_user.csv
function readUserData
{
  param
  (
    [string]
    [Parameter(Mandatory=$true,HelpMessage='Type name')]
  	$name,

	[string]
    [Parameter(Mandatory=$true,HelpMessage='Type surname')]
  	$surName, 

    [string]
    [Parameter(Mandatory=$true,HelpMessage='Type department')]
  	$department
  )
  
  $usersAmount = (Get-ADUser -Filter * | measure).Count
  $currentUserEmail = Get-ADUser -Filter "EmailAddress -eq '$($name).$($surName)@$($domainName)'"
  if($currentUserEmail) {
    $surName = "$($surName)$($usersAmount)"
  }

  createNewUser $name $surName $department
  Write-Host "Created user $($name).$($surName)@$($domainName)" -ForegroundColor Green

  $creationTime = (Get-ADUser -Filter "EmailAddress -eq '$($name).$($surName)@$($domainName)'" -Properties whenCreated).whenCreated
  addToFile "create_user.csv" "$($creator)|$($creationTime)|$($name).$($surName)"
} 

#Creates a new user based on the data from the readUserData function
# Adds login and password to csv file
function createNewUser($name, $surName, $department) {
    $readPass = randomPass
    $password = ConvertTo-SecureString $readPass -AsPlainText -Force

    New-ADUser `
        -UserPrincipalName "$($name).$($surName)" `
        -SamAccountName "$($name).$($surName)" `
        -EmailAddress "$($name).$($surName)@$($domainName)" `
        -DisplayName "$($name) $($surName)" `
        -Name "$($name).$($surName)" `
        -GivenName "$($name)" `
        -Surname "$($surName)" `
        -AccountPassword $password `
        -Department $department `
        -Enabled $true `
        -Path "OU=$($ou),$($domainNameDN)"
    addToFile "user_name.csv" "$($name).$($surName)|$($readPass)"
}

# Create and add users to AD from a csv file provided by the user
function createUsersFromCsv {
  $usersCsvName = Read-Host "Type the name of the file containing user data with the heading: 'name|surname|department' or use the file 'users_to_add.csv'"
  $csvUsers = Import-Csv "$($dirPath)\$($usersCsvName)" -Delimiter "|"    
  $csvUsers | ForEach-Object {
    Write-host "Created user: $($_.name) $($_.surName) $($_.department)" -ForegroundColor Green
    readUserData $_.name $_.surName $.department
  }
}

# Disables a particular account based on given login
function disableADAccount {
  $accountToDisable = Read-Host "Type the username to disable the account (e.g. john.doe))"
  Disable-ADAccount -Identity $accountToDisable
  Write-Host "Disabled account: $($accountToDisable)" -ForegroundColor Green

  $creationTime = (Get-ADUser -Filter "EmailAddress -eq '$($accountToDisable)@$($domainName)'" -Properties whenCreated).whenCreated
  addToFile "disabled_accounts_data.csv" "$($creator)|$($creationTime)|$($accountToDisable)@c"
}

# Changes the password for a user in the domain based on the given login
function changeUserPassword {
  $accountToChangePass = Read-Host "Enter the user's login to change his password (e.g. john.doe)"
  $newPass = Read-Host "Type new password"

  Set-ADAccountPassword -Identity "$($accountToChangePass)" -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$($newPass)" -Force)

  Write-Host "Account password changed: $($accountToChangePass)" -ForegroundColor Green
  $creationTime = (Get-ADUser -Filter "EmailAddress -eq '$($accountToChangePass)@$($domainName)'" -Properties whenCreated).whenCreated
  addToFile "password_change_data.csv" "$($creator)|$($creationTime)|$($accountToChangePass)@$($domainName)"
}

# Creates a new OU based on a variable
function addNewOU {
  $ouCheck = Get-ADOrganizationalUnit -Filter "distinguishedName -eq 'OU=$($ou), $($domainNameDN)'"
  
  if(-not($ouCheck)) {
     New-ADOrganizationalUnit -Name $ou -Path $($domainNameDN) -ProtectedFromAccidentalDeletion $false
      Write-Host "Created OU: $($ou)" -ForegroundColor Green
  }
} 

# Creates a group in a new OU (for easy management) based on user input
function addNewGroup {
    $groupName = Read-Host "Type a name for the new resource group:"
    $newOU = "OU=$($ou),$($domainNameDN)"
    New-ADGroup -Name "$($groupName)" -SamAccountName "$($groupName)" -DisplayName "$($groupName)" `
    -GroupCategory Security -GroupScope Global -Path $newOU
    Write-Host "Created new group: $($groupName)" -ForegroundColor Green
    
    $creationTime = (Get-ADGroup -Filter "SamAccountName -eq '$($groupName)'" -Properties whenCreated).whenCreated
    addToFile "create_group.csv" "$($creator)|$($creationTime)|$($groupName)"
}

# Adds a new user to the specified group based on the user's login
function addGroupMember {
  $group = Read-Host "Type the name of the group to which you want to add the user: "
  $member = Read-Host "Type the login of the user to be added to the selected group:"
  $userStatment = Get-ADGroupMember -Identity $group | Where-Object {$_.name -eq $member}
  if(-not($userStatment)){
    Add-ADGroupMember -Identity $group -Members $member  
    Write-Host "User '$($member)' added to group '$($group)'" -ForegroundColor Green
    addToFile "changing_group_membership.csv" "$($creator)|$($member)|$($group)"
  } else {
    Write-Host "User $($member) already exists in $($group)" -ForegroundColor Red
  }
}

<#----- Generate reports -----#>

# generates a list to a .csv file with all members of each group
function reportGroupAccounts {
  (Get-ADGroup -Filter * -Properties name | Select-Object name).name | ForEach-Object {
    createFileWithHeader "$($_).csv" "login"
    $currentGroup = $_
    (Get-ADGroupMember -Identity $_).name | ForEach-Object {
      addToFile "$($currentGroup).csv" $_
    }
  }
}


# Generates the specified data to a .csv file with all accounts disabled
function  reportDisabledAccounts {
  Get-ADUser -Filter {(Enabled -eq $False)}  -Properties SamAccountName, DistinguishedName, SID, modifyTimeStamp | `
  Select-Object SamAccountName, DistinguishedName, SID, modifyTimeStamp | ForEach-Object {
    addToFile "disabled_accounts.csv" "$($_.SamAccountName)|$($_.distinguishedName)|$($_.SID)|$($_.modifyTimeStamp)"
  } 
}

# Generates a report with the most important information about users in AD
function reportADUsersInfo {
  Get-ADUser -Filter * -Properties givenName, surName, userPrincipalName, samAccountName, distinguishedName, whenCreated, modifyTimeStamp, LastLogon, PasswordLastSet | `
  Select-Object givenName, surName, userPrincipalName, samAccountName, distinguishedName, whenCreated, modifyTimeStamp, LastLogon, PasswordLastSet | ForEach-Object {
    addToFile "users.csv" "$($_.givenName)|$($_.surName)|$($_.userPrincipalName)|$($_.samAccountName)|$($_.distinguishedName)|$($_.whenCreated)||$($_.modifyTimeStamp)|$($_.LastLogon)|$($_.PasswordLastSet)"
  } 
}

# Generates a report with information about all computer accounts in the domain
function reportADCoumputersInfo {
  Get-ADComputer -Filter * -Properties Name, SID, distinguishedName, Enabled, LastLogonDate, Created `
    | Select-Object Name, SID, distinguishedName, Enabled, LastLogonDate, Created | ForEach-Object {
      $os = (Get-ComputerInfo).windowsProductName
      $filePath = "$($domainName)_$($os).csv"
      createFileWithHeader "$($filePath)" "Computer name|SID|DistinguishedName|Account status|Last password change|Creation date"
      addToFile "$($filePath)" "$($_.Name)|$($_.SID)|$($_.distinguishedName)|$($_.Enabled)|$($_.LastLogon)|$($_.Created)"
  }
}

# Generates OU information and transfers it to a csv file
function reportOUInfo {
  Get-ADOrganizationalUnit -Filter * -Properties distinguishedName, name | Select-Object distinguishedName, name | Sort-Object distinguishedName `
  | ForEach-Object {
    addToFile "$($os).csv" "$($_.name)|$($_.distinguishedName)"
  }
}

# It shows the intro in the control panel
function showGreetings {
  Clear-Host
  Write-Host "____________________| Automation of active directory resources |____________________" -ForegroundColor Magenta
  Write-Host  "-------> Use the numeric keys as instructed to perform AD actions<------- " -ForegroundColor Green
  Write-Host "All created data and log files can be found in: $($dirPath)" -ForegroundColor Green 
}


<#----- Variables -----#>

$domainName = getDomainName
$domainNameDN = (Get-ADDomain).DistinguishedName

$ou = "myNewOU"
$dirPath = "C:\PS\generated"
$creator = $env:UserName
$os = $($(Get-ComputerInfo).windowsProductName)


<#----- Initial data initialization functions -----#>

# Create directory path for csv files
verifyAndCreateDirPath $dirPath

# Creates all needed csv files for logs and data
createFileWithHeader "users_to_add.csv" "name|surname|department"
createFileWithHeader "username.csv" "login|password"
createFileWithHeader "create_user.csv" "author|creation date|username"
createFileWithHeader "disabled_accounts_data.csv" "author|creation date|username"
createFileWithHeader "password_change_data.csv" "author|creation date|username"
createFileWithHeader "create_group.csv" "group author|creation date|group name"
createFileWithHeader "changing_group_membership.csv" "author|username|group"
createFileWithHeader "disabled_accounts.csv" "Account name|DistinguishedName|SID|last modification"
createFileWithHeader "users.csv" "name|surName|login(UPN)|samacount|localization in ADDS (DN)|creation date|last modification|last login|last password change"
createFileWithHeader "$($os).csv" "OU|DistinguishedName"

#Creates OU
addNewOU

<#----- Control panel -----#>
# Loading data and performing individual functions responsible for individual actions
do {
  showGreetings
  Write-Host "[1] Obsługa kont użytkowników"
  Write-Host "[2] Obsługa kont grup"
  Write-Host "[3] Generowanie raportów"

  $mainSelect = Read-Host "Wybierz opcję"
  switch ($mainSelect) {
    '1' { 
      Write-Host "Wybrałeś opcję [$($mainSelect)]" -ForegroundColor Green
      Write-Host "[1] Tworzenie konta użytkownika"
      Write-Host "[2] Tworzenie wielu kont na podstawie pliku csv"
      Write-Host "[3] Wyłączenie konta użytkownika"
      Write-Host "[4] Zmiana hasła konta uzytkownika"
      $accountsSelect = Read-Host "Wybierz opcje"
      switch ($accountsSelect) {
        '1' { readUserData }
        '2' { createUsersFromCsv }
        '3' { disableADAccount }
        '4' { changeUserPassword }
      }
     }
    '2' { 
      Write-Host "Wybrałeś opcję [$($mainSelect)]" -ForegroundColor Green
      Write-Host "[1] Utworzenie nowej grupy"
      Write-Host "[2] Dodanie nowego użytkownika do grupy"
      $groupsSelect = Read-Host "Wybierz opcje"
      switch ($groupsSelect) {
        '1' { addNewGroup }
        '2' { addGroupMember }
      }
     }
    '3' { 
      Write-Host "Wybrałeś opcję [$($mainSelect)]" -ForegroundColor Green
      Write-Host "[1] Generowanie list grup z członkami"
      Write-Host "[2] Generowanie listy wyłączonych kont w domenie"
      Write-Host "[3] Generowanie list szczegółowych informacji o kontach użytkowników"
      Write-Host "[4] Generowanie list szczegółowych informacji o kontach komputerów w domenie"
      Write-Host "[5] Generowanie listy jednostek organizacyjnych w domenie (alfabetycznie względem OU)"
      $reportsSelect = Read-Host "Wybierz opcję"
      switch ($reportsSelect) {
        '1'{ reportGroupAccounts }
        '2'{ reportDisabledAccounts }
        '3'{ reportADUsersInfo }
        '4'{ reportADCoumputersInfo }
        '5'{ reportOUInfo }
      }
     }
  }
  pause
}
until ($selection -eq 'q')