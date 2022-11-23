#cleans terminal
Clear-Host #

<#----- Functions -----#>

# Creates csv file with headers if doesn't exists
# First argument - fileName (witho extension), Secound argument - header values
function createCsvWithHeader($fileName, $headers) {
  $currPath = "$($dirPath)\$($fileName)"
  if(-not(Test-Path -Path $currPath)) {
    Set-Content $currPath -Value $headers
    Write-Host "Created directory: $($currPath)" -ForegroundColor Green
  }
}

# Adds content to existing csv file 
# First argument - fileName (with extension), Secound argument - header values
function addToCsv($fileName, $headers) {
   $headers | Add-Content -Path "$($dirPath)\$($fileName)"
   Write-Host "Added new data to csv: $($fileName)" -ForegroundColor Green
}

# Gets domain name
function getDomainName {
    return Get-WMIObject Win32_ComputerSystem | Select-Object -ExpandProperty Domain
}

# Generates random password
function randomPass {
    $newPass = ''
    1..12 | ForEach-Object {
        $newPass += [char](Get-Random -Minimum 48 -Maximum 122)
}
    return $newPass
}

# Creates directory path if doesn't exists
function verifyAndCreateDirPath($path) {
    if(Test-Path -Path $path) {
        Write-Host "Direcotry path exists" -ForegroundColor Red
    } else {
        New-Item $path -ItemType Directory -Force
        Write-Host "Created new direcotry path:  $($path)" -ForegroundColor Green
    }
}

# Gets nescesary user info to create user account in createNewUser function
# and sends logs to 18838_create_user.csv file 
function readUserData
{
  param
  (
    [string]
    [Parameter(Mandatory=$true,HelpMessage='Type first name')]
  	$name,

	[string]
    [Parameter(Mandatory=$true,HelpMessage='Type last name')]
  	$surName, 

    [string]
    [Parameter(Mandatory=$true,HelpMessage='Type your department')]
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
  addToCsv "18838_create_user.csv" "$($creator)|$($creationTime)|$($name).$($surName)"
} 

# Creates new user by data from readUserData function
# Adds login and password to csv
function createNewUser($name, $surName, $department) {
    $readPass = randomPass
    $password = ConvertTo-SecureString $readPass -AsPlainText -Force

    New-ADUser `
        -UserPrincipalName "$($name).$($surName)" `
        -SamAccountName "$($name).$($surName)" `
        -EmailAddress "$($name).$($surName)@$($domainName)" `
        -DisplayName "$($name) $($surName)" `
        -Name "$($name).$($surName)" `
        -AccountPassword $password `
        -Department $department `
        -Enabled $true `
        -Path "OU=$($ou),$($domainNameDN)"
    addToCsv "18838_nazwa_uzytkownika.csv" "$($name).$($surName)|$($readPass)"
}

# Creates and adds users to AD from csv file
function createUsersFromCsv {
  $csvUsers = Import-Csv "$($dirPath)\$($usersCsvName).csv" -Delimiter "|"    
  $csvUsers | ForEach-Object {
    readUserData $_.imie $_.nazwisko $._dzial
  }
}

# Disables user account by login
function disableADAccount {
  $accountToDisable = Read-Host "Type AD account login to disable"
  Disable-ADAccount -Identity $accountToDisable
  Write-Host "Account disabled $($accountToDisable)" -ForegroundColor Green

  $creationTime = (Get-ADUser -Filter "EmailAddress -eq '$($accountToDisable)@$($domainName)'" -Properties whenCreated).whenCreated
  addToCsv "18838_wylaczone_konta_data.csv" "$($creator)|$($creationTime)|$($accountToDisable)@TCO18838.pl"
}

# Changes password for user in domain.
function changeUserPassword {
  $accountToChangePass = Read-Host "Type AD account login to change his password"
  $newPass = Read-Host "Type new password"

  Set-ADAccountPassword -Identity "$($accountToChangePass)" -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$($newPass)" -Force)

  Write-Host "Password changed for account: $($accountToChangePass)" -ForegroundColor Green
  $creationTime = (Get-ADUser -Filter "EmailAddress -eq '$($accountToChangePass)@$($domainName)'" -Properties whenCreated).whenCreated
  addToCsv "18838_zmiana_hasla_data.csv" "$($creator)|$($creationTime)|$($accountToChangePass)@TCO18838.pl"
}

#Create OU
function addNewOU {
  $ouCheck = Get-ADOrganizationalUnit -Filter "distinguishedName -eq 'OU=$($ou), $($domainNameDN)'"
  
  if(-not($ouCheck)) {
     New-ADOrganizationalUnit -Name $ou -Path $($domainNameDN) -ProtectedFromAccidentalDeletion $false
      Write-Host "Added new OU: $($ou)" -ForegroundColor Green
  }
} 

# Creates groups
function addNewGroup {
    $groupName = Read-Host "Type name for new group:"
    $newOU = "OU=$($ou),$($domainNameDN)"
    New-ADGroup -Name "$($groupName)" -SamAccountName "$($groupName)" -DisplayName "$($groupName)" `
    -GroupCategory Security -GroupScope Global -Path $newOU
    Write-Host "New grup created: $($groupName)" -ForegroundColor Green
    
    $creationTime = (Get-ADGroup -Filter "SamAccountName -eq 'test'" -Properties whenCreated).whenCreated
    addToCsv "18838_create_group.csv" "$($creator)|$($creationTime)|$($groupName)"
}

# Adds new user to specific group by user login
function addGroupMember {
  $group = Read-Host "Type the group to which you want to add the user: "
  $member = Read-Host "Type the user login you want to add to the group:"
  $userStatment = Get-ADGroupMember -Identity $group | Where-Object {$_.name -eq $member}
  if(-not($userStatment)){
    Add-ADGroupMember -Identity $group -Members $member  
    Write-Host "New user $($member) added to group $($group)" -ForegroundColor Green
    addToCsv "18838 zmiana członkostwa grup.txt"$($creator)"|$($member)|$($group)"
  } else {
    Write-Host "User $($member) exists in group $($group)" -ForegroundColor Red
  }
}

<#----- Generate reports -----#>

# Generates list into .txt file with all members in each group
function reportGroupAccounts {
  (Get-ADGroup -Filter * -Properties name | Select-Object name).name | ForEach-Object {
    createCsvWithHeader "$($index) $($_).txt" "login"
    $currentGroup = $_
    (Get-ADGroupMember -Identity $_).name | ForEach-Object {
      addToCsv "$($index) $($currentGroup).txt" $_
    }
  }
}

# Generates specified data into .csv file with all disabled accounts
function  reportDisabledAccounts {
  Get-ADUser -Filter {(Enabled -eq $False)}  -Properties SamAccountName, DistinguishedName, SID, modifyTimeStamp | `
  Select-Object SamAccountName, DistinguishedName, SID, modifyTimeStamp | ForEach-Object {
    addToCsv "18838 wyłączone konta.csv" "$($_.SamAccountName)|$($_.distinguishedName)|$($_.SID)|$($_.modifyTimeStamp)"
  } 
}
reportDisabledAccounts
<#----- Variables -----#>

$domainName = getDomainName
$domainNameDN = (Get-ADDomain).DistinguishedName

$index = "18838"
$ou = $index
$dirPath = "C:\wit\18838"
$usersCsvName = "Użytkownicy" # +Później read-host i do funkcji menu
$creator = $env:UserName


<#----- Launch function -----#>

#Creates directory path for csv files
verifyAndCreateDirPath $dirPath

# Creates once all nescesary csv files for log and data
createCsvWithHeader "18838 nazwa uzytkownika" "login|haslo.csv"
createCsvWithHeader "18838_create_user" "autor|data utworzenia|nazwa użytkownika.csv"
createCsvWithHeader "18838 wylaczone konta data" "autor|data utworzenia|nazwa użytkownika.csv"
createCsvWithHeader "18838 zmiana hasla data" "autor|data utworzenia|nazwa użytkownika.csv"
createCsvWithHeader "18838 create group" "autor grupy|data utworzenia|nazwa grupy.csv"
createCsvWithHeader "18838 zmiana członkostwa grup.txt" "autor|nazwa użytkownika|grupa"
createCsvWithHeader "18838 wyłączone konta.csv" "Nazwa konta|DistinguishedName|SID|Data ostatniej modyfikacji"

#Creates once initial OU to keep all object
addNewOU

# Initializes user data reading
#readUserData