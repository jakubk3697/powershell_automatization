#cleans terminal
Clear-Host #

<#----- Functions -----#>

# Creates csv file with headers if doesn't exists
# First argument - fileName (without extension), Secound argument - header values
function createCsvWithHeader($fileName, $headers) {
  $currPath = "$($dirPath)\$($fileName).csv"
  if(-not(Test-Path -Path $currPath)) {
    Set-Content $currPath -Value $headers
    Write-Host "Created directory: $($currPath)" -ForegroundColor Green
  }
}

# Adds content to existing csv file 
# First argument - fileName (without extension), Secound argument - header values
function addToCsv($fileName, $headers) {
   $headers | Add-Content -Path "$($dirPath)\$($fileName).csv"
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
    [Parameter(Mandatory=$true,HelpMessage='Wpisz swoje imie')]
  	$name,

	[string]
    [Parameter(Mandatory=$true,HelpMessage='Wpisz swoje nazwisko')]
  	$surName, 

    [string]
    [Parameter(Mandatory=$true,HelpMessage='Wpisz nazwę swojego działu')]
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
  $creator = $env:UserName
  addToCsv "18838_create_user" "$($creator)|$($creationTime)|$($name).$($surName)"
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
    addToCsv "18838_nazwa_uzytkownika" "$($name).$($surName)|$($readPass)"
}

# Creates and adds users to AD from csv file
function createUsersFromCsv {
  $csvUsers = Import-Csv "$($dirPath)\$($usersCsvName)" -Delimiter "|"    
  $csvUsers | ForEach-Object {
    readUserData $_.imie $_.nazwisko $._dzial
  }
}

<#----- Variables -----#>

$domainName = getDomainName
$domainNameDN = (Get-ADDomain).DistinguishedName

$index = "18838"
$ou = $index
$dirPath = "C:\wit\18838"
$usersCsvName = "Użytkownicy.csv"


<#----- Launch function -----#>

#Creates directory path for csv files
verifyAndCreateDirPath $dirPath

# Creates all nescesary csv files for log and data
createCsvWithHeader "18838_nazwa_uzytkownika" "login|haslo"
createCsvWithHeader "18838_create_user" "twórca|data utworzenia|nazwa użytkownika"

# Initializes user data reading
readUserData