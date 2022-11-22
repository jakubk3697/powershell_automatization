clear #

<#----- Functions -----#>

# Creates csv file with headers
# First argument - header values, Secound argument - fileName (without extension)
function createCsvWithHeader($headers, $fileName) {
   Set-Content "$($dirPath)\$($fileName).csv" -Value $headers
}

# Adds content to existing csv file 
# First argument - header values, Secound argument - fileName (without extension)
function addToCsv($headers, $fileName) {
   $headers | Add-Content -Path "$($dirPath)\$($fileName).csv"
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
  Get-ADUser -Filter "EmailAddress -eq 'jan.kow@$($domainName)'"
  if(Get-ADUser -Filter "EmailAddress -eq '$($name).$($surName)@$($domainName)'") {
    $surName = "$($surName)$($usersAmount )"
  }

  createNewUser $name $surName $department
}  

# Creates new user by data from readUserData function
function createNewUser($name, $surName, $department) {
    $password = ConvertTo-SecureString randomPass -AsPlainText -Force

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
}

<#----- Variables -----#>

$domainName = getDomainName
$domainNameDN = (Get-ADDomain).DistinguishedName

$index = "18838"
$ou = $index
$dirPath = "C:\wit\18838"
$csvFilePath = "$($dirPath)\AZ3_Wersje_OS_$($index).csv"

verifyAndCreateDirPath
readUserData

<# Creates all nescesary csv files for log and data #>
Set-Content "$($dirPath)\18838_create_user.csv" -Value "login|haslo"