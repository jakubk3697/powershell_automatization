<#----- Funkcje -----#>

# Tworzy plik csv z nagłówkami, jeżeli nie istnieje
# Pierwszy argument - nazwa pliku z rozszerzeniem, drugi argument - wartości nagłówków
function createFileWithHeader($fileName, $headers) {
  $currPath = "$($dirPath)\$($fileName)"
  if(-not(Test-Path -Path $currPath)) {
    Set-Content $currPath -Value $headers
    Write-Host "Utworzono plik: $($currPath)" -ForegroundColor Green
  }
}

# Dodaje zawartość do istniejącego pliku csv 
# Pierwszy argument - nazwa pliku z rozszerzeniem, drugi argument - wartości nagłówków
function addToFile($fileName, $headers) {
   $headers | Add-Content -Path "$($dirPath)\$($fileName)"
   $line = $_.InvocationInfo.ScriptLineNumber
   Write-Host "Dodano dane do pliku: $($fileName)" $line -ForegroundColor Green 
}

# Uzyskuje nazwę domeny
function getDomainName {
    return Get-WMIObject Win32_ComputerSystem | Select-Object -ExpandProperty Domain
}

# Generuje losowe hasło
function randomPass {
    $newPass = ''
    1..12 | ForEach-Object {
        $newPass += [char](Get-Random -Minimum 48 -Maximum 122)
}
    return $newPass
}

# Tworzy ścieżkę do katalogu, jeśli nie istnieje
function verifyAndCreateDirPath($path) {
    if(Test-Path -Path $path) {
        Write-Host "Podana ścieżka katalogowa istnieje" -ForegroundColor Red
    } else {
        New-Item $path -ItemType Directory -Force
        Write-Host "Utworzono nową ścieżkę katalogową:  $($path)" -ForegroundColor Green
    }
}

# Pobiera niezbędne informacje o użytkowniku do utworzenia konta użytkownika w funkcji createNewUser
# i wysyła logi do pliku 18838_create_user.csv  
function readUserData
{
  param
  (
    [string]
    [Parameter(Mandatory=$true,HelpMessage='Wpisz imie')]
  	$name,

	[string]
    [Parameter(Mandatory=$true,HelpMessage='Wpisz nazwisko')]
  	$surName, 

    [string]
    [Parameter(Mandatory=$true,HelpMessage='Wpisz nazwę działu')]
  	$department
  )
  
  $usersAmount = (Get-ADUser -Filter * | measure).Count
  $currentUserEmail = Get-ADUser -Filter "EmailAddress -eq '$($name).$($surName)@$($domainName)'"
  if($currentUserEmail) {
    $surName = "$($surName)$($usersAmount)"
  }

  createNewUser $name $surName $department
  Write-Host "Utworzono użytkownika $($name).$($surName)@$($domainName)" -ForegroundColor Green

  $creationTime = (Get-ADUser -Filter "EmailAddress -eq '$($name).$($surName)@$($domainName)'" -Properties whenCreated).whenCreated
  addToFile "18838_create_user.csv" "$($creator)|$($creationTime)|$($name).$($surName)"
} 

#Tworzy nowego użytkownika na podstawie danych z funkcji readUserData
# Dodaje login i hasło do csv
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
    addToFile "18838_nazwa_uzytkownika.csv" "$($name).$($surName)|$($readPass)"
}

# Tworzenie i dodawanie użytkowników do AD z pliku csv podanego przez uzytkownika
function createUsersFromCsv {
  $usersCsvName = Read-Host "Wpisz nazwę pliku zawierającego dane użytkowników z nagłówkiem: 'imie|nazwisko|dzial' lub skorzystaj z pliku 'Uzytkownicy.csv'"
  $csvUsers = Import-Csv "$($dirPath)\$($usersCsvName)" -Delimiter "|"    
  $csvUsers | ForEach-Object {
    Write-host "Dodano użytkownika: $($_.imie) $($_.nazwisko) $($_.dzial)" -ForegroundColor Green
    readUserData $_.imie $_.nazwisko $._dzial
  }
}

# Wyłacza poszczególne konto w oparciu o podany login
function disableADAccount {
  $accountToDisable = Read-Host "Wpisz login użytkownika w celu wyłączenia konta (np. jan.kowalski)"
  Disable-ADAccount -Identity $accountToDisable
  Write-Host "Wyłączono konto: $($accountToDisable)" -ForegroundColor Green

  $creationTime = (Get-ADUser -Filter "EmailAddress -eq '$($accountToDisable)@$($domainName)'" -Properties whenCreated).whenCreated
  addToFile "18838_wylaczone_konta_data.csv" "$($creator)|$($creationTime)|$($accountToDisable)@TCO18838.pl"
}

# Zmienia hasło dla użytkownika w domenie w oparciu o podany login
function changeUserPassword {
  $accountToChangePass = Read-Host "Wpisz login użytkownika w celu zmiany jego hasła (np. jan.kowalski)"
  $newPass = Read-Host "Wpisz nowe hasło"

  Set-ADAccountPassword -Identity "$($accountToChangePass)" -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$($newPass)" -Force)

  Write-Host "Zmieniono hasło dla konta: $($accountToChangePass)" -ForegroundColor Green
  $creationTime = (Get-ADUser -Filter "EmailAddress -eq '$($accountToChangePass)@$($domainName)'" -Properties whenCreated).whenCreated
  addToFile "18838_zmiana_hasla_data.csv" "$($creator)|$($creationTime)|$($accountToChangePass)@TCO18838.pl"
}

#Tworzy nowe OU w oparciu o zmienną
function addNewOU {
  $ouCheck = Get-ADOrganizationalUnit -Filter "distinguishedName -eq 'OU=$($ou), $($domainNameDN)'"
  
  if(-not($ouCheck)) {
     New-ADOrganizationalUnit -Name $ou -Path $($domainNameDN) -ProtectedFromAccidentalDeletion $false
      Write-Host "Dodano OU: $($ou)" -ForegroundColor Green
  }
} 

# Tworzy grupę w nowym OU (dla łatwego zarządzania) na podstawie danych od użytkownia
function addNewGroup {
    $groupName = Read-Host "Wpisz nazwę dla nowej grupy zasobów:"
    $newOU = "OU=$($ou),$($domainNameDN)"
    New-ADGroup -Name "$($groupName)" -SamAccountName "$($groupName)" -DisplayName "$($groupName)" `
    -GroupCategory Security -GroupScope Global -Path $newOU
    Write-Host "Utworzono nową grupę: $($groupName)" -ForegroundColor Green
    
    $creationTime = (Get-ADGroup -Filter "SamAccountName -eq 'test'" -Properties whenCreated).whenCreated
    addToFile "18838_create_group.csv" "$($creator)|$($creationTime)|$($groupName)"
}

# Dodaje nowego użytkownika do określonej grupy na podstawie loginu użytkownika
function addGroupMember {
  $group = Read-Host "Wpisz nazwę grupy do której chcesz dodać użytkownika: "
  $member = Read-Host "Wpisz login użytkownika, który ma zostać dodany do wybranej grupy:"
  $userStatment = Get-ADGroupMember -Identity $group | Where-Object {$_.name -eq $member}
  if(-not($userStatment)){
    Add-ADGroupMember -Identity $group -Members $member  
    Write-Host "Użytkownik '$($member)' dodany do grupy '$($group)'" -ForegroundColor Green
    addToFile "18838 zmiana czlonkostwa grup.txt"$($creator)"|$($member)|$($group)"
  } else {
    Write-Host "Użytkownik $($member) istnieje w grupie $($group)" -ForegroundColor Red
  }
}

<#----- Generate reports -----#>

# Generuje listę do pliku .txt z wszystkimi członkami każdej grupy
function reportGroupAccounts {
  (Get-ADGroup -Filter * -Properties name | Select-Object name).name | ForEach-Object {
    createFileWithHeader "$($index) $($_).txt" "login"
    $currentGroup = $_
    (Get-ADGroupMember -Identity $_).name | ForEach-Object {
      addToFile "$($index) $($currentGroup).txt" $_
    }
  }
}


# Generuje określone dane do pliku .csv z wszystkimi wyłączonymi kontami
function  reportDisabledAccounts {
  Get-ADUser -Filter {(Enabled -eq $False)}  -Properties SamAccountName, DistinguishedName, SID, modifyTimeStamp | `
  Select-Object SamAccountName, DistinguishedName, SID, modifyTimeStamp | ForEach-Object {
    addToFile "18838 wylaczone konta.csv" "$($_.SamAccountName)|$($_.distinguishedName)|$($_.SID)|$($_.modifyTimeStamp)"
  } 
}

# Generuje raport z najważniejszymi informacjami o użytkownikach w AD
function reportADUsersInfo {
  Get-ADUser -Filter * -Properties givenName, surName, userPrincipalName, samAccountName, distinguishedName, whenCreated, modifyTimeStamp, LastLogon, PasswordLastSet | `
  Select-Object givenName, surName, userPrincipalName, samAccountName, distinguishedName, whenCreated, modifyTimeStamp, LastLogon, PasswordLastSet | ForEach-Object {
    addToFile "18838 uzytkownicy.csv" "$($_.givenName)|$($_.surName)|$($_.userPrincipalName)|$($_.samAccountName)|$($_.distinguishedName)|$($_.whenCreated)||$($_.modifyTimeStamp)|$($_.LastLogon)|$($_.PasswordLastSet)"
  } 
}

# Generuje raport z informacjami o wszystkich kontach komputerów w domenie
function reportADCoumputersInfo {
  Get-ADComputer -Filter * -Properties Name, SID, distinguishedName, Enabled, LastLogonDate, Created `
    | Select-Object Name, SID, distinguishedName, Enabled, LastLogonDate, Created | ForEach-Object {
      $os = (Get-ComputerInfo).windowsProductName
      $filePath = "$($index)_$($domainName)_$($os).csv"
      createFileWithHeader "$($filePath)" "Nazwa komputera|SID obiektu|DistinguishedName|Status konta|Ostatnia zmiana hasla|Data utworzenia"
      addToFile "$($filePath)" "$($_.Name)|$($_.SID)|$($_.distinguishedName)|$($_.Enabled)|$($_.LastLogon)|$($_.Created)"
  }
}

# Generuje informacje o OU i przekazuje je pliku csv
function reportOUInfo {
  Get-ADOrganizationalUnit -Filter * -Properties distinguishedName, name | Select-Object distinguishedName, name | Sort-Object distinguishedName `
  | ForEach-Object {
    addToFile "18838_$($os).csv" "$($_.name)|$($_.distinguishedName)"
  }
}

# Pokazuje intro w panelu sterowania
function showGreetings {
  Clear-Host
  Write-Host "____________________| Automatyzacja zasobów active directory |____________________" -ForegroundColor Magenta
  Write-Host  "-------> Używaj klawiszy numerycznych zgodnie z poleceniami pomocniczymi, aby wykonać poszczególne akcje AD <------- " -ForegroundColor Green
  Write-Host "Wszystkie utworzone pliki z danymi oraz logami znajdziesz w: $($dirPath)" -ForegroundColor Green 
}


<#----- Zmienne -----#>

$domainName = getDomainName
$domainNameDN = (Get-ADDomain).DistinguishedName

$index = "18838"
$ou = $index
$dirPath = "C:\wit\18838"
$creator = $env:UserName
$os = $($(Get-ComputerInfo).windowsProductName)


<#----- Funkcje inicjalizujące wstępne dane -----#>

#Tworzy ścieżkę katalogową dla plików csv
verifyAndCreateDirPath $dirPath

# Tworzy wszystkie potrzebne pliki csv dla logów i danych
createFileWithHeader "Uzytkownicy.csv" "imie|nazwisko|dzial"
createFileWithHeader "18838 nazwa uzytkownika.csv" "login|haslo"
createFileWithHeader "18838_create_user.csv" "autor|data utworzenia|nazwa użytkownika"
createFileWithHeader "18838 wylaczone konta data.txt" "autor|data utworzenia|nazwa użytkownika"
createFileWithHeader "18838 zmiana hasla data.txt" "autor|data utworzenia|nazwa użytkownika"
createFileWithHeader "18838 create group.csv" "autor grupy|data utworzenia|nazwa grupy"
createFileWithHeader "18838 zmiana członkostwa grup.txt" "autor|nazwa użytkownika|grupa"
createFileWithHeader "18838 wylaczone konta.csv" "Nazwa konta|DistinguishedName|SID|Data ostatniej modyfikacji"
createFileWithHeader "18838 uzytkownicy.csv" "imie|nazwisko|login(UPN)|samacount|lokalizacja w ADDS (DN)|data utworzenia|ostatnia modyfikacja|ostatnie logowanie|ostatnia zmiana hasla"
createFileWithHeader "18838_$($os).csv" "Nazwa OU|DistinguishedName"

#Utworzenie OU
addNewOU

<#----- Panel kontrolny -----#>
# Wczytywanie danych i wykonywanie poszczególnych funkcji odpowiedzialnych za poszczególne działania
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