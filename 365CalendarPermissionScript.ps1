#   written by J. Norrby, 2024

#   Current language support: English, Spanish, French, German, Italian, Portuguese, Dutch, Chinese (Simplified), Japanese, 
#   Korean, Russian, Arabic, Turkish, Hindi, Bengali, Urdu, Swedish, Polish, Danish, Finnish, Greek, Hebrew, Thai, Vietnamese, Norwegian

Write-Host "`n`n`n`n`n`n`n - - - Microsoft 365 Calendar Permission Script - - - `n`n`n`n`n"
Write-Host "Enter ? for help with any prompt`n"

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    $YN = Read-Host "ExchangeOnlineManagement module is not installed on your system but is required to run this script. Do you want to install it? (Y/N)"

    if ($response -eq "Y") {
        try {
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
            Write-Host "ExchangeOnlineManagement module has been installed successfully." -ForegroundColor Green
        } catch {
            Write-Host "An error occurred while trying to install the ExchangeOnlineManagement module." -ForegroundColor Red
        }
    } else {
        Write-Host "Installation of ExchangeOnlineManagement module was skipped." -ForegroundColor Yellow
    }
} else {
}

Import-Module ExchangeOnlineManagement

do {
    $AdminUPN = Read-Host "Admin UPN"
    Write-Host " "

    if ($AdminUPN -eq "?") {
        Write-Host "`n---HELP---`n`nAdmin UPN is the username of an admin account on the Microsoft 365 tenant, e.g admin@company.onmicrosoft.com`n`n----------`n`n"
    }
    elseif ($AdminUPN -like "*@*.*"){
        Break
    }
    else {
        Write-Host "Invalid input. Please enter a valid admin account UPN, e.g., admin@company.onmicrosoft.com`n" -ForegroundColor Red
    }
} while ($true)

Write-Host "Logging in as $adminUPN, MFA might be required..." -ForegroundColor Green
Start-Sleep -Seconds 2
Connect-ExchangeOnline -UserPrincipalName $AdminUPN

Write-Host " "

do {
    $global:TargetMailbox = Read-Host "Target UPN"

    if ($TargetMailbox -eq "?") {
        Write-Host "`n---HELP---`n`nTarget UPN is the user who's calendar you wish to give other users permissions to. e.g., user@company.com`n`n----------`n`n"
    }
    elseif ($TargetMailbox -like "*@*.*"){
        Write-Host "is " -NoNewline
        Write-Host "$TargetMailbox" -ForegroundColor Cyan -NoNewline
        Write-Host " correct?" -NoNewline
        $YN = Read-host "[Y/N]"
        if ($YN -eq "Y") {
            Break
        }
        elseif ($YN -eq "N") {
        }
    }
    else {
        Write-Host "Invalid input. Please enter a valid user account UPN, e.g., user@company.com`n" -ForegroundColor Red
    }
} while ($true)

$Folders = Get-MailboxFolderStatistics -Identity $TargetMailbox | Select-Object Name, FolderPath
$UserCalendar = $null

foreach ($folder in $folders) {
    switch -Regex ($folder.Name) {
        "calendar" { $UserCalendar = "Calendar"; break }
        "calendario" { $UserCalendar = "calendario"; break }
        "calendrier" { $UserCalendar = "Calendrier"; break }
        "kalender" { $UserCalendar = "kalender"; break }
        "calendário" { $UserCalendar = "calendário"; break }
        "agenda" { $UserCalendar = "agenda"; break }
        "日历" { $UserCalendar = "日历"; break }
        "カレンダー" { $UserCalendar = "カレンダー"; break }
        "캘린더" { $UserCalendar = "캘린더"; break }
        "календарь" { $UserCalendar = "календарь"; break }
        "التقويم" { $UserCalendar = "التقويم"; break }
        "takvim" { $UserCalendar = "takvim"; break }
        "कैलेंडर" { $UserCalendar = "कैलेंडर"; break }
        "ক্যালেন্ডার" { $UserCalendar = "ক্যালেন্ডার"; break }
        "کیلنڈر" { $UserCalendar = "کیلنڈر"; break }
        "kalender" { $UserCalendar = "kalender"; break }
        "kalendarz" { $UserCalendar = "kalendarz"; break }
        "kalenteri" { $UserCalendar = "kalenteri"; break }
        "ημερολόγιο" { $UserCalendar = "ημερολόγιο"; break }
        "לוח שנה" { $UserCalendar = "לוח שנה"; break }
        "ปฏิทิน" { $UserCalendar = "ปฏิทิน"; break }
        "lịch" { $UserCalendar = "lịch"; break }
    }
}

do {
    $list = @("`n---Available Permissions---`n", "1 reviewer - read", "2 contributor - write", "3 editor - write/read/modify can't delete posts", "4 publishingauthor - write/read/delete but only on their own posts", "5 owner - full access")
    $list | ForEach-Object { Write-Host $_ }
    $global:Permission = read-host "`nPermission"

    if ($Permission -eq "?") {
        $list | ForEach-Object { Write-Host $_ }
        Write-Host "Enter one of the permissions above or its ID"
    }
    elseif ($Permission -eq "reviewer" -or "1") {
        $Permission = "reviewer"
        break
    }
    elseif ($Permission -eq "contributor" -or "2") {
        $Permission = "contributor"
        break
    }
    elseif ($Permission -eq "editor" -or "3") {
        $Permission = "editor"
        break
    }
    elseif ($Permission -eq "publishingauthor" -or "4") {
        $Permission = "publishingauthor"
        break
    }
    elseif ($Permission -eq "owner" -or "5") {
        $Permission = "owner"
        break
    }
    else {
        Write-Host "Invalid input. Please enter a valid choice of permission or ? to see available permissions`n" -ForegroundColor Red
    }
} while ($true)
Function TargetUsers {
do {
    $global:TargetUser = Read-Host "User UPN"

    if ($TargetUser -eq "?") {
        Write-Host "`n---HELP---`n`nUser UPN is the user who you wish to give permissions to. e.g., user@company.com`n`n----------`n`n"
    }
    elseif ($TargetUser -like "*@*.*"){
        Write-Host "Do you wish to give " -NoNewline
        Write-Host "$Targetuser" -ForegroundColor Cyan -NoNewline
        Write-Host ", " -NoNewline
        Write-Host "$Permission" -ForegroundColor Cyan -NoNewline
        Write-Host " permissions on " -NoNewline
        Write-host "$TargetMailbox" -ForegroundColor Cyan -NoNewline
        Write-Host "'s " -NoNewline
        Write-Host "$UserCalendar" -ForegroundColor Cyan -NoNewline
        $YN = Read-host "? [Y/N]"
        if ($YN -eq "Y") {
            Write-Host "`nGiving $Targetuser $Permission permission on $Targetmailbox's $UserCalendar ...`n" -ForegroundColor Yellow
            Add-MailboxFolderPermission -Identity $TargetMailbox":\"$UserCalendar -user $TargetUser -AccessRights $Permission
            Break
        }
        elseif ($YN -eq "N") {
        }
    }
    else {
        Write-Host "Invalid input. Please enter a valid user account UPN, e.g., user@company.com`n" -ForegroundColor Red
    }
} while ($true)
}

Write-Host " "

TargetUsers

do {
    
    Write-Host "do you wish to give another user " -NoNewline
    Write-Host "$Permission" -ForegroundColor Cyan -NoNewline
    Write-Host " permissions on " -NoNewline
    Write-Host "$TargetMaiblox" -ForegroundColor Cyan -NoNewline
    Write-Host "'s" -NoNewline
    Write-Host "$Usercalendar" -ForegroundColor Cyan -NoNewline  
    $YN = Read-host "? [Y/N]"

    if ($YN -eq "Y") {
        TargetUsers
    }
    else {
        Break
    }
} while ($true)

Write-Host "Quitting script..." -ForegroundColor Red
Start-sleep -seconds 3