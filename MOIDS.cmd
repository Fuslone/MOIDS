@echo off
setlocal enabledelayedexpansion
title Microsoft Office Installer Download Script
mode con: cols=79
color F0
:start
set InstallationType=
set platform=
set version=
set FullVersion=
set specification=
set ProductReleaseID=
set language=
set FullLanguage=
set confirmation=
set FullURL=
cd %USERPROFILE%\Downloads
mode con: lines=9
echo.
echo                              Welcome to use MOIDS
echo.
echo             To select an option,enter the content in parentheses.
echo.
echo                                   [1] Start
echo                                   [0] Exit
echo.
choice /n /c 10
if !errorlevel! == 2 exit /b
:main
set confirmation=yes
:InstallationType
mode con: lines=14
echo.
echo                               Installation Types
echo.
echo                      ------------------------------------
echo.
echo                             [1] Online x32
echo                             [2] Online x64
echo                             [3] Offline x32 ^& x64
echo.
echo                             [0] Go Back
echo.
echo                      ------------------------------------
echo.
choice /n /c 1230
if !errorlevel! == 1 (
    set "InstallationType=Online x32"
    set platform=x86
) else if !errorlevel! == 2 (
    set "InstallationType=Online x64"
    set platform=x64
) else if !errorlevel! == 3 (
    set "InstallationType=Offline x32 & x64"
)
if "!confirmation!" == "no" goto confirmation
if !errorlevel! == 4 goto start
:version
mode con: lines=17
echo.
echo                                    Versions
echo.
echo                      ------------------------------------
echo.
echo                               [1] Microsoft 365
echo                               [2] Office 2024
echo                               [3] Office 2021
echo                               [4] Office 2019
echo                               [5] Office 2016
echo                               [6] Office 2013
echo.
echo                               [0] Go Back
echo.
echo                      ------------------------------------
echo.
choice /n /c 1234560
if not !errorlevel! == 7 (
    if !errorlevel! == 1 (
        set "FullVersion=Microsoft 365"
        set version=365
    ) else if !errorlevel! == 2 (
        set "FullVersion=Office 2024"
        set version=2024
    ) else if !errorlevel! == 3 (
        set "FullVersion=Office 2021"
        set version=2021
    ) else if !errorlevel! == 4 (
        set "FullVersion=Office 2019"
        set version=2019
    ) else if !errorlevel! == 5 (
        set "FullVersion=Office 2016"
        set version=2016
    ) else if !errorlevel! == 6 (
        set "FullVersion=Office 2013"
        set version=2013
    )
    goto !version!
) else (
    if "!confirmation!" == "no" goto confirmation
    goto InstallationType
)
:365
mode con: lines=19
echo.
echo                                 Specification
echo.
echo                      ------------------------------------
echo.
echo                                [1] ProPlus
echo                                [2] AppsBasic
echo                                [3] Business
echo                                [4] EduCloud
echo                                [5] HomePrem
echo                                [6] SmallBusPrem
echo                                [7] ProjectPro
echo                                [8] VisioPro
echo.
echo                                [0] Go Back
echo.
echo                      ------------------------------------
echo.
choice /n /c 123456780
if not !errorlevel! == 9 (
    if !errorlevel! == 1 (
        set specification=ProPlus
    ) else if !errorlevel! == 2 (
        set specification=AppsBasic
    ) else if !errorlevel! == 3 (
        set specification=Business
    ) else if !errorlevel! == 4 (
        set specification=EduCloud
    ) else if !errorlevel! == 5 (
        set specification=HomePrem
    ) else if !errorlevel! == 6 (
        set specification=SmallBusPrem
    ) else if !errorlevel! == 7 (
        set specification=ProjectPro
    ) else if !errorlevel! == 8 (
        set specification=VisioPro
    )
    set ProductReleaseID=O365!specification!Retail
    if "!confirmation!" == "no" goto confirmation
    goto language
) else (
    if "!confirmation!" == "no" goto confirmation
    goto version
)
:2024
mode con: lines=23
echo.
echo                                 Specification
echo.
echo                      ------------------------------------
echo.
echo                                [A] Access
echo                                [B] Excel
echo                                [C] Home
echo                                [D] HomeBusiness
echo                                [E] Outlook
echo                                [F] PowerPoint
echo                                [G] ProjectPro
echo                                [H] ProjectStd
echo                                [I] ProPlus
echo                                [J] VisioPro
echo                                [K] VisioStd
echo                                [L] Word
echo.
echo                                [0] Go Back
echo.
echo                      ------------------------------------
echo.
choice /n /c ABCDEFGHIJKL0
if not !errorlevel! == 13 (
    if !errorlevel! == 1 (
        set specification=Access
    ) else if !errorlevel! == 2 (
        set specification=Excel
    ) else if !errorlevel! == 3 (
        set specification=Home
    ) else if !errorlevel! == 4 (
        set specification=HomeBusiness
    ) else if !errorlevel! == 5 (
        set specification=Outlook
    ) else if !errorlevel! == 6 (
        set specification=PowerPoint
    ) else if !errorlevel! == 7 (
        set specification=ProjectPro
    ) else if !errorlevel! == 8 (
        set specification=ProjectStd
    ) else if !errorlevel! == 9 (
        set specification=ProPlus
    ) else if !errorlevel! == 10 (
        set specification=VisioPro
    ) else if !errorlevel! == 11 (
        set specification=VisioStd
    ) else if !errorlevel! == 12 (
        set specification=Word
    )
    set ProductReleaseID=!specification!2024Retail
    if "!confirmation!" == "no" goto confirmation
    goto language
) else (
    if "!confirmation!" == "no" goto confirmation
    goto version
)
:2021
mode con: lines=29
echo.
echo                                 Specification
echo.
echo                      ------------------------------------
echo.
echo                            [A] Access
echo                            [B] Excel
echo                            [C] HomeBusiness
echo                            [D] HomeStudent
echo                            [E] OneNote
echo                            [F] Outlook
echo                            [G] Personal
echo                            [H] PowerPoint
echo                            [I] Professional
echo                            [J] ProjectPro
echo                            [K] ProjectStd
echo                            [L] ProPlus
echo                            [M] Publisher
echo                            [N] SkypeforBusiness
echo                            [O] Standard
echo                            [P] VisioPro
echo                            [Q] VisioStd
echo                            [R] Word
echo.
echo                            [0] Go Back
echo.
echo                      ------------------------------------
echo.
choice /n /c ABCDEFGHIJKLMNOPQR0
if not !errorlevel! == 19 (
    if !errorlevel! == 1 (
        set specification=Access
    ) else if !errorlevel! == 2 (
        set specification=Excel
    ) else if !errorlevel! == 3 (
        set specification=HomeBusiness
    ) else if !errorlevel! == 4 (
        set specification=HomeStudent
    ) else if !errorlevel! == 5 (
        set specification=OneNote
    ) else if !errorlevel! == 6 (
        set specification=Outlook
    ) else if !errorlevel! == 7 (
        set specification=Personal
    ) else if !errorlevel! == 8 (
        set specification=PowerPoint
    ) else if !errorlevel! == 9 (
        set specification=Professional
    ) else if !errorlevel! == 10 (
        set specification=ProjectPro
    ) else if !errorlevel! == 11 (
        set specification=ProjectStd
    ) else if !errorlevel! == 12 (
        set specification=ProPlus
    ) else if !errorlevel! == 13 (
        set specification=Publisher
    ) else if !errorlevel! == 14 (
        set specification=SkypeforBusiness
    ) else if !errorlevel! == 15 (
        set specification=Standard
    ) else if !errorlevel! == 16 (
        set specification=VisioPro
    ) else if !errorlevel! == 17 (
        set specification=VisioStd
    ) else if !errorlevel! == 18 (
        set specification=Word
    )
    set ProductReleaseID=!specification!2021Retail
    if "!confirmation!" == "no" goto confirmation
    goto language
) else (
    if "!confirmation!" == "no" goto confirmation
    goto version
)
:2019
mode con: lines=30
echo.
echo                                 Specification
echo.
echo                      ------------------------------------
echo.
echo                           [A] Access
echo                           [B] AccessRuntime
echo                           [C] Excel
echo                           [D] HomeBusiness
echo                           [E] HomeStudent
echo                           [F] Outlook
echo                           [G] Personal
echo                           [H] PowerPoint
echo                           [I] Professional
echo                           [J] ProjectPro
echo                           [K] ProjectStd
echo                           [L] ProPlus
echo                           [M] Publisher
echo                           [N] SkypeforBusiness
echo                           [O] SkypeforBusinessEntry
echo                           [P] Standard
echo                           [Q] VisioPro
echo                           [R] VisioStd
echo                           [S] Word
echo.
echo                           [0] Go Back
echo.
echo                      ------------------------------------
echo.
choice /n /c ABCDEFGHIJKLMNOPQRS0
if not !errorlevel! == 20 (
    if !errorlevel! == 1 (
        set specification=Access
    ) else if !errorlevel! == 2 (
        set specification=AccessRuntime
    ) else if !errorlevel! == 3 (
        set specification=Excel
    ) else if !errorlevel! == 4 (
        set specification=HomeBusiness
    ) else if !errorlevel! == 5 (
        set specification=HomeStudent
    ) else if !errorlevel! == 6 (
        set specification=Outlook
    ) else if !errorlevel! == 7 (
        set specification=Personal
    ) else if !errorlevel! == 8 (
        set specification=PowerPoint
    ) else if !errorlevel! == 9 (
        set specification=Professional
    ) else if !errorlevel! == 10 (
        set specification=ProjectPro
    ) else if !errorlevel! == 11 (
        set specification=ProjectStd
    ) else if !errorlevel! == 12 (
        set specification=ProPlus
    ) else if !errorlevel! == 13 (
        set specification=Publisher
    ) else if !errorlevel! == 14 (
        set specification=SkypeforBusiness
    ) else if !errorlevel! == 15 (
        set specification=SkypeforBusinessEntry
    ) else if !errorlevel! == 16 (
        set specification=Standard
    ) else if !errorlevel! == 17 (
        set specification=VisioPro
    ) else if !errorlevel! == 18 (
        set specification=VisioStd
    ) else if !errorlevel! == 19 (
        set specification=Word
    )
    set ProductReleaseID=!specification!2019Retail
    if "!confirmation!" == "no" goto confirmation
    goto language
) else (
    if "!confirmation!" == "no" goto confirmation
    goto version
)
:2016
mode con: lines=37
echo.
echo                                 Specification
echo.
echo                      ------------------------------------
echo.
echo                           [A] Access
echo                           [B] AccessRuntime
echo                           [C] Excel
echo                           [D] HomeBusinessPipc
echo                           [E] HomeBusiness
echo                           [F] HomeStudent
echo                           [G] HomeStudentVNext
echo                           [H] OneNoteFree
echo                           [I] OneNote
echo                           [J] Outlook
echo                           [K] PersonalPipc
echo                           [L] Personal
echo                           [M] PowerPoint
echo                           [N] ProfessionalPipc
echo                           [O] Professional
echo                           [P] ProjectPro
echo                           [Q] ProjectStd
echo                           [R] ProPlus
echo                           [S] Publisher
echo                           [T] SkypeforBusinessEntry
echo                           [U] SkypeforBusiness
echo                           [V] SkypeServiceBypass
echo                           [W] Standard
echo                           [X] VisioPro
echo                           [Y] VisioStd
echo                           [Z] Word
echo.
echo                           [0] Go Back
echo.
echo                      ------------------------------------
echo.
choice /n /c ABCDEFGHIJKLMNOPQRSTUVWXYZ0
if not !errorlevel! == 27 (
    if !errorlevel! == 1 (
        set specification=Access
    ) else if !errorlevel! == 2 (
        set specification=AccessRuntime
    ) else if !errorlevel! == 3 (
        set specification=Excel
    ) else if !errorlevel! == 4 (
        set specification=HomeBusinessPipc
    ) else if !errorlevel! == 5 (
        set specification=HomeBusiness
    ) else if !errorlevel! == 6 (
        set specification=HomeStudent
    ) else if !errorlevel! == 7 (
        set specification=HomeStudentVNext
    ) else if !errorlevel! == 8 (
        set specification=OneNoteFree
    ) else if !errorlevel! == 9 (
        set specification=OneNote
    ) else if !errorlevel! == 10 (
        set specification=Outlook
    ) else if !errorlevel! == 11 (
        set specification=PersonalPipc
    ) else if !errorlevel! == 12 (
        set specification=Personal
    ) else if !errorlevel! == 13 (
        set specification=PowerPoint
    ) else if !errorlevel! == 14 (
        set specification=ProfessionalPipc
    ) else if !errorlevel! == 15 (
        set specification=Professional
    ) else if !errorlevel! == 16 (
        set specification=ProjectPro
    ) else if !errorlevel! == 17 (
        set specification=ProjectStd
    ) else if !errorlevel! == 18 (
        set specification=ProPlus
    ) else if !errorlevel! == 19 (
        set specification=Publisher
    ) else if !errorlevel! == 20 (
        set specification=SkypeforBusinessEntry
    ) else if !errorlevel! == 21 (
        set specification=SkypeforBusiness
    ) else if !errorlevel! == 22 (
        set specification=SkypeServiceBypass
    ) else if !errorlevel! == 23 (
        set specification=Standard
    ) else if !errorlevel! == 24 (
        set specification=VisioPro
    ) else if !errorlevel! == 25 (
        set specification=VisioStd
    ) else if !errorlevel! == 26 (
        set specification=Word
    )
    set ProductReleaseID=!specification!Retail
    if "!confirmation!" == "no" goto confirmation
    goto language
) else (
    if "!confirmation!" == "no" goto confirmation
    goto version
)
:2013
mode con: lines=42
echo.
echo                                 Specification
echo.
echo                      ------------------------------------
echo.
echo                              [1] Access
echo                              [2] Excel
echo                              [3] Groove
echo                              [4] HomeBusinessPipc
echo                              [5] HomeBusiness
echo                              [A] HomeStudent
echo                              [B] InfoPath
echo                              [C] LyncAcademic
echo                              [D] LyncEntry
echo                              [E] Lync
echo                              [F] O365Business
echo                              [G] O365HomePrem
echo                              [H] O365ProPlus
echo                              [I] O365SmallBusPrem
echo                              [J] OneNoteFree
echo                              [K] OneNote
echo                              [L] Outlook
echo                              [M] PersonalPipc
echo                              [N] Personal
echo                              [O] PowerPoint
echo                              [P] ProfessionalPipc
echo                              [Q] Professional
echo                              [R] ProjectPro
echo                              [S] ProjectStd
echo                              [T] ProPlus
echo                              [U] Publisher
echo                              [V] SPD
echo                              [W] Standard
echo                              [X] VisioPro
echo                              [Y] VisioStd
echo                              [Z] Word
echo.
echo                              [0] Go Back
echo.
echo                      ------------------------------------
echo.
choice /n /c 12345ABCDEFGHIJKLMNOPQRSTUVWXYZ0
if not !errorlevel! == 32 (
    if !errorlevel! == 1 (
        set specification=Access
    ) else if !errorlevel! == 2 (
        set specification=Excel
    ) else if !errorlevel! == 3 (
        set specification=Groove
    ) else if !errorlevel! == 4 (
        set specification=HomeBusinessPipc
    ) else if !errorlevel! == 5 (
        set specification=HomeBusiness
    ) else if !errorlevel! == 6 (
        set specification=HomeStudent
    ) else if !errorlevel! == 7 (
        set specification=InfoPath
    ) else if !errorlevel! == 8 (
        set specification=LyncAcademic
    ) else if !errorlevel! == 9 (
        set specification=LyncEntry
    ) else if !errorlevel! == 10 (
        set specification=Lync
    ) else if !errorlevel! == 11 (
        set specification=O365Business
    ) else if !errorlevel! == 12 (
        set specification=O365HomePrem
    ) else if !errorlevel! == 13 (
        set specification=O365ProPlus
    ) else if !errorlevel! == 14 (
        set specification=O365SmallBusPrem
    ) else if !errorlevel! == 15 (
        set specification=OneNoteFree
    ) else if !errorlevel! == 16 (
        set specification=OneNote
    ) else if !errorlevel! == 17 (
        set specification=Outlook
    ) else if !errorlevel! == 18 (
        set specification=PersonalPipc
    ) else if !errorlevel! == 19 (
        set specification=Personal
    ) else if !errorlevel! == 20 (
        set specification=PowerPoint
    ) else if !errorlevel! == 21 (
        set specification=ProfessionalPipc
    ) else if !errorlevel! == 22 (
        set specification=Professional
    ) else if !errorlevel! == 23 (
        set specification=ProjectPro
    ) else if !errorlevel! == 24 (
        set specification=ProjectStd
    ) else if !errorlevel! == 25 (
        set specification=ProPlus
    ) else if !errorlevel! == 26 (
        set specification=Publisher
    ) else if !errorlevel! == 27 (
        set specification=SPD
    ) else if !errorlevel! == 28 (
        set specification=Standard
    ) else if !errorlevel! == 29 (
        set specification=VisioPro
    ) else if !errorlevel! == 30 (
        set specification=VisioStd
    ) else if !errorlevel! == 31 (
        set specification=Word
    )
    set ProductReleaseID=!specification!Retail
    if "!confirmation!" == "no" goto confirmation
    goto language
) else (
    if "!confirmation!" == "no" goto confirmation
    goto version
)
:language
mode con: lines=54
echo.
echo                                    Language
echo.
echo                      ------------------------------------
echo.
echo                              [ar-SA] Arabic
echo                              [bg-BG] Bulgarian
echo                              [cs-CZ] Czech
echo                              [da-DK] Danish
echo                              [de-DE] German
echo                              [el-GR] Greek
echo                              [en-GB] English UK
echo                              [en-US] English
echo                              [es-ES] Spanish
echo                              [es-MX] Spanish Mexico
echo                              [et-EE] Estonian
echo                              [fi-FI] Finnish
echo                              [fr-CA] French Canada
echo                              [fr-FR] French
echo                              [he-IL] Hebrew
echo                              [hi-IN] Hindi
echo                              [hr-HR] Croatian
echo                              [hu-HU] Hungarian
echo                              [id-ID] Indonesian
echo                              [it-IT] Italian
echo                              [ja-JP] Japanese
echo                              [kk-KZ] Kazakh
echo                              [ko-KR] Korean
echo                              [lt-LT] Lithuanian
echo                              [lv-LV] Latvian
echo                              [ms-MY] Malay (Latin)
echo                              [nb-NO] Norwegian Bokmal
echo                              [nl-NL] Dutch
echo                              [pl-PL] Polish
echo                              [pt-BR] Portuguese (Brazil)
echo                              [pt-PT] Portuguese
echo                              [ro-RO] Romanian
echo                              [ru-RU] Russian
echo                              [sk-SK] Slovak
echo                              [sl-SI] Slovenian
echo                              [sr-latn-RS] Serbian (Latin, Serbia)
echo                              [sv-SE] Swedish
echo                              [th-TH] Thai
echo                              [tr-TR] Turkish
echo                              [uk-UA] Ukrainian
echo                              [vi-VN] Vietnamese
echo                              [zh-CN] Chinese (Simplified)
echo                              [zh-TW] Chinese (Traditional)
echo.
echo                              [0] Go Back
echo.
echo                      ------------------------------------
echo.
set /p language=
if "!language!" == "ar-SA" (
    set FullLanguage=Arabic
) else if "!language!" == "bg-BG" (
    set FullLanguage=Bulgarian
) else if "!language!" == "cs-CZ" (
    set FullLanguage=Czech
) else if "!language!" == "da-DK" (
    set FullLanguage=Danish
) else if "!language!" == "de-DE" (
    set FullLanguage=German
) else if "!language!" == "el-GR" (
    set FullLanguage=Greek
) else if "!language!" == "en-GB" (
    set "FullLanguage=English UK"
) else if "!language!" == "en-US" (
    set FullLanguage=English
) else if "!language!" == "es-ES" (
    set FullLanguage=Spanish
) else if "!language!" == "es-MX" (
    set "FullLanguage=Spanish Mexico"
) else if "!language!" == "et-EE" (
    set FullLanguage=Estonian
) else if "!language!" == "fi-FI" (
    set FullLanguage=Finnish
) else if "!language!" == "fr-CA" (
    set "FullLanguage=French Canada"
) else if "!language!" == "fr-FR" (
    set FullLanguage=French
) else if "!language!" == "he-IL" (
    set FullLanguage=Hebrew
) else if "!language!" == "hi-IN" (
    set FullLanguage=Hindi
) else if "!language!" == "hr-HR" (
    set FullLanguage=Croatian
) else if "!language!" == "hu-HU" (
    set FullLanguage=Hungarian
) else if "!language!" == "id-ID" (
    set FullLanguage=Indonesian
) else if "!language!" == "it-IT" (
    set FullLanguage=Italian
) else if "!language!" == "ja-JP" (
    set FullLanguage=Japanese
) else if "!language!" == "kk-KZ" (
    set FullLanguage=Kazakh
) else if "!language!" == "ko-KR" (
    set FullLanguage=Korean
) else if "!language!" == "lt-LT" (
    set FullLanguage=Lithuanian
) else if "!language!" == "lv-LV" (
    set FullLanguage=Latvian
) else if "!language!" == "ms-MY" (
    set "FullLanguage=Malay (Latin)"
) else if "!language!" == "nb-NO" (
    set "FullLanguage=Norwegian Bokmal"
) else if "!language!" == "nl-NL" (
    set FullLanguage=Dutch
) else if "!language!" == "pl-PL" (
    set FullLanguage=Polish
) else if "!language!" == "pt-BR" (
    set "FullLanguage=Portuguese (Brazil)"
) else if "!language!" == "pt-PT" (
    set FullLanguage=Portuguese
) else if "!language!" == "ro-RO" (
    set FullLanguage=Romanian
) else if "!language!" == "ru-RU" (
    set FullLanguage=Russian
) else if "!language!" == "sk-SK" (
    set FullLanguage=Slovak
) else if "!language!" == "sl-SI" (
    set FullLanguage=Slovenian
) else if "!language!" == "sr-latn-RS" (
    set "FullLanguage=Serbian (Latin, Serbia)"
) else if "!language!" == "sv-SE" (
    set FullLanguage=Swedish
) else if "!language!" == "th-TH" (
    set FullLanguage=Thai
) else if "!language!" == "tr-TR" (
    set FullLanguage=Turkish
) else if "!language!" == "uk-UA" (
    set FullLanguage=Ukrainian
) else if "!language!" == "vi-VN" (
    set FullLanguage=Vietnamese
) else if "!language!" == "zh-CN" (
    set "FullLanguage=Chinese (Simplified)"
) else if "!language!" == "zh-TW" (
    set "FullLanguage=Chinese (Traditional)"
) else if "!language!" == "0" (
    if "!confirmation!" == "no" goto confirmation
    goto version
) else (
    mode con: lines=4
    echo.
    echo             Please enter the content in parentheses.
    echo.
    timeout /t 3 > nul
    goto language
)
:confirmation
mode con: lines=18
set confirmation=no
echo.
echo                                  Confirmation
echo.
echo                      -------------------------------------
echo.
echo                  [1] Installation Type : !InstallationType!
echo                  [2] Version : !FullVersion!
echo                  [3] Specification : !specification!
echo                  [4] Language : !FullLanguage!
echo.
echo                  [0] Go Back To The Menu
echo.
echo                      -------------------------------------
echo.
echo                        Are you sure ? If yes,enter "y".
echo            If not,enter the number of the value you want to modify.
echo.
choice /n /c Y12340
if !errorlevel! == 2 (
    goto InstallationType
) else if !errorlevel! == 3 (
    goto version
) else if !errorlevel! == 4 (
    goto !version!
) else if !errorlevel! == 5 (
    goto language
) else if !errorlevel! == 6 (
    goto start
)
if "!InstallationType!" == "Offline x32 & x64" (
    if "!FullVersion!" == "Office 2013" (
        set "FullURL=https://officecdn.microsoft.com/db/39168d7e-077b-48e7-872c-b232c3e72675/media/!language!/!ProductReleaseID!.img"
    ) else (
        set "FullURL=https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/!language!/!ProductReleaseID!.img"
    )
) else (
    if "!FullVersion!" == "Office 2013" (
        set "FullURL=https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=!ProductReleaseID!&platform=!platform!&language=!language!&version=O15GA"
    ) else (
        set "FullURL=https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=!ProductReleaseID!&platform=!platform!&language=!language!&version=O16GA"
    )
)
:check-and-detect
mode con: lines=8
echo.
echo Checking network connection...
echo.
ping officecdn.microsoft.com > nul
ping c2rsetup.officeapps.live.com > nul
if !errorlevel! == 1 (
    powershell -c "Write-Host 'There is a problem with the network,please check the connection.' -ForegroundColor Red"
    echo.
    powershell -c "Write-Host 'Please check the network connection and then continue.' -ForegroundColor Blue"
    echo.
    timeout /t 3 > nul
    goto confirmation
) else (
    powershell -c "Write-Host 'The network connection is normal.' -ForegroundColor Green"
    echo.
    timeout /t 3 > nul
    cls
)
echo.
echo Detecting whether the resource exists...
echo.
powershell -c "(Invoke-WebRequest -Uri '!FullURL!' -Method Head).Headers" | findstr "Content-Disposition" > nul
if !errorlevel! == 1 (
    powershell -c "Write-Host 'The resource does not exist.' -ForegroundColor Red"
    echo.
    powershell -c "Write-Host 'Please change the parameters and try again.' -ForegroundColor Blue"
    echo.
    timeout /t 3 > nul
    goto confirmation
) else (
    powershell -c "Write-Host 'The resource exists and the download will start soon.' -ForegroundColor Green"
    echo.
    timeout /t 3 > nul
)
:download
mode con: lines=12
echo.
echo Starting download...
echo.
curl -# -O -J "!FullURL!"
if !errorlevel! == 0 (
    powershell -c "Write-Host 'The download is complete.' -ForegroundColor Green"
    echo.
    echo Returning to the menu soon.
    echo.
    pause
    goto start
) else (
    echo.
    powershell -c "Write-Host 'The download failed.' -ForegroundColor Red"
    echo.
    powershell -c "Write-Host 'Please try again.' -ForegroundColor Blue"
    echo.
    pause
    goto confirmation
)