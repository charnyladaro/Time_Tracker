@echo off

set "folder=%userprofile%\Downloads\Beta-Database"
set "file=%folder%\Time_TrackerV1.0 beta version.xlsm"
set "icon=%folder%\Timer.ico"
set "desktop=%userprofile\Desktop\"

Move %folder% %Desktop%
mklink /H "%desktop%\Beta-Database.lnk" %file%
set "file=%desktop%\Beta-Database.lnk"

set "shell=WScript.Shell"
set "link=shell.CreateShortcut(file)"
link.IconLocation = icon
link.save
