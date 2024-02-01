'Copy selected folder

'strUserName = Environ("username")
set objFSO = CreateObject("Scripting.FileSystemObject")
strFolder = "C:\Users\Environ("username")\Downloads\SelectedFolder\"
strNewPath = "C:\"

objFSO.MoveFolder strFolder, strNewPath

'create shortcut to specific file in copied folder on desktop
Set objShell = CreateObject("WScript.Shell")
strDesktop = objShell.SpecialFolders("Desktop")
strShortcut = strDesktop & "\Shortcut.lnk"
objShortcut = objShell.CreateShortcut(strShortcut)
objShortcut.TargetPath = strNewPath & "Beta-Database"\Time_TrackerV1.0 beta version.xlsm"
objShortcut.Description = "Shortcut to Time_TrackerV1.0 beta version.xlsm"
objShortcut.IconLocation = "C:\Beta-Database\Timer.ico, 0"
objShortcut.Save

