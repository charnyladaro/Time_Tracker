Set fso = CreateObject("Scripting.FileSystemObject")
 
'Move OldFolderName to C:\Dst and rename to NewFolderName
'fso.MoveFolder "C:\OlderFolderName", C:\Dst\NewFolderName
'Set wshShell = CreateObject( "WScript.Shell" )
'UserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
'WScript.Echo "User Name: " & UserName
'Move all OldFolders to C:\Dst. This will keep their names
fso.MoveFolder "C:\Users\charn\Downloads\Installer\Time_Logger", "C:\"
'fso.MoveFolder "C:\Users\" + UserName + "\Downloads\Installer", "C:\"

Set objShell = WScript.CreateObject("WScript.Shell")
 
'All users Desktop
allUsersDesktop = objShell.SpecialFolders("AllUsersDesktop")
 
'The current users Desktop
usersDesktop = objShell.SpecialFolders("Desktop")
 
'Where to create the new shorcut
Set objShortCut = objShell.CreateShortcut(usersDesktop & "\Time_Logger.lnk")
 
'What does the shortcut point to
objShortCut.TargetPath = "C:\Time_Logger\Success.txt"
 
objShortCut.IconLocation = "C:\Time_Logger\Timer.ico"
'Add a description
objShortCut.Description = "Run the Notepad."
 
'Create the shortcut
objShortCut.Save
WScript.Sleep 3000
x=msgbox("Time Tracker Successfully Installed", 0 + 64, "Time Logger Installer")