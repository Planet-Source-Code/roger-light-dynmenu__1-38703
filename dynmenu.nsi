; dynmenu.nsi
;
; Dynamic Menus
;

SetDateSave on
SetDatablockOptimize on
SetOverwrite ifnewer
CRCCheck on
SilentInstall normal

Name "DynMenu v1.4.0"
;Icon "dynmenu.ico"
OutFile "DynMenu_inst.exe"

LicenseText "DynMenu is released under the GNU Public License v2. Source for this distribution is available at http://www.atchoo.org"
LicenseData "copying.txt"

InstallDir "$PROGRAMFILES\DynMenu\"
InstallDirRegKey HKEY_LOCAL_MACHINE "Software\DynMenu" "InstallPath"
DirText "Choose a directory to install in to:"
DirShow show

; uninstall stuff
UninstallText "This will uninstall DynMenu. Hit next to continue."
UninstallIcon "uninst.ico"
ComponentText "This will install DynMenu on your computer."
; Select which optional things you want installed."
;EnabledBitmap inst_dyn1.bmp
;DisabledBitmap inst_dyn2.bmp
!ifndef NOINSTTYPES ; only if not defined
;InstType "Most"
;InstType "Full"
;InstType "More"
;InstType "Base"
!endif

AutoCloseWindow false
;ShowInstDetails show

Section "" ; empty string makes it hidden, so would starting with -
; write uninstall strings
WriteRegStr HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\DynMenu" "DisplayName" "DynMenu (remove only)"
WriteRegStr HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\DynMenu" "UninstallString" '"$INSTDIR\dynmenu-uninst.exe"'

WriteRegStr HKEY_CLASSES_ROOT ".dym" "" "dynamic_menu"
WriteRegStr HKEY_CLASSES_ROOT ".dym\ShellNew" "" ""
WriteRegStr HKEY_CLASSES_ROOT "dynamic_menu\shell\open\command" "" '$INSTDIR\DynMenu.exe %1'
WriteRegStr HKEY_CLASSES_ROOT "dynamic_menu\shell\edit" "" 'Edit Menu'
WriteRegStr HKEY_CLASSES_ROOT "dynamic_menu\shell\edit\command" "" '$INSTDIR\DynMenu.exe /edit %1'
WriteRegStr HKEY_CLASSES_ROOT "dynamic_menu\shell\addtosystray" "" 'Add Menu to SysTray'
WriteRegStr HKEY_CLASSES_ROOT "dynamic_menu\shell\addtosystray\command" "" '$INSTDIR\DynMenu.exe /systray %1'
WriteRegStr HKEY_CLASSES_ROOT "dynamic_menu\DefaultIcon" "" "$INSTDIR\dym.ico"

SetOutPath $SYSDIR
File "C:\Windows\System32\scrrun.dll"
RegDLL "$SYSDIR\scrrun.dll"
File "C:\Windows\System32\Comdlg32.ocx"
RegDLL "$SYSDIR\Comdlg32.ocx"
;File "E:\ggz2k\wggz\libs\ggzxml\bin\ggzmlparse.dll"
;RegDLL "$SYSDIR\ggzmlparse.dll"
;Push $SYSDIR\ggzmlparse.dll
;call AddSharedDLL

SetOutPath $INSTDIR
File "DynMenu.exe"
File "DynMenu.txt"
;File "ChangeLog.html"
;File "DynMenu.chm"
File "dym.ico"
File "atchoo.ico"
File "drive.bmp"
File "dynmenu.bmp"
File "txt.bmp"
File "test.dym"
File "test2.dym"
File "sites.dym"
File "Copying.txt"
File "changelog.txt"
WriteUninstaller "$INSTDIR\dynmenu-uninst.exe"


;=============================

;MessageBox MB_YESNO|MB_ICONQUESTION "View readme file?" IDNO 1
;ExecShell open '"$INSTDIR\dynmenu.txt"'
MessageBox MB_YESNO|MB_ICONQUESTION "Add a link to DynMenu to the desktop?" IDNO lblNoLink
CreateShortCut "$DESKTOP\DynMenu.lnk" "$INSTDIR\DynMenu.exe"

lblNoLink:
ExecShell open '"$INSTDIR"'
Sleep 500
BringToFront
SectionEnd

; special uninstall section.
Section "Uninstall"
DeleteRegKey HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\DynMenu"

DeleteRegKey HKEY_CLASSES_ROOT ".dym"
DeleteRegKey HKEY_CLASSES_ROOT "dynamic_menu"
;Push $SYSDIR\ggzmlparse.dll
;Call un.RemoveSharedDLL 
Delete "$INSTDIR\DynMenu.exe"
Delete "$INSTDIR\dynmenu-uninst.exe"
Delete "$INSTDIR\DynMenu.txt"
Delete "$INSTDIR\dymlarge.ico"
Delete "$INSTDIR\atchoo.ico"
Delete "$INSTDIR\drive.bmp"
Delete "$INSTDIR\dynmenu.bmp"
Delete "$INSTDIR\txt.bmp"
Delete "$INSTDIR\test.dym"
Delete "$INSTDIR\test2.dym"
Delete "$INSTDIR\sites.dym"
Delete "$INSTDIR\copying.txt"
Delete "$INSTDIR\changelog.txt"
Delete "$DESKTOP\DynMenu.lnk"
RMDir "$INSTDIR"
SectionEnd
; eof