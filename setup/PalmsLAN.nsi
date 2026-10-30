; Palms LAN Center Setup: installs Palms LAN Center on Windows 98 and later (NSIS 3, ANSI).
; makensis -DVERSION=0.9.3-alpha -DNUMVER=0.9.3 -DSRC=<folder with the release files> -DOUT=<setup.exe> PalmsLAN.nsi

Unicode false
SetCompressor /SOLID lzma
RequestExecutionLevel admin
XPStyle on

!include "LogicLib.nsh"
!include "WinVer.nsh"

!ifndef VERSION
  !define VERSION "0.0.0"
!endif
!ifndef NUMVER
  !define NUMVER "${VERSION}"
!endif
!ifndef SRC
  !define SRC "files"
!endif
!ifndef OUT
  !define OUT "PalmsLAN-${VERSION}-setup.exe"
!endif
!ifndef SIZEKB
  !define SIZEKB 2048
!endif
!define REG "Software\Palms Arcade\LAN Center"
!define UNINST "Software\Microsoft\Windows\CurrentVersion\Uninstall\PalmsLANCenter"

Name "Palms LAN Center ${VERSION}"
Caption "Palms LAN Center ${VERSION} Setup"
UninstallCaption "Palms LAN Center Uninstall"
OutFile "${OUT}"
Icon "palm.ico"
UninstallIcon "palm.ico"
BrandingText "Palms Arcade"
InstallDir "$PROGRAMFILES\PalmsLANCenter"
InstallDirRegKey HKLM "${REG}" "program_folder"
ShowInstDetails show
ShowUninstDetails show
InstallColors /windows

VIProductVersion "${NUMVER}.0"
VIAddVersionKey "ProductName" "Palms LAN Center"
VIAddVersionKey "CompanyName" "Palms Arcade"
VIAddVersionKey "FileDescription" "Palms LAN Center Setup"
VIAddVersionKey "FileVersion" "${NUMVER}"
VIAddVersionKey "ProductVersion" "${VERSION}"
VIAddVersionKey "LegalCopyright" "Palms Arcade"

Var GAMESDIR

ComponentText "Select the components to install. Click Next to continue." "" "Components to install:"
DirText "Setup will install Palms LAN Center in the following folder. To install in a different folder, click Browse and select another folder. Click Next to continue." "Program folder" "" "Select the folder to install Palms LAN Center in:"
UninstallText "This removes Palms LAN Center from this PC, together with all games it installed (including their configuration files and saved games) and its settings. Click Uninstall to continue."

Page components
Page directory
PageEx directory
  Caption ": Games Folder"
  DirVar $GAMESDIR
  DirText "Setup will set Palms LAN Center to install games from the Game Library in the following folder. To use a different folder, click Browse and select another folder. Click Install to continue." "Games folder" "" "Select the folder to install games to:"
  PageCallbacks "" "" GamesLeave
PageExEnd
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles

; ------------------------------------------------------------------ installing

Section "Palms LAN Center (required)" SecMain
  SectionIn RO
  Call CheckClosed

  SetOutPath "$INSTDIR"
  SetOverwrite on
  File "${SRC}\PalmsLAN.exe"
  File "${SRC}\palmsget.exe"
  File "${SRC}\7za.exe"
  File "${SRC}\ding.wav"
  File "${SRC}\dingdong.wav"
  File "${SRC}\horn.wav"
  File "${SRC}\triple.wav"
  File "${SRC}\palms-clean.vbs"

  ; the Visual Basic 6.0 runtime: Windows 2000 and XP have it, Windows 98 may not
!if /FileExists "${SRC}\vbrun60sp6.exe"
  ${IfNot} ${FileExists} "$SYSDIR\MSVBVM60.DLL"
    DetailPrint "Installing the Visual Basic 6.0 runtime..."
    InitPluginsDir
    SetOutPath "$PLUGINSDIR"
    File "${SRC}\vbrun60sp6.exe"
    ExecWait '"$PLUGINSDIR\vbrun60sp6.exe" /Q' $0
    SetOutPath "$INSTDIR"
    ${IfNot} ${FileExists} "$SYSDIR\MSVBVM60.DLL"
      MessageBox MB_OK|MB_ICONEXCLAMATION "Unable to install the Visual Basic 6.0 runtime. Palms LAN Center requires it.$\r$\n$\r$\nRestart Windows and run Setup again." /SD IDOK
    ${EndIf}
  ${EndIf}
!else
  ${IfNot} ${FileExists} "$SYSDIR\MSVBVM60.DLL"
    MessageBox MB_OK|MB_ICONEXCLAMATION "The Visual Basic 6.0 runtime (MSVBVM60.DLL) is not installed. Palms LAN Center requires it.$\r$\n$\r$\nInstall the Visual Basic 6.0 runtime (vbrun60sp6.exe), then start Palms LAN Center." /SD IDOK
  ${EndIf}
!endif

  ; the Winsock control, shared with other programs: in the system folder, and registered
  DetailPrint "Registering MSWINSCK.OCX..."
  SetOutPath "$SYSDIR"
  SetOverwrite ifnewer
  File "${SRC}\MSWINSCK.OCX"
  SetOverwrite on
  SetOutPath "$INSTDIR"
  ClearErrors
  RegDLL "$SYSDIR\MSWINSCK.OCX"
  ${If} ${Errors}
    MessageBox MB_OK|MB_ICONEXCLAMATION "Unable to register MSWINSCK.OCX. Palms LAN Center can't connect to the Palms Server without it.$\r$\n$\r$\nLog on to Windows as an administrator and run Setup again." /SD IDOK
  ${EndIf}

  ; settings Palms LAN Center reads (the rest it makes itself)
  WriteRegStr HKLM "${REG}" "games_folder" "$GAMESDIR"
  WriteRegStr HKLM "${REG}" "program_folder" "$INSTDIR"

  ; Windows XP firewall (Windows 98 has none)
  ${If} ${AtLeastWinXP}
    DetailPrint "Adding a Windows Firewall exception..."
    nsExec::Exec 'netsh firewall add allowedprogram program="$INSTDIR\PalmsLAN.exe" name="Palms LAN Center" mode=ENABLE'
    Pop $0
  ${EndIf}

  ; Start menu, for everyone using this PC
  SetShellVarContext all
  CreateDirectory "$SMPROGRAMS\Palms Arcade"
  CreateShortCut "$SMPROGRAMS\Palms Arcade\Palms LAN Center.lnk" "$INSTDIR\PalmsLAN.exe" "" "$INSTDIR\PalmsLAN.exe" 0
  CreateShortCut "$SMPROGRAMS\Palms Arcade\Uninstall Palms LAN Center.lnk" "$INSTDIR\uninst.exe"

  ; Add or Remove Programs
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${UNINST}" "DisplayName" "Palms LAN Center"
  WriteRegStr HKLM "${UNINST}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "${UNINST}" "Publisher" "Palms Arcade"
  WriteRegStr HKLM "${UNINST}" "DisplayIcon" "$INSTDIR\PalmsLAN.exe,0"
  WriteRegStr HKLM "${UNINST}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKLM "${UNINST}" "UninstallString" '"$INSTDIR\uninst.exe"'
  WriteRegDWORD HKLM "${UNINST}" "NoModify" 1
  WriteRegDWORD HKLM "${UNINST}" "NoRepair" 1
  WriteRegDWORD HKLM "${UNINST}" "EstimatedSize" ${SIZEKB}
SectionEnd

Section "Desktop shortcut" SecDesktop
  SetShellVarContext all
  CreateShortCut "$DESKTOP\Palms LAN Center.lnk" "$INSTDIR\PalmsLAN.exe" "" "$INSTDIR\PalmsLAN.exe" 0
SectionEnd

Section /o "Start Palms LAN Center with Windows" SecStartup
  SetShellVarContext all
  CreateShortCut "$SMSTARTUP\Palms LAN Center.lnk" "$INSTDIR\PalmsLAN.exe" "" "$INSTDIR\PalmsLAN.exe" 0
SectionEnd

Function .onInit
  ; the games folder Palms LAN Center already uses on this PC, or the default
  ReadRegStr $GAMESDIR HKLM "${REG}" "games_folder"
  ${If} $GAMESDIR == ""
    StrCpy $GAMESDIR "C:\LANPartyGames"
  ${EndIf}
FunctionEnd

; a folder on a local drive, like C:\LANPartyGames or D:\Games (not a network folder)
Function GamesLeave
  StrCpy $0 $GAMESDIR 2 1
  ${If} $0 != ":\"
    MessageBox MB_OK|MB_ICONEXCLAMATION "Select a folder on a local drive (for example, C:\LANPartyGames or D:\Games)."
    Abort
  ${EndIf}
FunctionEnd

; PalmsLAN.exe can't be replaced while it's running
!macro CHECK_CLOSED UN
Function ${UN}CheckClosed
  ${Do}
    ${IfNot} ${FileExists} "$INSTDIR\PalmsLAN.exe"
      ${Break}
    ${EndIf}
    ClearErrors
    FileOpen $0 "$INSTDIR\PalmsLAN.exe" a
    ${IfNot} ${Errors}
      FileClose $0
      ${Break}
    ${EndIf}
    MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "Palms LAN Center is running. Close it, then click Retry." /SD IDCANCEL IDRETRY retry
    Quit
    retry:
  ${Loop}
FunctionEnd
!macroend
!insertmacro CHECK_CLOSED ""
!insertmacro CHECK_CLOSED "un."

; ------------------------------------------------------------------ uninstalling

Section "Uninstall"
  Call un.CheckClosed

  ; every game it installed (folders, registry settings, shortcuts, firewall exceptions)
  ; and its settings: palms-clean.vbs does that, quietly
  ${If} ${FileExists} "$INSTDIR\palms-clean.vbs"
    DetailPrint "Removing the games installed by Palms LAN Center..."
    ExecWait 'wscript.exe "$INSTDIR\palms-clean.vbs" /quiet'
  ${EndIf}
  DeleteRegKey HKLM "Software\Palms Arcade"

  ${If} ${AtLeastWinXP}
    nsExec::Exec 'netsh firewall delete allowedprogram program="$INSTDIR\PalmsLAN.exe"'
    Pop $0
  ${EndIf}

  SetShellVarContext all
  Delete "$DESKTOP\Palms LAN Center.lnk"
  Delete "$SMSTARTUP\Palms LAN Center.lnk"
  Delete "$SMPROGRAMS\Palms Arcade\Palms LAN Center.lnk"
  Delete "$SMPROGRAMS\Palms Arcade\Uninstall Palms LAN Center.lnk"
  RMDir "$SMPROGRAMS\Palms Arcade"
  ; palms-update.vbs makes one for the user who ran it
  SetShellVarContext current
  Delete "$DESKTOP\Palms LAN Center.lnk"

  ; the program, and what it made in its folder (MSWINSCK.OCX and the Visual Basic runtime stay:
  ; other programs use them too)
  Delete "$INSTDIR\PalmsLAN.exe"
  Delete "$INSTDIR\palmsget.exe"
  Delete "$INSTDIR\7za.exe"
  Delete "$INSTDIR\MSWINSCK.OCX"
  Delete "$INSTDIR\ding.wav"
  Delete "$INSTDIR\dingdong.wav"
  Delete "$INSTDIR\horn.wav"
  Delete "$INSTDIR\triple.wav"
  Delete "$INSTDIR\palms-clean.vbs"
  Delete "$INSTDIR\palms-update.vbs"
  RMDir /r "$INSTDIR\icons"
  RMDir /r "$INSTDIR\temp"
  Delete "$INSTDIR\uninst.exe"
  RMDir "$INSTDIR"

  DeleteRegKey HKLM "${UNINST}"
SectionEnd
