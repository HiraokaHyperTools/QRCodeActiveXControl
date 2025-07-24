; example2.nsi
;
; This script is based on example1.nsi, but it remember the directory, 
; has uninstall support and (optionally) installs start menu shortcuts.
;
; It will install example2.nsi into a directory that the user selects,

;--------------------------------

Unicode true
XPStyle on

!define APP "QRCodeActiveXControl"
!define COM "HIRAOKA HYPERS TOOLS, Inc."
!define BINDIR "RELEASE"
!define BINDIR64 "x64\RELEASE"
!define BINDIRA64X "ARM64EC\RELEASE"

!system 'DefineAsmVer.exe "${BINDIR}\QRCStamp.ocx" "!define VER ""[SVER]"" " > Tmpver.nsh'
!include "Tmpver.nsh"
!searchreplace APV ${VER} "." "_"

!define TTL "QRCodeActiveXControl ${VER}"

!system 'MySign "${BINDIR}\QRCStamp.ocx" "${BINDIR64}\QRCStamp.ocx" "${BINDIRA64X}\QRCStamp.ocx"'
!finalize 'MySign "%1"'

; The name of the installer
Name "${TTL}"

; The file to write
OutFile "Setup_${APP}_${APV}.exe"

; The default installation directory
InstallDir "$PROGRAMFILES\${APP}"

; Registry key to check for directory (so if you install again, it will 
; overwrite the old one automatically)
InstallDirRegKey HKLM "Software\${COM}\${APP}" "Install_Dir"

; Request application privileges for Windows Vista
RequestExecutionLevel admin

!include "FileFunc.nsh"
!include "LogicLib.nsh"
!include "x64.nsh"

!insertmacro Locate

;--------------------------------

; Pages

Page license
Page directory
Page components
Page instfiles

LicenseData "LICENSE"

UninstPage uninstConfirm
UninstPage instfiles

;--------------------------------

; The stuff to install
Section ""

  SectionIn RO
  
  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put file there
  SetOutPath $INSTDIR
  File "${BINDIR}\QRCStamp.ocx"
  ${If} ${IsNativeAMD64}
    SetOutPath $INSTDIR\x64
    File "${BINDIR64}\QRCStamp.ocx"
  ${ElseIf} ${IsNativeARM64}
    SetOutPath $INSTDIR\x64
    File "${BINDIRA64X}\QRCStamp.ocx"
  ${EndIf}

  SetOutPath $INSTDIR

  ; Write the installation path into the registry
  WriteRegStr HKLM "SOFTWARE\${COM}\${APP}" "Install_Dir" "$INSTDIR"
  
  ; Write the uninstall keys for Windows
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}" "DisplayName" "${TTL}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}" "NoRepair" 1
  WriteUninstaller "uninstall.exe"
  
SectionEnd

Section "Register QRCStamp.ocx (32 bit)"
  ExecWait 'REGSVR32 /s "$INSTDIR\QRCStamp.ocx"' $0
  DetailPrint "Result: $0"
SectionEnd

Section "Register QRCStamp.ocx (64 bit)"
  ExecWait 'REGSVR32 /s "$INSTDIR\x64\QRCStamp.ocx"' $0
  DetailPrint "Result: $0"
SectionEnd

Section "Clear QRCStampLib.exd files"
  ${Locate} "$TEMP" "/L=F /M=QRCStampLib.exd" "DeleteIt"
SectionEnd

Function DeleteIt
  Delete "$R9"
  Push $0
FunctionEnd

; Optional section (can be disabled by the user)
Section /o "Start Menu Shortcuts"

  CreateDirectory "$SMPROGRAMS\${APP}"
  CreateShortCut "$SMPROGRAMS\${APP}\Uninstall.lnk" "$INSTDIR\uninstall.exe" "" "$INSTDIR\uninstall.exe" 0
  
SectionEnd

;--------------------------------

; Uninstaller

Section "Uninstall"
  
  ; Remove registry keys
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}"
  DeleteRegKey HKLM "SOFTWARE\${COM}\${APP}"

  ; Remove files and uninstaller
  ExecWait 'REGSVR32 /s /u "$INSTDIR\x64\QRCStamp.ocx"' $0
  DetailPrint "Result: $0"
  Delete                   "$INSTDIR\x64\QRCStamp.ocx"
  
  ExecWait 'REGSVR32 /s /u "$INSTDIR\QRCStamp.ocx"' $0
  DetailPrint "Result: $0"
  Delete                   "$INSTDIR\QRCStamp.ocx"
  
  Delete "$INSTDIR\uninstall.exe"

  ; Remove shortcuts, if any
  Delete "$SMPROGRAMS\${APP}\Uninstall.lnk"

  ; Remove directories used
  RMDir "$SMPROGRAMS\${APP}"
  RMDir "$INSTDIR"

SectionEnd
