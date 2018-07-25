!define PRODUCT_NAME "Outlook Mail Alert"

!include MUI2.nsh

; MUI Settings
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

;Registry keys
!define PRODUCT_DIR_REGKEY "Software\OutlookMailAlert"
!define REG_SHELL_CLEANER_PATH "*\shell\OutlookMailAlert"
!define REG_SHELL_CLEANER_COMMAND_PATH "*\shell\OutlookMailAlert\command"

;--------------------------------
;General

Name "${PRODUCT_NAME}"
OutFile "OutlookMessageAllert.exe"
RequestExecutionLevel admin
InstallDir "$PROGRAMFILES\Outlook Mail Alert"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" "Install_Dir"
ShowInstDetails show
ShowUnInstDetails show

;--------------------------------
;Pages

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

; Language files
!insertmacro MUI_LANGUAGE "English"

; MUI end ------

Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd

;--------------------------------

Section "Outlook Mail Alert" MainSec
  SectionIn RO
  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer
  File "..\OutlookMessageAllert\bin\Release\OutlookMessageAllert.dll"
  File "..\OutlookMessageAllert\bin\Release\OutlookMessageAllert.vsto"
  File "..\OutlookMessageAllert\bin\Release\OutlookMessageAllert.dll.config"
  
  File "..\OutlookMessageAllert\bin\Release\Microsoft.Office.Tools.Common.dll"
  File "..\OutlookMessageAllert\bin\Release\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
  File "..\OutlookMessageAllert\bin\Release\Microsoft.Office.Tools.dll"
  File "..\OutlookMessageAllert\bin\Release\Microsoft.Office.Tools.Outlook.dll"
  File "..\OutlookMessageAllert\bin\Release\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll"
  File "..\OutlookMessageAllert\bin\Release\Microsoft.Office.Tools.v4.0.Framework.dll"
  File "..\OutlookMessageAllert\bin\Release\Microsoft.VisualStudio.Tools.Applications.Runtime.dll"

  ; Write the installation path into the registry
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "Install_Dir" "$INSTDIR"

  ; Write the uninstall keys for Windows
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OutlookMailAlert" "DisplayName" "Outlook Mail Alert"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OutlookMailAlert" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OutlookMailAlert" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OutlookMailAlert" "NoRepair" 1
  
  
  
  ;Create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  
  ;register addin in outlook
  WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\OutlookMessageAllert" "Description" "OutlookMessageAllert"
  WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\OutlookMessageAllert" "FriendlyName" "OutlookMessageAllert"
  WriteRegDWORD HKCU "Software\Microsoft\Office\Outlook\Addins\OutlookMessageAllert" "LoadBehavior" "3"
  ; 0x00000003 (3)
  WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\OutlookMessageAllert" "Manifest"   '"$INSTDIR\OutlookMessageAllert.vsto|vstolocal"'
  ; file:///C:/Data/work/OutlookMessageAllert/OutlookMessageAllert/bin/Release/OutlookMessageAllert.vsto|vstolocal
     
SectionEnd

;--------------------------------
;Descriptions

  ;Language strings
  LangString DESC_MainSec ${LANG_ENGLISH} "Outlook Mail Alert"

  ;Assign language strings to sections
  !insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
    !insertmacro MUI_DESCRIPTION_TEXT ${MainSec} $(DESC_MainSec)
  !insertmacro MUI_FUNCTION_DESCRIPTION_END

Section Uninstall

  Delete "$INSTDIR\OutlookMessageAllert.dll"
  Delete "$INSTDIR\OutlookMessageAllert.vsto"
  Delete "$INSTDIR\OutlookMessageAllert.dll.config"
  Delete "$INSTDIR\Microsoft.Office.Tools.Common.dll"
  Delete "$INSTDIR\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
  Delete "$INSTDIR\Microsoft.Office.Tools.dll"
  Delete "$INSTDIR\Microsoft.Office.Tools.Outlook.dll"
  Delete "$INSTDIR\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll"
  Delete "$INSTDIR\Microsoft.Office.Tools.v4.0.Framework.dll"
  Delete "$INSTDIR\Microsoft.VisualStudio.Tools.Applications.Runtime.dll"
  Delete "$INSTDIR\Uninstall.exe"

  RMDir "$INSTDIR"

  ; Remove registry keys
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OutlookMailAlert"
  DeleteRegKey HKLM "SOFTWARE\OutlookMailAlert"
  DeleteRegKey HKCR "${REG_SHELL_CLEANER_PATH}"
  DeleteRegKey HKCR "${REG_SHELL_CLEANER_COMMAND_PATH}"
  
  
SectionEnd