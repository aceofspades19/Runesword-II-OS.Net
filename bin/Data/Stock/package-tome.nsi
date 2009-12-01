; tome packaging script
;
;--------------------------------
!define PRODUCT_NAME "RuneSword: "
;!define PRODUCT_VERSION "Tome patch for 2.4.6"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\RuneSword.exe"
;!define TOMEDIR "Eternia\The Land Beyond"

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
;OutFile "${PRODUCT_VERSION}.exe"
OutFile "${LOCALDIR}\Tomes\${TOMEDIR}\${PRODUCT_VERSION}.exe"

InstallDir "$PROGRAMFILES\RuneSword"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""

;--------------------------------
; Pages

Page directory
Page instfiles

; The stuff to install
Section "" 
  ;SetOutPath $INSTDIR
  SetOverwrite ifnewer
  SetOutPath "$INSTDIR\Data"
  File /nonfatal /r "${LOCALDIR}\Tomes\${TOMEDIR}\Data\*.*"
  SetOutPath "$INSTDIR\Tomes\${TOMEDIR}"
  ;File /r /x Data "*.*"  
  !ifndef TOMENAME
	File /r /x Data "${LOCALDIR}\Tomes\${TOMEDIR}\*.*"
  !else
	File "${LOCALDIR}\Tomes\${TOMEDIR}\${TOMENAME}*.*"
  !endif
SectionEnd ; end the section