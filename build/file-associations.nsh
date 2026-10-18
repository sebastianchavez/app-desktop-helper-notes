; Register Explorer commands without changing the default application for any extension.
; Input: EXT is the extension without a dot. Output: registry entries for this extension.
!macro HelperNotesRegisterExtension EXT
  WriteRegNone SHELL_CONTEXT "Software\Classes\.${EXT}\OpenWithProgids" "HelperNotes.Document"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".${EXT}" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\SystemFileAssociations\.${EXT}\shell\HelperNotes" "" "Abrir con Helper Notes"
  WriteRegStr SHELL_CONTEXT "Software\Classes\SystemFileAssociations\.${EXT}\shell\HelperNotes" "Icon" '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\",0'
  WriteRegStr SHELL_CONTEXT "Software\Classes\SystemFileAssociations\.${EXT}\shell\HelperNotes\command" "" '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\" $\"%1$\"'
!macroend

!macro HelperNotesUnregisterExtension EXT
  ; Input: EXT without a dot. Output: removes only Helper Notes entries for the extension.
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.${EXT}\OpenWithProgids" "HelperNotes.Document"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\SystemFileAssociations\.${EXT}\shell\HelperNotes"
!macroend

!macro customInstall
  ; Input: installer context and APP_EXECUTABLE_FILENAME. Output: Explorer/Open With registrations.
  WriteRegStr SHELL_CONTEXT "Software\Classes\HelperNotes.Document" "" "Helper Notes Document"
  WriteRegStr SHELL_CONTEXT "Software\Classes\HelperNotes.Document\shell\open\command" "" '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\" $\"%1$\"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}" "FriendlyAppName" "Helper Notes"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}\shell\open\command" "" '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\" $\"%1$\"'
  !insertmacro HelperNotesRegisterExtension "txt"
  !insertmacro HelperNotesRegisterExtension "html"
  !insertmacro HelperNotesRegisterExtension "htm"
  !insertmacro HelperNotesRegisterExtension "docx"
  !insertmacro HelperNotesRegisterExtension "pdf"
  !insertmacro HelperNotesRegisterExtension "xlsx"
  !insertmacro HelperNotesRegisterExtension "xls"
  !insertmacro HelperNotesRegisterExtension "csv"
  !insertmacro HelperNotesRegisterExtension "pptx"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro customUnInstall
  ; Input: uninstall context and APP_EXECUTABLE_FILENAME. Output: removes Helper Notes registrations.
  !insertmacro HelperNotesUnregisterExtension "txt"
  !insertmacro HelperNotesUnregisterExtension "html"
  !insertmacro HelperNotesUnregisterExtension "htm"
  !insertmacro HelperNotesUnregisterExtension "docx"
  !insertmacro HelperNotesUnregisterExtension "pdf"
  !insertmacro HelperNotesUnregisterExtension "xlsx"
  !insertmacro HelperNotesUnregisterExtension "xls"
  !insertmacro HelperNotesUnregisterExtension "csv"
  !insertmacro HelperNotesUnregisterExtension "pptx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\HelperNotes.Document"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
