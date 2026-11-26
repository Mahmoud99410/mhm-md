; Adds "Convert to Markdown (MHM MD)" to the Windows right-click menu for supported files and folders
!macro MHM_ADD EXT
  WriteRegStr HKCU "Software\Classes\SystemFileAssociations\${EXT}\shell\MHMMD" "" "Convert to Markdown (MHM MD)"
  WriteRegStr HKCU "Software\Classes\SystemFileAssociations\${EXT}\shell\MHMMD" "Icon" "$INSTDIR\${APP_EXECUTABLE_FILENAME}"
  WriteRegStr HKCU "Software\Classes\SystemFileAssociations\${EXT}\shell\MHMMD\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
!macroend

!macro MHM_DEL EXT
  DeleteRegKey HKCU "Software\Classes\SystemFileAssociations\${EXT}\shell\MHMMD"
!macroend

!macro customInstall
  !insertmacro MHM_ADD ".pdf"
  !insertmacro MHM_ADD ".docx"
  !insertmacro MHM_ADD ".pptx"
  !insertmacro MHM_ADD ".xlsx"
  !insertmacro MHM_ADD ".txt"
  !insertmacro MHM_ADD ".csv"
  !insertmacro MHM_ADD ".png"
  !insertmacro MHM_ADD ".jpg"
  WriteRegStr HKCU "Software\Classes\Directory\shell\MHMMD" "" "Convert folder to Markdown (MHM MD)"
  WriteRegStr HKCU "Software\Classes\Directory\shell\MHMMD" "Icon" "$INSTDIR\${APP_EXECUTABLE_FILENAME}"
  WriteRegStr HKCU "Software\Classes\Directory\shell\MHMMD\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
!macroend

!macro customUnInstall
  !insertmacro MHM_DEL ".pdf"
  !insertmacro MHM_DEL ".docx"
  !insertmacro MHM_DEL ".pptx"
  !insertmacro MHM_DEL ".xlsx"
  !insertmacro MHM_DEL ".txt"
  !insertmacro MHM_DEL ".csv"
  !insertmacro MHM_DEL ".png"
  !insertmacro MHM_DEL ".jpg"
  DeleteRegKey HKCU "Software\Classes\Directory\shell\MHMMD"
!macroend
