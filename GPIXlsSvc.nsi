/* Инсталятор для ----------------------------------------------------
  __________________._______  ___.__           _________             
 /  _____/\______   \   \   \/  /|  |   ______/   _____/__  __ ____  
/   \  ___ |     ___/   |\     / |  |  /  ___/\_____  \\  \/ // ___\ 
\    \_\  \|    |   |   |/     \ |  |__\___ \ /        \\   /\  \___ 
 \______  /|____|   |___/___/\  \|____/____  >_______  / \_/  \___  >
        \/                    \_/          \/        \/           \/ 
------------------------------------------------------ Версии 1.5.0 */
!include "MUI2.nsh"
!include "nsDialogs.nsh"

!define APPLICATION_NAME "GPIXlsSvc"
!define APPLICATION_VERSION "1.5.0"
!define APPLICATION_FULL_NAME "Gazprom Inform Xls Parse Service"

Name "${APPLICATION_NAME}"
InstallDir "$PROGRAMFILES\GPI\${APPLICATION_NAME}"
OutFile "bin\Install\${APPLICATION_NAME}_v${APPLICATION_VERSION}_Setup.exe"
Icon "${NSISDIR}\Contrib\Graphics\Icons\orange-install.ico"
RequestExecutionLevel admin

!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\orange-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\orange-uninstall.ico"
!define MUI_ABORTWARNING

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
Page Custom pageCreate pageLeave
!insertmacro MUI_PAGE_INSTFILES
  
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "Russian"

Var HWND_SERVICE_REGISTRATION
Var HWND_SERVICE_RUNNING
Var STATE_DIALOG_CONTROL
Var STATE_SERVICE_REGISTRATION
Var STATE_SERVICE_RUNING
Var STATE_SERVICE_EXIST

Function pageCreate
	
	!insertmacro MUI_HEADER_TEXT "Настройка службы GPIXlsSvc" "Автоматический запуск и регистрация службы в системе"
	nsDialogs::Create 1018
	
		Pop $STATE_DIALOG_CONTROL
		
		${If} $STATE_DIALOG_CONTROL == error
			Abort
		${EndIf}
		
		${NSD_CreateCheckBox} 0 0 100% 12u "Регистрация службы в системе"
		Pop $HWND_SERVICE_REGISTRATION
		${NSD_Check} $HWND_SERVICE_REGISTRATION
		${NSD_CreateCheckBox} 0 20 100% 12u "Запуск службы после завершения установки"
		Pop $HWND_SERVICE_RUNNING
	
	nsDialogs::Show
	
FunctionEnd

Function pageLeave
	
	${NSD_GetState} $HWND_SERVICE_REGISTRATION $STATE_SERVICE_REGISTRATION
	${NSD_GetState} $HWND_SERVICE_RUNNING $STATE_SERVICE_RUNING
	
FunctionEnd

Section "Установка ${APPLICATION_NAME} v${APPLICATION_VERSION}" Install
	
	SetOutPath "$INSTDIR"
	File "bin\Publish\${APPLICATION_NAME}.exe"
	File "bin\Publish\${APPLICATION_NAME}.ini"
	File "bin\Publish\Readme.txt"
	WriteUninstaller "$INSTDIR\Uninstall.exe"
	${If} $STATE_SERVICE_REGISTRATION == ${BST_CHECKED}
		DetailPrint "Регистрация службы ${APPLICATION_NAME} в системе..."
		SimpleSC::InstallService "${APPLICATION_NAME}" "${APPLICATION_NAME}" "16" "2" "$INSTDIR\${APPLICATION_NAME}.exe" "" "" ""
		${If} $STATE_SERVICE_RUNING == ${BST_CHECKED}
			DetailPrint "Запуск службы ${APPLICATION_NAME}..."
			SimpleSC::StartService "${APPLICATION_NAME}" "" 30
		${EndIf}
	${EndIf}
	
SectionEnd

Section "Uninstall"
	
	SimpleSC::ExistsService "${APPLICATION_NAME}"
  	Pop $STATE_SERVICE_EXIST
	${If} $STATE_SERVICE_EXIST == 0
		SimpleSC::RemoveService "${APPLICATION_NAME}"
	${EndIf}
	RMDir /r "$INSTDIR\*.*"
	RMDir "$INSTDIR"
	
SectionEnd