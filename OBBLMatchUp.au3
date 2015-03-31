
#region initiate
	AutoItSetOption("MustDeclareVars", 1) ;Enforce good practice
	AutoItSetOption("TrayIconDebug", 1) ;0=no info, 1=debug line info
;~ 	#RequireAdmin
#endregion initiate

#region ;**** Directives created by AutoIt3Wrapper_GUI ****
	#AutoIt3Wrapper_Res_SaveSource=Y ;Include this source in compiled Executable for resilience
	#AutoIt3Wrapper_Res_Comment=
	#AutoIt3Wrapper_Icon=Icons\BB.ico
	#AutoIt3Wrapper_Res_Field = Compiled|%date% %time%
	#AutoIt3Wrapper_Res_Field = Compiled By|Tom
	#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
	#AutoIt3Wrapper_Res_Description=Automatically Upload you Cyanide games to your OBBLM server
	#AutoIt3Wrapper_Res_Fileversion=0.1.0 ;This version is used to determine if the script should update itself
	#AutoIt3Wrapper_Run_Tidy=Y ; tidy the script prior to compiling
	#Tidy_Parameters= /reel /sf /ri
	#AutoIt3Wrapper_UseX64=N ; Compile 32bit version
	#AutoIt3Wrapper_Change2CUI=n ; Create GUI rather than CUI
	#AutoIt3Wrapper_res_compatibility=Vista,Windows7 ; "Vista", "Windows7" both allowed separated by a comma.
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

#region includes
	#include <UDFs\_SelfUpdate.au3>
	#include <UDFs\_MonitorDirectory.au3>
	#include <upload.au3> ; Functions to upload

	#include <FileConstants.au3>
	#include <Constants.au3>
	#include <Misc.au3>
	#include <ButtonConstants.au3>
	#include <EditConstants.au3>
	#include <GUIConstantsEx.au3>
	#include <GUIListBox.au3>
	#include <WindowsConstants.au3>
	#include <GuiListView.au3>
	#include <IE.au3>
	#include <Inet.au3>
	#include <File.au3>
	#include <ComboConstants.au3>
	#include <StaticConstants.au3>
	#include <Crypt.au3>
	#include <WinAPI.au3>
#endregion includes

#region globals
	Global Const $bSINGLETON = _Singleton("{36AEC0A7-14A6-4D7A-AA97-D80A1469F8E9}", 1) ; Used to see if the installer is already running on machine.

	Global Const $sVERSION_INFO = "v" & _GetOwnVersion() ; Used to display script version number and check for updates

	;Below is an array of error codes and descriptions. If the Main() function returns an error, this is used to inform the user.
	Global $aERRORS[40] = [7, _
			@ScriptName & " already running", _
			"Username/Password incorrect", _
			"No response from server", _
			"Unexpected response from server", _
			"Not currently used", _
			"Unknown error with upload", _
			"User cancelled"]

	Global Const $aBB_DIR[3] = [2, "BloodBowlChaos", "BloodBowlLegendary"]
	Global Const $sMY_DOCS = _UnUNC(@MyDocumentsDir)

	Global $sBB_DIR = ""
	For $dir In $aBB_DIR
		$sBB_DIR = $sMY_DOCS & "\" & $dir
		If FileExists($sBB_DIR) Then ExitLoop
	Next
	Global Const $sBB_DIR_REPLAY = $sBB_DIR & "\Saves\Replays"
	Global $sREPLAY_FILE = ""
	Global Const $sBB_DIR_MATCH_REPORT = $sBB_DIR & "\MatchReport.sqlite"

	Global $sTEMP_FOLDER = @TempDir & "\OBBLMatchUp" ; Locations for temporary files used in the installation
	;If the Temp folder exists, add a random character and try, try again...
	While FileExists($sTEMP_FOLDER)
		$sTEMP_FOLDER &= Chr(Random(65, 122, 1))
	WEnd
	;Make sure it ends with a backslash
	If StringRight($sTEMP_FOLDER, 1) <> "\" Then $sTEMP_FOLDER &= "\"

	;Locations to check for updated versions of script
	Global Const $sUPDATE_PATH[4] = [2, "", _
			""]

	;Name of update file
	Global Const $sUPDATE_FILE = @ScriptName
	Global Const $sSERVER_SUCCESS_MESSAGE = "Upload was successful."

	Global Const $sINI_FILE = StringRegExpReplace(@ScriptName, "\.[^.]+$", "") & ".ini"

	Global $sServerName = IniRead($sINI_FILE, "SERVER", "SERVER NAME", "www.yourOBBLM.com")
	Global $iServerPort = IniRead($sINI_FILE, "SERVER", "SERVER PORT", "80")
	Global $sUsername = IniRead($sINI_FILE, "USER", "USERNAME", "")
	Global $sPassword = IniRead($sINI_FILE, "USER", "PASSWORD", "")
	Global $sReroll = IniRead($sINI_FILE, "OTHER", "REROLL", "2")

#endregion globals

; Loop through possbile update locations looking for a more recent version
For $i = 1 To $sUPDATE_PATH[0]
	_UpDateMe($sUPDATE_PATH[$i], $sUPDATE_FILE)
Next

;Run _main() and check for errors
If Not _main() Then
	Local $iError = @error
	If @extended Then $aERRORS[$iError] &= '"' & @CRLF & '"Extended error code: ' & @extended
	If $iError > $aERRORS[0] Then ; Check for error number higher than known error list
		MsgBox(48, "Error", "OBBLMatchUp returned with an unknown error")
	Else
		MsgBox(48, "Error", "OBBLMatchUp returned with status: " & $iError _ ; notify user of error
				 & @CRLF & @CRLF & '"' & $aERRORS[$iError] & '"')
	EndIf
EndIf

DirRemove($sTEMP_FOLDER, 1) ; Delete Temp folder

#region main

	Func _main()

		TraySetToolTip("OBBLMatchUp running..." & @CRLF & $sVERSION_INFO)

		If $bSINGLETON = 0 Then
			Return SetError(1, 0, 0) ;"Program already running"
		EndIf

		ProgressOn("OBBLMatchUp", "", "Copying files...", -1, -1)

		;Files
		;==============================
		DirCreate($sTEMP_FOLDER)

		FileInstall(".\bin\7za.exe", $sTEMP_FOLDER, 1)

		ProgressSet(40, "Unpacking files...")

		ProgressSet(80, "Launching...")

		ProgressSet(100, "Complete")

		ProgressOff()

		#region ### START Koda GUI section ### Form=\\isad.isadroot.ex.ac.uk\UOE\User\Desktop\frmMakeAdmin.kxf
			Local $frmOBBLMatchUp = GUICreate("OBBLMatchUp", 236, 248)
			Local $lblServer = GUICtrlCreateLabel("Server", 24, 24, 263, 17)
			Local $inServer = GUICtrlCreateInput($sServerName, 24, 40, 185, 21)
			Local $lblRerolls = GUICtrlCreateLabel("Reroll below", 128, 72, 66, 30)
			Local $cmbRerolls = GUICtrlCreateCombo($sReroll, 128, 88, 73, 25)
			GUICtrlSetData($cmbRerolls, "2|3|4|5|6")
			Local $btnCheck = GUICtrlCreateButton("Check server", 24, 80, 89, 25, $BS_DEFPUSHBUTTON)
			Local $inUser = GUICtrlCreateInput($sUsername, 24, 144, 185, 21)
			Local $lblUser = GUICtrlCreateLabel("Username", 24, 128, 58, 17)
			Local $btnUpload = GUICtrlCreateButton("Launch Uploader", 24, 184, 185, 33)
			Local $lblVersion = GUICtrlCreateLabel($sVERSION_INFO, 195, 233, 75, 25)
			GUICtrlSetColor($lblVersion, 0x716F64)
			GUISetState(@SW_SHOW, $frmOBBLMatchUp)
		#endregion ### END Koda GUI section ###

		While 1
			Local $nMsg = GUIGetMsg()
			Switch $nMsg
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_DISABLE, $frmOBBLMatchUp)
					GUISetState(@SW_HIDE, $frmOBBLMatchUp)
					Return SetError(7, 0, 0) ;"User cancelled"
				Case $btnCheck ;Query logged on users
					GUICtrlSetData($inServer, StringRegExpReplace(GUICtrlRead($inServer), "(?i)^https?://", ""))
					TCPStartup()
					Local $sServerIP = TCPNameToIP(GUICtrlRead($inServer))
					TCPShutdown()

					If ($sServerIP = "") Then
						GUICtrlSetData($lblServer, "Server (unreachable)")
						If BitAND(GUICtrlGetState($btnUpload), $GUI_ENABLE) Then GUICtrlSetState($btnUpload, $GUI_DISABLE)
					Else
;~ 						GUICtrlSetData($cmbRerolls, "")
						If BitAND(GUICtrlGetState($btnUpload), $GUI_DISABLE) Then GUICtrlSetState($btnUpload, $GUI_ENABLE)

						GUICtrlSetData($lblServer, "Server (" & $sServerIP & ")")
						GUICtrlSetState($btnUpload, $GUI_DEFBUTTON)
					EndIf

				Case $btnUpload
					Local $sPass = InputBox("Password", "Please enter your password:", $sPassword, "*M", Default, 125, Default, Default, Default, $frmOBBLMatchUp)
					Local $sUser = ""
					If $sPass Then
						If Not @error Then
							GUISetState(@SW_DISABLE, $frmOBBLMatchUp)
							GUISetState(@SW_HIDE, $frmOBBLMatchUp)
							$sUsername = GUICtrlRead($inUser)
							$sServerName = GUICtrlRead($inServer)
							$sReroll = GUICtrlRead($cmbRerolls)

							IniWrite($sINI_FILE, "SERVER", "SERVER NAME", $sServerName)
							IniWrite($sINI_FILE, "SERVER", "SERVER PORT", $iServerPort)
							IniWrite($sINI_FILE, "USER", "USERNAME", $sUsername)
							IniWrite($sINI_FILE, "USER", "PASSWORD", $sPassword)
							IniWrite($sINI_FILE, "OTHER", "REROLL", $sReroll)
							ExitLoop
						EndIf
					EndIf
;~ 				Case $cmbRerolls
			EndSwitch
		WEnd

		Local $frmLog = GUICreate("Progress Log", 731, 385, 194, 126)
		Local $lvMessages = GUICtrlCreateListView("Action|Time|Detail", 8, 8, 714, 366)
		GUISetState(@SW_SHOW, $frmLog)

		_UpdateLog($lvMessages, "Submitted", "User: " & $sUsername & "; Server: " & $sServerName & "; Reroll threshold: " & $sReroll)

		_UpdateLog($lvMessages, "Monitoring", "Checking " & $sBB_DIR_REPLAY & " for new file...")

		OnAutoItExitRegister("_Exit")
		Local $aDirs[1]
		$aDirs[0] = $sBB_DIR_REPLAY
		_MonitorDirectory($aDirs, False, 250, "_ReportFileChanges")
		While ($sREPLAY_FILE = "")
			Sleep(20)
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_DISABLE, $frmLog)
					GUISetState(@SW_HIDE, $frmLog)
					Return SetError(7, 0, 0) ;"User cancelled"
			EndSwitch
		WEnd

		_UpdateLog($lvMessages, "Found", "File " & $sREPLAY_FILE & " found.")

		_UpdateLog($lvMessages, "Monitoring", "Checking " & $sBB_DIR_MATCH_REPORT & " for changes...")

		_ChangeFileMon($sBB_DIR_MATCH_REPORT)
		If @error Then
			GUISetState(@SW_DISABLE, $frmLog)
			GUISetState(@SW_HIDE, $frmLog)
			Return SetError(7, 0, 0) ;"User cancelled"
		EndIf

		_UpdateLog($lvMessages, "Change", $sBB_DIR_MATCH_REPORT & " changed")

		Local $sUploadFile = "MatchReport_" & StringLeft(StringRight($sREPLAY_FILE, 22), 19) & ".zip"
		_UpdateLog($lvMessages, "Zipping", $sUploadFile)
		If FileExists($sUploadFile) Then FileDelete($sUploadFile)

		While _FileInUse($sREPLAY_FILE)
			Sleep(200)
		WEnd
		ShellExecuteWait($sTEMP_FOLDER & "7za.exe", 'a -tzip "' & $sUploadFile & '" "' & $sREPLAY_FILE & '"', $sBB_DIR, "open", @SW_HIDE)
		_UpdateLog($lvMessages, "Zipping", $sREPLAY_FILE)
		While _FileInUse($sBB_DIR_MATCH_REPORT)
			Sleep(200)
		WEnd
		ShellExecuteWait($sTEMP_FOLDER & "7za.exe", 'a -tzip "' & $sUploadFile & '" "' & $sBB_DIR_MATCH_REPORT & '"', $sBB_DIR, "open", @SW_HIDE)
		_UpdateLog($lvMessages, "Zipping", $sBB_DIR_MATCH_REPORT)

		$sUploadFile = $sBB_DIR & "\" & $sUploadFile

		_UpdateLog($lvMessages, "Upload", "Uploading " & $sUploadFile)
		Local $sUploadResult = _OBBLMUpload($sUsername, $sPass, $sServerName, $sUploadFile, $iServerPort, $sReroll)

		If Not $sUploadResult Then
			Local $iError = @error
			Switch $iError
				Case 0
				Case 1
					Return SetError(2, $iError, 0) ; Wrong username/password
				Case 2
					Return SetError(3, $iError, 0) ; No response
				Case 3
					Return SetError(4, $iError, 0) ; Unexpected response
				Case Else
					Return SetError(6, $iError, 0) ; Unknown error with upload
			EndSwitch
		Else
			_UpdateLog($lvMessages, "Response", 'Server response: "' & $sUploadResult & '"')
			If $sUploadResult <> $sSERVER_SUCCESS_MESSAGE Then
				_UpdateLog($lvMessages, "Warning", "Unexpected server response")
				MsgBox(64, "Unexpected server response", $sUploadResult)
			EndIf
		EndIf

		_UpdateLog($lvMessages, "Finished", "Close this window when ready")
		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_DISABLE, $frmLog)
					GUISetState(@SW_HIDE, $frmLog)
					ExitLoop
			EndSwitch
		WEnd

		Return SetError(0, 0, 1)
	EndFunc   ;==>_main
#endregion main

#region functions

	Func _ChangeFileMon($sFile)
		_Crypt_Startup()
		Local $sModTime = FileGetTime($sFile, 0, 1)
		Local $sHash = _Crypt_HashFile($sFile, $CALG_MD5)
		While 1
			If FileGetTime($sFile, 0, 1) <> $sModTime Then
				If _Crypt_HashFile($sFile, $CALG_MD5) <> $sHash Then ExitLoop
				$sModTime = FileGetTime($sFile, 0, 1)
			EndIf
			Sleep(60)
			Local $nMsg = GUIGetMsg()
			Switch $nMsg
				Case $GUI_EVENT_CLOSE
					_Crypt_Shutdown()
					Return SetError(1, 0, 0)
			EndSwitch
		WEnd
		_Crypt_Shutdown()
		Return SetError(0, 0, $sFile)
	EndFunc   ;==>_ChangeFileMon

;~ Stop monitoring before exiting at least
	Func _Exit()
		_MonitorDirectory()
		Exit
	EndFunc   ;==>_Exit

	Func _FileInUse($sFilePath) ; By Nessie. Modified by guinness. http://www.autoitscript.com/forum/topic/153060-fileinuse-to-check-if-file-is-in-used-or-not-work-on-both-local-and-network-drive/?p=1100491
		Local $iAccess = $FILE_SHARE_WRITE
		If DriveGetType($sFilePath) = 'NETWORK' Then $iAccess = $FILE_SHARE_READ
		Local Const $hFileOpen = _WinAPI_CreateFile($sFilePath, $CREATE_ALWAYS, $iAccess)
		Local $fReturn = True
		If $hFileOpen Then
			_WinAPI_CloseHandle($hFileOpen)
			$fReturn = False
		EndIf

		If $fReturn Then
			$fReturn = _WinAPI_GetLastError() = 32 ; $ERROR_SHARING_VIOLATION The process cannot access the file because it is being used by another process.
		EndIf
		Return $fReturn
	EndFunc   ;==>_FileInUse

	Func _GetOwnVersion()
		Local $sVersion = FileGetVersion(@ScriptFullPath)
		If Not @Compiled Then
			Local $sScriptSource = FileRead(@ScriptFullPath)
			Local $aVersion = StringRegExp($sScriptSource, "#AutoIt3Wrapper_Res_Fileversion=(?: [ ]*)?([0-9.]+)", 1)
			If IsArray($aVersion) Then
				$sVersion = $aVersion[0]
			Else
				Return SetError(1, 0, "")
			EndIf
		EndIf
		Return $sVersion
	EndFunc   ;==>_GetOwnVersion

	Func _NewFileMon($sDir, $sPattern = ".*")
		Local $sFile = ""
		Local $strComputer = "."
		Local $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\cimv2")

		Local $sSelect = "SELECT * FROM __InstanceCreationEvent WITHIN 5 WHERE " _
				 & "Targetinstance ISA ""CIM_DirectoryContainsFile"" and " _
				 & "TargetInstance.GroupComponent=" _
				 & '"Win32_Directory.Name=\"' & StringReplace($sDir, "\", "\\\\") & '\""'

		Local $colMonitoredEvents = $objWMIService.ExecNotificationQuery($sSelect)

		While 1
			Local $objEventObject = $colMonitoredEvents.NextEvent()

			If $objEventObject.Path_.Class() = "__InstanceCreationEvent" Then
				$sFile = StringRegExpReplace($objEventObject.TargetInstance.PartComponent(), '^[^"]+\.Name="([^"]+)"$', "\1")
				$sFile = StringReplace($sFile, '\\', '\')
				If StringRegExp($sFile, $sPattern) And FileExists($sFile) Then ExitLoop
			EndIf
		WEnd
		Return $sFile
	EndFunc   ;==>_NewFileMon

;~ This function will be called by _MonitorDirectory when changes are registered
	Func _ReportFileChanges($Action, $FilePath)

		ConsoleWrite($Action & ": " & $FilePath & @CRLF)
		If $Action = "Created" And StringRegExp($FilePath, "^.*Replay_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.db$") Then
			$sREPLAY_FILE = $FilePath
		EndIf

	EndFunc   ;==>_ReportFileChanges

	Func _TimeStamp()
		Local $hTime = _Date_Time_GetSystemTime()
		Return _Date_Time_SystemTimeToDateTimeStr($hTime)
	EndFunc   ;==>_TimeStamp

	Func _UnUNC($Path = @MyDocumentsDir)
		Local $sDriveLetter = ""
		Local $sDriveUnc = ""
		For $i = 65 To 90
;~ 			$sDriveLetter = Chr($i) & ":"
			$sDriveUnc = DriveMapGet($sDriveLetter)
			If $sDriveUnc Then
				$Path = StringReplace($Path, $sDriveUnc, $sDriveLetter, 1)
			EndIf
		Next
		Return $Path
	EndFunc   ;==>_UnUNC

	Func _UpdateLog($hHandle, $sMessage, $sDesc = "")
		Local $iItem = _GUICtrlListView_InsertItem($hHandle, $sMessage)
		_GUICtrlListView_AddSubItem($hHandle, $iItem, _TimeStamp(), 1)
		_GUICtrlListView_AddSubItem($hHandle, $iItem, $sDesc, 2)
		For $i = 0 To _GUICtrlListView_GetColumnCount($hHandle) - 1
			_GUICtrlListView_SetColumnWidth($hHandle, $i, $LVSCW_AUTOSIZE_USEHEADER)
		Next
	EndFunc   ;==>_UpdateLog

	Func _UpDateMe($sUpdatePath, $sUpdateFile)
		If StringRight($sUpdatePath, 1) <> "\" Then $sUpdatePath &= "\"
		If @Compiled Then
			If FileExists($sUpdatePath & $sUpdateFile) Then
				Local $sRemoteVer = FileGetVersion($sUpdatePath & $sUpdateFile)
				If FileGetVersion(@ScriptFullPath) < FileGetVersion($sUpdatePath & $sUpdateFile) Then
					TrayTip("OBBLMatchUp Update", "Updating from " & FileGetVersion(@ScriptFullPath) & " to " & $sRemoteVer, 5, 2)
					Sleep(5000)
					_SelfUpdate($sUpdatePath & $sUpdateFile, True, 0, False, Default, $CMDLINERAW)
					If Not @error Then Exit
				EndIf
			EndIf
		EndIf
		Return SetError(0, 0, 1)
	EndFunc   ;==>_UpDateMe
#endregion functions

