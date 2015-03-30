
#include-once

;~ =========================== FUNCTION _MonitorDirectory() ==============================
#cs
	Description:     Monitors the user defined directories for file activity.
	Original:        http://www.autoitscript.com/forum/index.php?showtopic=69044&hl=folderspy&st=0
	Modified:        Jack Chen
	Syntax:         _MonitorDirectory($Dirs = "", $Subtree = True, $TimerMs = 250, $Function = "_ReportChanges")
	Parameters:
	$Dirs          - Optional: Zero-based array of valid directories to be monitored.
	$Subtree       - Optional: Subtrees will be monitored if $Subtree = True.
	$TimerMs       - Optional: Timer to register changes in milliseconds.
	$Function      - Optional: Function to launch when changes are registered. e.g. _ReportChanges
	Syntax of your function must be e.g._ReportChanges($Action, $FilePath)
	Possible actions: Created, Deleted, Modified, Rename-, Rename+, Unknown
	Remarks:        Call _MonitorDirectory() without parameters to stop monitoring all directories.
	THIS SHOULD BE DONE BEFORE EXITING SCRIPT AT LEAST.
#ce

Func _MonitorDirectory($Dirs = "", $Subtree = True, $TimerMs = 250, $Function = "_ReportChanges")
	Local Static $nMax, $hBuffer, $hEvents, $aSubtree, $sFunction
	If IsArray($Dirs) Then
		$nMax = UBound($Dirs)
	ElseIf $nMax < 1 Then
		Return
	EndIf
	Local Static $aDirHandles[$nMax], $aOverlapped[$nMax], $aDirs[$nMax]
	If IsArray($Dirs) Then
		$aDirs = $Dirs
		$aSubtree = $Subtree
		$sFunction = $Function
;~      $hBuffer = DllStructCreate("byte[4096]")
		$hBuffer = DllStructCreate("byte[65536]")
		For $i = 0 To $nMax - 1
			If StringRight($aDirs[$i], 1) <> "\" Then $aDirs[$i] &= "\"
;~  http://msdn.microsoft.com/en-us/library/aa363858%28VS.85%29.aspx
			Local $aResult = DllCall("kernel32.dll", "hwnd", "CreateFile", "Str", $aDirs[$i], "Int", 0x1, "Int", BitOR(0x1, 0x4, 0x2), "ptr", 0, "int", 0x3, "int", BitOR(0x2000000, 0x40000000), "int", 0)
			$aDirHandles[$i] = $aResult[0]
			$aOverlapped[$i] = DllStructCreate("ulong_ptr Internal;ulong_ptr InternalHigh;dword Offset;dword OffsetHigh;handle hEvent")
			For $j = 1 To 5
				DllStructSetData($aOverlapped, $j, 0)
			Next
			_SetReadDirectory($aDirHandles[$i], $hBuffer, $aOverlapped[$i], True, $aSubtree)
		Next
		$hEvents = DllStructCreate("hwnd hEvent[" & UBound($aOverlapped) & "]")
		For $j = 1 To UBound($aOverlapped)
			DllStructSetData($hEvents, "hEvent", DllStructGetData($aOverlapped[$j - 1], "hEvent"), $j)
		Next
		AdlibRegister("_GetChanges", $TimerMs)
	ElseIf $Dirs = "ReadDirChanges" Then
		_GetDirectoryChanges($aDirHandles, $hBuffer, $aOverlapped, $hEvents, $aDirs, $aSubtree, $sFunction)
	ElseIf $Dirs = "" Then
		AdlibUnRegister("_GetChanges")
;~  Close Handle
		For $i = 0 To $nMax - 1
			DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $aDirHandles[$i])
			DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $aOverlapped[$i])
		Next
		DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hEvents)
		$nMax = 0
		$hBuffer = ""
		$hEvents = ""
		$aDirHandles = ""
		$aOverlapped = ""
		$aDirs = ""
		$aSubtree = ""
		$sFunction = ""
	EndIf
EndFunc   ;==>_MonitorDirectory

Func _SetReadDirectory($hDir, $hBuffer, $hOverlapped, $bInitial, $bSubtree)
	Local $hEvent, $pBuffer, $nBufferLength, $pOverlapped
	$pBuffer = DllStructGetPtr($hBuffer)
	$nBufferLength = DllStructGetSize($hBuffer)
	$pOverlapped = DllStructGetPtr($hOverlapped)
	If $bInitial Then
		$hEvent = DllCall("kernel32.dll", "hwnd", "CreateEvent", "UInt", 0, "Int", True, "Int", False, "UInt", 0)
		DllStructSetData($hOverlapped, "hEvent", $hEvent[0])
	EndIf
;~  http://msdn.microsoft.com/en-us/library/aa365465%28VS.85%29.aspx
	Local $aResult = DllCall("kernel32.dll", "Int", "ReadDirectoryChangesW", "hwnd", $hDir, "ptr", $pBuffer, "dword", $nBufferLength, "int", $bSubtree, "dword", BitOR(0x1, 0x2, 0x4, 0x8, 0x10, 0x40, 0x100), "Uint", 0, "Uint", $pOverlapped, "Uint", 0)
	Return $aResult[0]
EndFunc   ;==>_SetReadDirectory

Func _GetChanges()
	_MonitorDirectory("ReadDirChanges")
EndFunc   ;==>_GetChanges

Func _GetDirectoryChanges($aDirHandles, $hBuffer, $aOverlapped, $hEvents, $aDirs, $aSubtree, $sFunction)
	Local $aMsg, $i, $nBytes = 0
	$aMsg = DllCall("User32.dll", "dword", "MsgWaitForMultipleObjectsEx", "dword", UBound($aOverlapped), "ptr", DllStructGetPtr($hEvents), "dword", -1, "dword", 0x4FF, "dword", 0x6)
	$i = $aMsg[0]
	Switch $i
		Case 0 To UBound($aDirHandles) - 1
			DllCall("Kernel32.dll", "Uint", "ResetEvent", "uint", DllStructGetData($aOverlapped[$i], "hEvent"))
			_ParseFileMessages($hBuffer, $aDirs[$i], $sFunction)
			_SetReadDirectory($aDirHandles[$i], $hBuffer, $aOverlapped[$i], False, $aSubtree)
			Return $nBytes
	EndSwitch
	Return 0
EndFunc   ;==>_GetDirectoryChanges

Func _ParseFileMessages($hBuffer, $sDir, $sFunction)
	Local $hFileNameInfo, $pBuffer, $FilePath
	Local $nFileNameInfoOffset = 12, $nOffset = 0, $nNext = 1
	$pBuffer = DllStructGetPtr($hBuffer)
	While $nNext <> 0
		$hFileNameInfo = DllStructCreate("dword NextEntryOffset;dword Action;dword FileNameLength", $pBuffer + $nOffset)
		Local $hFileName = DllStructCreate("wchar FileName[" & DllStructGetData($hFileNameInfo, "FileNameLength") / 2 & "]", $pBuffer + $nOffset + $nFileNameInfoOffset)
		Local $Action = ""
		Switch DllStructGetData($hFileNameInfo, "Action")
			Case 0x1 ; $FILE_ACTION_ADDED
				$Action = "Created"
			Case 0x2 ; $FILE_ACTION_REMOVED
				$Action = "Deleted"
			Case 0x3 ; $FILE_ACTION_MODIFIED
				$Action = "Modified"
			Case 0x4 ; $FILE_ACTION_RENAMED_OLD_NAME
				$Action = "Rename-"
			Case 0x5 ; $FILE_ACTION_RENAMED_NEW_NAME
				$Action = "Rename+"
			Case Else
				$Action = "Unknown"
		EndSwitch

		$FilePath = $sDir & DllStructGetData($hFileName, "FileName")
		Call($sFunction, $Action, $FilePath) ; Launch the specified function
		$nNext = DllStructGetData($hFileNameInfo, "NextEntryOffset")
		$nOffset += $nNext
	WEnd
EndFunc   ;==>_ParseFileMessages
;~ ===========================End of FUNCTION _MonitorDirectory() ==============================
