#include-once
#include <FileConstants.au3>

; #FUNCTION# ====================================================================================================================
; Name ..........: _SelfUpdate
; Description ...: Update the running executable with another executable.
; Syntax ........: _SelfUpdate($sUpdatePath[, $fRestart = False[, $iDelay = 5[, $fUsePID = False]]])
; Parameters ....: $sUpdatePath         - Path string of the update file to overwrite the running executable.
;                  $fRestart            - [optional] Restart the application (True) or to not restart (False) after overwriting. Default is False.
;                  $iDelay              - [optional] An integer value for the delay to wait (in seconds) before stopping the process and deleting the executable.
;                                         If 0 is specified then the batch will wait indefinitely until the process no longer exits. Default is 5 (seconds).
;                  $fUsePID             - [optional] Use the process name (False) or PID (True). Default is False.
;                  $sBackupPath         - [optional] String to add to backup file. If blank or default no backup is created.
;                  $sCmdLine            - [optional] Command line arguments for restarted script
; Return values .: Success - Returns the PID of the batch file.
;                  Failure - Returns 0 & sets @error to non-zero
; Author ........: guinness
; Modified ......:
; Remarks .......: The current executable is renamed to AppName_Backup.exe if $sBackupPath is set to True.
; Example .......: Yes
; ===============================================================================================================================
Func _SelfUpdate($sUpdatePath, $fRestart = Default, $iDelay = 5, $fUsePID = Default, $sBackupPath = Default, $sCmdLine = Default)
	Local $sAppID = @ScriptName
	Local $sDelay = 'IF %TIMER% GTR ' & $iDelay & ' GOTO KILL' & @CRLF
	Local $sBackup = ''
	Local $sImageName = 'IMAGENAME'
	Local $sRestart = ''
	Local $sScriptPath = @ScriptFullPath
	Local $sTempFileName = @ScriptName

	Local $sExtension = StringRegExpReplace($sTempFileName, "(^.*)(\.[^.]+$)", "\2", 1) ; Script extension
	$sTempFileName = StringRegExpReplace($sTempFileName, "(^.*)(\.[^.]+$)", "\1", 1) ; ScriptName less the extension
	Local $sTempBatFile = $sTempFileName

	If @Compiled = 0 Or FileExists($sUpdatePath) = 0 Then
		Return SetError(1, 0, 0)
	EndIf

	If $sBackupPath <> "" And $sBackupPath <> Default Then
		$sBackupPath = StringRegExpReplace($sBackupPath, '[/?<>\\:*|\"^]+', "", 0) ; Remove illegal windows filename chars
		$sBackup = 'COPY /Y /V /D /Z /B "' & $sScriptPath & '" "' & @ScriptDir & '\' & $sTempFileName & $sBackupPath & $sExtension & '"' & @CRLF
	EndIf

	While FileExists(@TempDir & '\' & $sTempBatFile & '.bat')
		$sTempBatFile &= Chr(Random(65, 122, 1))
	WEnd
	$sTempBatFile = @TempDir & '\' & $sTempBatFile & '.bat'

	While FileExists(@TempDir & '\' & $sTempFileName & '.bak')
		$sTempFileName &= Chr(Random(65, 122, 1))
	WEnd
	$sTempFileName = @TempDir & '\' & $sTempFileName & '.bak'

	If $iDelay = Default Then
		$iDelay = 5
	EndIf

	If $iDelay = 0 Then
		$sDelay = ''
	EndIf

	If $fUsePID Then
		$sAppID = @AutoItPID
		$sImageName = 'PID'
	EndIf

	If $fRestart Then
		$sRestart = 'START "" "' & $sScriptPath & '"'
		If $sCmdLine <> "" And $sCmdLine <> Default Then
			$sRestart &= " " & $sCmdLine
		EndIf
	EndIf

	Local $sFileContent = 'SET TIMER=0' & @CRLF _
			 & 'SET TRIES=1' & @CRLF _
			 & ':START' & @CRLF _
			 & 'PING -w 1000 -n 2 127.0.0.1' & @CRLF _
			 & $sDelay _
			 & 'SET /A TIMER+=1' & @CRLF _
			 & @CRLF _
			 & 'TASKLIST /NH /FI "' & $sImageName & ' EQ ' & $sAppID & '" | FIND /I "' & $sAppID & '" && GOTO START' & @CRLF _
			 & @CRLF _
			 & ':KILL' & @CRLF _
			 & 'TASKKILL /F /FI "' & $sImageName & ' EQ ' & $sAppID & '"' & @CRLF _
			 & 'COPY /Y /V /D /Z /B "' & $sScriptPath & '" "' & $sTempFileName & '"' & @CRLF _
			 & $sBackup _
			 & @CRLF _
			 & ':COPY' & @CRLF _
			 & 'COPY /Y /V /D /Z /B "' & $sUpdatePath & '" "' & $sScriptPath & '" && GOTO RESTART' & @CRLF _
			 & 'PING -w 1000 -n 2 127.0.0.1' & @CRLF _
			 & 'SET /A TRIES+=1' & @CRLF _
			 & 'IF %TRIES% GTR 5 GOTO RECOVER' & @CRLF _
			 & 'GOTO COPY' & @CRLF _
			 & @CRLF _
			 & ':RECOVER' & @CRLF _
			 & 'COPY /Y /V /D /Z /B "' & $sTempFileName & '" "' & $sScriptPath & '"' & @CRLF _
			 & ':RESTART' & @CRLF _
			 & 'DEL "' & $sTempFileName & '"' & @CRLF _
			 & $sRestart & @CRLF _
			 & 'DEL "' & $sTempBatFile & '"' & @CRLF _
			 & 'GOTO EOF'


	Local $hFileHandle = FileOpen($sTempBatFile, $FO_OVERWRITE)
	If $hFileHandle = -1 Then
		Return SetError(2, 0, 0)
	EndIf
	FileWrite($hFileHandle, $sFileContent)
	FileClose($hFileHandle)
	Return Run($sTempBatFile, @TempDir, @SW_HIDE)
EndFunc   ;==>_SelfUpdate