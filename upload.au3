

#include-once
#include <UDFs\WinHttp.au3>

; #FUNCTION# ====================================================================================================================
; Name ..........: _OBBLMlogin
; Description ...: Update the running executable with another executable.
; Syntax ........: _OBBLMlogin($sUsername, $sPassword , $sUrl [, $iPort = False])
; Parameters ....: $sUsername           - Username (Coach name) on OBBLM.
;                  $sPassword           - Password
;                  $sUrl                - OBBLM website
;                  $iPort               - [optional] Port to use. Default is 80 (HTTP)
; Return values .: Success - Returns PHPSESSID cookie string.
;                  Failure - Returns 0 & sets @error to non-zero
; Author ........: TomG
; Modified ......:
; Remarks .......:
; Example .......:
; ===============================================================================================================================

Func _OBBLMlogin($sUsername, $sPassword, $sUrl, $iPort = 80)

	If StringLeft($sUrl, 7) <> "http://" Then $sUrl = "http://" & $sUrl

;~ Local $Cookie = "obblmpasswd=deleted;obblmuserid=deleted;PHPSESSID="
	Local $sSessionID = ""
	Local $sPD = "coach=" & $sUsername & "&passwd=" & $sPassword & "&remember=1&login=Login"
	Local $sHeaderContentType = "application/x-www-form-urlencoded"

	Local $oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
	$oHTTP.Open("POST", $sUrl)
	$oHTTP.SetRequestHeader("Content-Type", $sHeaderContentType)

	$oHTTP.Send($sPD)
	Local $HeaderResponses = $oHTTP.GetAllResponseHeaders()

;~ Local $oReceived = $oHTTP.ResponseText
	Local $oStatusCode = $oHTTP.Status

	If StringInStr($HeaderResponses, "obblmuserid=deleted;") Then Return SetError(1, 0, 0)

	If $oStatusCode <> 200 Then
		Return SetError(2, 0, 0)
	EndIf


	; Handle Cookies
	Local $aCookies = StringRegExp($HeaderResponses, 'Set-Cookie: (.+)\r\n', 3)

	If Not IsArray($aCookies) Then
		Return SetError(3, 0, 0)
	EndIf

	For $sCookie In $aCookies
		Local $aSessionID = StringRegExp($sCookie, "PHPSESSID=([a-z0-9]+;).+", 1)
		If IsArray($aSessionID) Then $sSessionID = $aSessionID[0]
	Next

	If StringLen($sSessionID) < 1 Then Return SetError(4, 0, 0)


	Return SetError(0, 0, $sSessionID)

EndFunc   ;==>_OBBLMlogin




#cs
Global Const $sUserName = "Tom"
Global Const $sPassword = "pw"
Global Const $sDomain = "www.knobbl.com"
Global Const $sReroll = 5
Global Const $sUploadFile = "C:\Folder\Match_2015-03-19_21-11-04.zip"

Local $RES = _OBBLMUpload($sUserName, $sPassword, $sDomain, $sUploadFile, $sReroll)
#ce


Func _OBBLMUpload($sUserName, $sPassword, $sDomain, $sUploadFile, $sPort = 80, $sReroll = 2)
	If $sPort = Default Then $sPort = 80
	If $sReroll = Default Then $sReroll = 2

	Local Const $sLoginPage = "index.php"
	Local Const $sUploadPage = "handler.php?type=leegmgr"
	Local Const $sLoginFailStr = "obblmuserid=deleted;"
	Local Const $sExpectReturn = '<li><a href="index.php?section=about">OBBLM</a></li>'
	Local Const $sLoginData = "coach=" & $sUserName & "&passwd=" & $sPassword & "&remember=1&login=Login"
	Local Const $sBoundary = "-----------------------------7da1c51404c4"
	Local $hRequest = ""
	Local $sReturned = ""
	Local $sHeader = ""
	Local Const $sMatchData =	$sBoundary & @CRLF & _
                    'Content-Disposition: form-data; name="MAX_FILE_SIZE"' & _
                    @CRLF & @CRLF & _
                    "256000" & @CRLF & _
                    $sBoundary & @CRLF & _
                    'Content-Disposition: form-data; name="userfile"; filename="' & $sUploadFile & '"' & _
                    @CRLF & _
                    "Content-Type: application/x-zip-compressed" & @CRLF & _
                    @CRLF & _
					FileRead($sUploadFile) & _
					@CRLF & $sBoundary & @CRLF & _
                    'Content-Disposition: form-data; name="ffatours"' & @CRLF & _
                    @CRLF & _
                    "1" & @CRLF & _
                    $sBoundary & @CRLF & _
                    'Content-Disposition: form-data; name="reroll"' & @CRLF & _
                    @CRLF & _
                    $sReroll & @CRLF & _
                    $sBoundary & "--" & _
                    @CRLF


	; Initialize and get session handle
	Local Const $hOpen = _WinHttpOpen("Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.2.6) Gecko/20100625 Firefox/3.6.6")
	; Get connection handle
	Local Const $hConnect = _WinHttpConnect($hOpen, $sDomain, $sPort)
	; Make a request
	$hRequest = _WinHttpOpenRequest($hConnect, "POST", $sLoginPage)

	; Send login.
	_WinHttpSendRequest($hRequest, "Content-Type: application/x-www-form-urlencoded", $sLoginData)
	; Wait for the response
	_WinHttpReceiveResponse($hRequest)

	; See what's returned
	If _WinHttpQueryDataAvailable($hRequest) Then ; if there is data
		Do
			$sReturned &= _WinHttpReadData($hRequest)
		Until @error
	EndIf

	If Not $sReturned Then
		;Close other handles
		_WinHttpCloseHandle($hConnect)
		_WinHttpCloseHandle($hOpen)
		Return SetError(2, 0, 0)
	EndIf

	If Not StringInStr($sReturned, $sExpectReturn) Then
		;Close other handles
		_WinHttpCloseHandle($hConnect)
		_WinHttpCloseHandle($hOpen)
		Return SetError(3, 0, 0)
	EndIf

	$sHeader = _WinHttpQueryHeaders($hRequest)

	; Close handles
	_WinHttpCloseHandle($hRequest)

	If StringInStr($sHeader, $sLoginFailStr) Then
		;Close other handles
		_WinHttpCloseHandle($hConnect)
		_WinHttpCloseHandle($hOpen)
		Return SetError(1, 0, 0)
	EndIf

	$hRequest = _WinHttpOpenRequest($hConnect, "POST", $sUploadPage)
	_WinHttpSendRequest($hRequest, "Content-Type: multipart/form-data; boundary=" & StringTrimLeft($sBoundary,2), Binary($sMatchData))
	_WinHttpReceiveResponse($hRequest)

	; See what's returned
	$sReturned = ""
	If _WinHttpQueryDataAvailable($hRequest) Then ; if there is data
		Do
			$sReturned &= _WinHttpReadData($hRequest)
		Until @error
	EndIf
;~ 	FileWrite("log.txt",$sReturned)

Local $sMarker = '<div class="section"> <!-- This container holds the section specific content -->'
$sReturned = StringRegExpReplace($sReturned,"(?s).*" & $sMarker & "\s*(?:<[^>]+>)*(.*?)(?:<|\r|\n).*","$1")
$sReturned = StringLeft(StringRegExpReplace($sReturned,"(?m)(?s)[^\w. :]",""),80)
If Not $sReturned Then $sReturned = "Unable to parse server response"

	; Get full header
	$sHeader = _WinHttpQueryHeaders($hRequest)

	; Close request handle
	_WinHttpCloseHandle($hRequest)

	;Close other handles
	_WinHttpCloseHandle($hConnect)
	_WinHttpCloseHandle($hOpen)

	Return SetError(0, 0, $sReturned)

EndFunc   ;==>_OBBLMUpload
