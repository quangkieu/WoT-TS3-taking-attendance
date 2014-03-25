#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Work\Icon1.ico
#AutoIt3Wrapper_Outfile=Last24h.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_AU3Check_Parameters=-w 1
#AutoIt3Wrapper_Run_Before="C:\Users\kieuq\Desktop\software\new folder (2)\ShowOriginalLine.exe" %in%
#AutoIt3Wrapper_Run_After="C:\Users\kieuq\Desktop\software\new folder (2)\ShowOriginalLine.exe" %in%
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/sci 1
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/striponly /sci 1 /mo
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;396046237
;Ca5cad3
#AutoIt3Wrapper_run_debug_mode=Y
#Region ****#~ endfunc_comment=1
;Directives created by AutoIt3Wrapper_GUI * * * *
#~ endfunc_comment=1
#EndRegion ****#~ endfunc_comment=1
;#RequireAdmin
Opt("TrayIconDebug", 1)
#include <MsgBoxConstants.au3>
#include <JSMN.au3>
#include <IE.au3>
#include <File.au3>
#include <Inet.au3>
#include <WindowsConstants.au3>
#include <Timers.au3>
#include <Array.au3>
#include <UnixTime.au3>
#include <_HTML_Ex.au3>
#include <GUIConstantsEx.au3>
#include <_Excel_Rewrite.au3>
#include "C:\Users\kieuq\Desktop\software\new folder (2)\_AutoItErrorTrap.au3"
;#include 'TS3.au3'

Global $clan = [['A', 1000002504],['R', 1000001787],['C', 1000006228]]
Global $oIE = _IECreateEmbedded()
TCPStartup()
Local $temp = @ScriptDir & '\temp\INFO.ini'
$temp = IniReadSection($temp, 'Server')
Global $serverfile = SetServer($temp, 'xlsxfile'), $serverIP = SetServer($temp, 'ServerIP'), $QueryPort = SetServer($temp, 'queryport'), $QueryTransfer = SetServer($temp, 'querytransfer')
Global $QueryVituralServer = SetServer($temp, 'queryvituralserver'), $QueryUser = SetServer($temp, 'queryuser'), $QueryPass = SetServer($temp, 'querypass')
Global $ChannelUpload = SetServer($temp, 'channelupload'), $ChannelPass = SetServer($temp, 'channelpass'), $ChannelFolder = SetServer($temp, 'channelfolder')
Global $MainSocket = TCPConnect($serverIP, $QueryPort)
If @error Then Exit
Start2()

Func SetServer(ByRef $iniA, $c)
	Local $this = $iniA, $a[1][1]
	$c = StringLower($c)
	$this = _ArraySearch($this, $c, 0, 0, 0, 1, 0)
	If $this <> -1 Then
		$this = $iniA[$this][1]
		If $this = '' Then
			$this = SetServer($a, $c)
		ElseIf $c = 'ServerIP' Then
			TCPStartup()
			$this = TCPNameToIP($this)
			TCPShutdown()
		EndIf
	Else
		Switch $c
			Case 'xlsxfile'
				$this = @ScriptDir & '\temp\TS3.xlsx'
			Case 'serverip'
				$this = '108.92.120.202'
			Case 'queryport'
				$this = 10011
			Case 'querytransfer'
				$this = 30033
			Case 'queryvituralserver'
				$this = 1
			Case 'queryuser'
				$this = 'serveradmin'
			Case 'querypass'
				$this = 'wL5FgSWV'
			Case 'channelupload'
				$this = 6
			Case 'channelpass'
				$this = ''
			Case 'channelfolder'
				$this = '\/In-game\sreport'
		EndSwitch
	EndIf
	Return $this
EndFunc   ;==>SetServer

Func __LogWrite($file, $text, $mode)
	Local $Ram = ProcessGetStats()
	_FileWriteLog($file, Int($Ram[0] / 1048576) & "MB | " & Int($Ram[1] / 1048576) & 'MB ' & $text, $mode)
	FileClose($file)
EndFunc   ;==>__LogWrite

Func _ArrayDeleteCol(ByRef $avWork, $iCol)
	If Not IsArray($avWork) Then Return SetError(1, 0, 0); Not an array
	If UBound($avWork, 0) <> 2 Then Return SetError(1, 1, 0); Not a 2D array
	If ($iCol < 0) Or ($iCol > (UBound($avWork, 2) - 1)) Then Return SetError(1, 2, 0); $iCol out of range
	If $iCol < UBound($avWork, 2) - 1 Then
		For $c = $iCol To UBound($avWork, 2) - 2
			For $r = 0 To UBound($avWork) - 1
				$avWork[$r][$c] = $avWork[$r][$c + 1]
			Next
		Next
	EndIf
	ReDim $avWork[UBound($avWork)][UBound($avWork, 2) - 1]
	Return 1
EndFunc   ;==>_ArrayDeleteCol

Func _TS3Upload($file, $c)
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Start Upload ' & $c & '.', 1)
	Local $this = 0, $string, $string1, $temp, $sendSocket
	$this = StringReplace($file, @ScriptDir & '\temp\SNS_' & $c & '\', '')
	Sleep(3000)
	TCPSend($MainSocket, "ftinitupload cid=" & $ChannelUpload & " cpw=" & $ChannelPass & " name=" & $ChannelFolder & "\/SNS_" & $c & "\/" & $this & " overwrite=1 resume=0 clientftfid=" & _ArraySearch($clan, $c, 0, 0, 0, 0, 1, 0) + 1 & " size=" & FileGetSize($file) & @CRLF)
	;TCPSend($MainSocket,"User-Agent: Mozilla/4.0"&@CRLF&@CRLF)
	Sleep(3000)
	$string = TCPRecv($MainSocket, 1024000)
	ConsoleWrite($string)
	$string = StringReplace($string, @LF & @CR, "")
	$string = StringReplace($string, 'TS3Welcome to the TeamSpeak 3 ServerQuery interface, type "help" for a list of commands and "help <command>" ' & _
			'for information on a specific command.error id=0 msg=ok', "")
	$string = StringReplace($string, 'error id=0 msg=ok', "")
	$string = StringReplace($string, '\s', ' ')
	$string = StringReplace($string, '\/', '/')
	$string = StringReplace($string, '\p', '|')
	$string = StringReplace($string, '\\', '\')
	If $string = '' Then
		$MainSocket = TCPConnect($serverIP, $QueryPort)
		;If @error Then MsgBox(0, '', @error)
		Sleep(5000)
		TCPSend($MainSocket, "login client_login_name=" & $QueryUser & " client_login_password=" & $QueryPass & "" & @CRLF)
		TCPSend($MainSocket, "use " & $QueryVituralServer & "" & @CRLF)
		Sleep(2000)
		TCPRecv($MainSocket, 100000000)
		Return 0
	ElseIf StringInStr($string, 'ftkey=') = 0 Then
		Return 0
	EndIf
	$string = StringReplace($string, 'ftkey=', "")
	$string = StringReplace($string, 'serverftfid=', "")
	$string = StringReplace($string, '0quit', '0' & @CR & @LF)
	$string = StringSplit($string, " ")
	$temp = FileOpen($file, 0)
	$string1 = FileRead($temp)
	FileClose($temp)
	$temp = $string
	;MsgBox(0,'',$string1)
	;MsgBox(0,'',@error)
	$sendSocket = TCPConnect($serverIP, $QueryTransfer)
	Sleep(3000)
	TCPSend($sendSocket, $string[3] & $string1 & @CRLF & 'quit')
	$string = TCPRecv($sendSocket, 1024000)
	;MsgBox(0,'',$string)
	Sleep(1000)
	TCPCloseSocket($sendSocket)
	TCPSend($MainSocket, 'ftstop serverftfid=' & $temp[2] & ' delete=0' & @CRLF)
	Sleep(1000)
	$temp = TCPRecv($MainSocket, 10000)
	ConsoleWrite($temp)
	If Not (StringInStr($temp, 'msg=file\stransfer\scanceled size=0')) Then Return 0
	$temp = 0
	Return 1
EndFunc   ;==>_TS3Upload

Func InsertNewLine(ByRef $a, ByRef $b, ByRef $c)
	_Excel_RangeWrite($a, 3, _Excel_RangeRead($a, 3, _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0]) + 2) & _Excel_RangeRead($a, 3, _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0]) + 3) & '1')) + 1, _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0]) + 3) & '1')
	_Excel_RangeInsert(_Excel_SheetList($a)[2][1], _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0])) & '3:' & _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0]) + 3) & '3', $xlShiftDown)
	;If @error Then MsgBox(0, '', @error & ' ' & _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0])) & '3:' & _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0]) + 3) & '3')
EndFunc   ;==>InsertNewLine

Func Example(ByRef $c)
	Local $sendSocket, $small, $this, $temp, $begin, $temp1, $string1, $oEx, $table1, $oBook
	Global $aTableData = 0
;~ 	TCPStartup()
;~ $MainSocket = TCPConnect('162.213.61.98', '80')
;~ sleep(500)
;~ TCPSend($MainSocket,"GET /leaderboard/#wot&lb_tr=1&lb_of=battles_count&lb_cd=2014-2-4&lb_sg=clan&lb_nick=quangkieu HTTP/1.1"&@CRLF)
;~ TCPSend($MainSocket,"Host: worldoftanks.com"&@CRLF)
;~ TCPSend($MainSocket,"User-Agent: Mozilla/4.0"&@CRLF&@CRLF)
;~ Sleep(3000)
;~ $string=TCPRecv($MainSocket, 1024000)
;~ TCPCloseSocket($MainSocket)
;~ TCPShutDown()
	; *******************************************************
	; Example 1 - Open a browser with the table example, get a reference to the second table
	;				on the page (index 1) and read its contents into a 2-D array
	; *******************************************************
	;$gui=GUICreate('Check playing in last 24 hours (only SNS-A for beta)',600,300,-1,-1, $WS_OVERLAPPEDWINDOW + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
	;$warning=GUICtrlCreateLabel('This auto code would not able to be stop. It only could be minimized',10,20)
	;$lable=GUICtrlCreateLabel('Starting',10,40,300)
	;$lable1=GUICtrlCreateLabel('',10,60,300,20)
	;GUISetState(@SW_SHOW)
	$small = GUICreate('', 0, 0, 0, 0, $WS_OVERLAPPEDWINDOW + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
	$this = GUICtrlCreateObj($oIE, 0, 0, 0, 0)
	;GUISetState(@SW_SHOW)
	_IENavigate($oIE, 'about:blank')
	Sleep(3000)
	;$oIE=_IE_Example("table")
	;_HTML_GetSource('http://quangkieu:Sangngan_25111993@worldoftanks.com/leaderboard/#wot&lb_tr=1&lb_of=battles_count&lb_cd=2014-2-6&lb_sg=clan&lb_nick=quangkieu');http://cw.wotapi.ru/na/1000005063?utc=-5&type=html')
	;$oForm=_IEFormGetCollection($oIE,0)
	;$Login=_IEFormElementGetObjByName($oForm,'login')
	;$Pass=_IEFormElementGetObjByName($oForm,'password')
	;_IEFormElementSetValue($Login,'gautthn@gmail.com')
	;_IEFormElementSetValue($Pass,'Sangngan_25111993')
	;_IEFormSubmit($oForm)
	;Sleep(10000)
	;_IENavigate($oIE,'http://worldoftanks.com/')
	;Sleep(10000)
	;_IENavigate($oIE,'http://worldoftanks.com/leaderboard/#wot&lb_tr=1&lb_of=battles_count&lb_cd=2014-2-7&lb_sg=clan&lb_nick=quangkieu')
	;MsgBox(0,"",$oie)
	;Local $oTable = _IETableGetCollection($oIE, 1)
	Do
		$temp = _HTML_GetSource('http://api.worldoftanks.com/wot/clan/info/?application_id=demo&fields=members_count,members.account_name&clan_id=' & $clan[_ArraySearch($clan, $c, 0, 0, 0, 0, 1, 0)][1])
		ConsoleWrite(@error)
		$temp = Jsmn_Decode($temp)
		ConsoleWrite($temp)
		$temp = Jsmn_ObjTo2DArray($temp)
		$temp = $temp[3][1]
		$temp = $temp[1][1]
		$temp = $temp[2][1]
		$temp[0][0] = UBound($temp, 1) - 1
		For $i = 1 To $temp[0][0]
			$aTableData = $temp[$i][1]
			$temp[$i][1] = $aTableData[1][1]
		Next
		$temp[0][1] = 0
		$aTableData = $temp
		$temp = 0
		IsArray($aTableData)
		ConsoleWrite(@error)
	Until IsArray($aTableData) = 1
	;_ArrayDisplay($aTableData)
	_ArraySort($aTableData, 0, 1, 0, 1)
	;MsgBox(0,'',@error)
	;ConsoleWrite($aTableData)
	;_ArrayDisplay($aTableData)
	Local $Data[UBound($aTableData, 1)][2]
	Local $aPlayerTable = 0
	Local $i, $oPlayer = 0
	$aTableData[0][0] = UBound($aTableData, 1) - 1
	$Data[0][0] = UBound($aTableData, 1) - 1
	;GUICtrlSetData($lable,'Downloading files to temp.')
	OnAutoItExitRegister("_CleanUp")
	For $i = 1 To $Data[0][0]
		GUIDelete($small)
		If @error Then MsgBox(0, '', @error)
		$oIE = _IECreateEmbedded()
		$small = GUICreate('', 0, 0, 0, 0, $WS_OVERLAPPEDWINDOW + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
		Sleep(10)
		$this = GUICtrlCreateObj($oIE, 0, 0, 0, 0)
		GUISetState(@SW_HIDE)
		_IENavigate($oIE, 'http://www.noobmeter.com/player/na/' & $aTableData[$i][1] & '/' & $aTableData[$i][0], 0)
		Sleep(10)
	Next
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Download ' & $c & '.', 1)
	For $i = 1 To $Data[0][0]
		$Data[$i][0] = InetGet('http://www.noobmeter.com/player/na/' & $aTableData[$i][1] & '/' & $aTableData[$i][0], @ScriptDir & '\temp\' & $aTableData[$i][1] & '.html', 1, 1)
		;ConsoleWrite(@error)
		;Sleep(200)
	Next
	$temp = 1
	$begin = _Timer_Init()
	Do
		$temp = $Data[0][0]
		For $i = 1 To $Data[0][0]
			If InetGetInfo($Data[$i][0], 4) <> 0 Then
				ConsoleWrite($i & ' e;')
				$Data[$i][0] = InetGet('http://www.noobmeter.com/player/na/' & $aTableData[$i][1] & '/' & $aTableData[$i][0], @ScriptDir & '\temp\' & $aTableData[$i][1] & '.html', 1, 1)
			ElseIf Not (InetGetInfo($Data[$i][0], 2)) Then
				ConsoleWrite($i & ' d;')
			Else
				$temp = $temp - 1
			EndIf
		Next
		;GUICtrlSetData($lable1,$Data[0][0]-$temp&'/'&$Data[0][0]&' files. Next retry '&120-_Timer_Diff($begin)/1000&' seconds')
		ConsoleWrite(@CR & @LF & '====================' & @CR & @LF)
		Sleep(1000)
		If (_Timer_Diff($begin) / 1000) >= 120 Then
			;GUICtrlSetData($lable1,$Data[0][0]-$temp&'/'&$Data[0][0]&' files. Retry')
			For $i = 1 To $Data[0][0]
				If Not (InetGetInfo($Data[$i][0], 2)) Or InetGetInfo($Data[$i][0], 4) <> 0 Then
					$Data[$i][0] = InetGet('http://www.noobmeter.com/player/na/' & $aTableData[$i][1] & '/' & $aTableData[$i][0], @ScriptDir & '\temp\' & $aTableData[$i][1] & '.html', 1, 1)
					;ConsoleWrite(@error)
					Sleep(200)
				EndIf
			Next
			ConsoleWrite('start new' & @CR & @LF)
			$begin = _Timer_Init()
		EndIf
	Until $temp = 0
	For $i = 1 To $Data[0][0]
		InetClose($Data[$i][0])
	Next

	Dim $Data[1][1] = [[UBound($aTableData, 1) - 1]]
	;GUICtrlSetData($lable1,'')
	;$oIE = _IECreateEmbedded()
	;GUICtrlCreateObj($oIE, 10, 40, 600, 360)
	;_IENavigate($oIE, "about:blank")
	;$oIE=_IECreate(Default)
	;$oIE=_IEAttach('blank page','instance',1)
	$oIE.document.location.reload(True)
	Local $aMon[] = [0, 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Process files for ' & $c & '.', 1)
	For $i = 1 To $Data[0][0]
		;InetGet('http://www.noobmeter.com/player/na/'&$aTableData[$i][1]&'/'&$aTableData[$i][0],@ScriptDir&'\temp\'&$aTableData[$i][1]&'.html',1,0)
		GUIDelete($small)
		$oIE = _IECreateEmbedded()
		$small = GUICreate('', 0, 0, 0, 0, $WS_OVERLAPPEDWINDOW + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
		$this = GUICtrlCreateObj($oIE, 0, 0, 0, 0)
		GUISetState(@SW_HIDE)
		$temp = $aTableData[$i][0]
		$aTableData[$i][0] = $aTableData[$i][1]
		$aTableData[$i][1] = $temp
		ConsoleWrite($aTableData[$i][0] & @CR & @LF)
		;GUICtrlSetData($lable,'Reading temp files.')
		$temp1 = FileOpen(@ScriptDir & '\temp\' & $aTableData[$i][0] & '.html')
		$temp = FileRead($temp1)
		FileClose($temp1)
		$temp = StringRegExpReplace($temp, '[\r\n\t]', " ")
		$temp = StringRegExpReplace($temp, '(?i)<div.*?searchform.*?</div>', "")
		$temp = StringRegExpReplace($temp, '(?i)<img.*?/>', "")
		$temp = StringRegExpReplace($temp, '(?i)<table.*?>.*?</table>', "", 1)
		$temp = StringRegExpReplace($temp, '(?i)<a href=".*?">', "")
		$temp = StringRegExpReplace($temp, '(?i)</a>', "")
		$temp1 = StringRegExp($temp, '(?i)<table.*?>.*?</table>', 1)
		;_ArrayDisplay($temp1)
		If IsArray($temp1) Then $temp = StringReplace($temp, $temp1[0], 'Copy then paste here Quang', 1)
		$temp = StringRegExpReplace($temp, '(?i)<table.*?>.*?</table>', "", 1)
		If IsArray($temp1) Then $temp = StringReplace($temp, 'Copy then paste here Quang', $temp1[0], 1)
		$temp = StringRegExpReplace($temp, '(?i)<link.*?/>', "")
		$temp = StringRegExpReplace($temp, '(?i)<form.*?>.*?</form>', "")
		$temp = StringRegExpReplace($temp, '(?i)<script.*?>.*?</script>', "")
		$temp = StringRegExpReplace($temp, '(?i)<div id="pr_chart".*?>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>', "")
		;MsgBox(0,'',@error)
		$temp1 = FileOpen(@ScriptDir & '\temp\' & $aTableData[$i][0] & '.html', 2)
		FileFlush(FileWrite($temp1, $temp))
		FileClose($temp1)
		;_IENavigate($oIE,@ScriptDir&'\temp\Warlion08.html')
		_IENavigate($oIE, @ScriptDir & '\temp\' & $aTableData[$i][0] & '.html')
		$oIE.document.location.reload(True)
		;Sleep(500)
		;$oIE=_IEAttach('blank page','instance',1)
		;if @error then Exit
		If @error Then
			;GUICtrlSetData($lable1,'Error in file, re-download for '&$aTableData[$i][0])
			Do
				$temp = InetGet('http://www.noobmeter.com/player/na/' & $aTableData[$i][0] & '/' & $aTableData[$i][1], @ScriptDir & '\temp\' & $aTableData[$i][0] & '.html', 1)
			Until Not (InetGetInfo($temp, 2)) And InetGetInfo($temp, 4) <> 0
			$temp1 = FileOpen(@ScriptDir & '\temp\' & $aTableData[$i][0] & '.html')
			$temp = FileRead($temp1)
			FileClose($temp1)
			$temp = StringRegExpReplace($temp, '[\r\n\t]', " ")
			$temp = StringRegExpReplace($temp, '(?i)<div.*?searchform.*?</div>', "")
			$temp = StringRegExpReplace($temp, '(?i)<img.*?/>', "")
			$temp = StringRegExpReplace($temp, '(?i)<table.*?>.*?</table>', "", 1)
			$temp = StringRegExpReplace($temp, '(?i)<a href=".*?">', "")
			$temp = StringRegExpReplace($temp, '(?i)</a>', "")
			$temp1 = StringRegExp($temp, '(?i)<table.*?>.*?</table>', 1)
			If IsArray($temp1) Then $temp = StringReplace($temp, $temp1[0], 'Copy then paste here Quang', 1)
			$temp = StringRegExpReplace($temp, '(?i)<table.*?>.*?</table>', "", 1)
			If IsArray($temp1) Then $temp = StringReplace($temp, 'Copy then paste here Quang', $temp1[0], 1)
			$temp = StringRegExpReplace($temp, '(?i)<link.*?/>', "")
			$temp = StringRegExpReplace($temp, '(?i)<form.*?>.*?</form>', "")
			$temp = StringRegExpReplace($temp, '(?i)<script.*?>.*?</script>', "")
			$temp = StringRegExpReplace($temp, '(?i)<div id="pr_chart".*?>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>', "")
			;MsgBox(0,'',@error)
			$temp1 = FileOpen(@ScriptDir & '\temp\' & $aTableData[$i][0] & '.html', 2)
			FileFlush(FileWrite($temp1, $temp))
			FileClose($temp1)
			GUIDelete($small)
			$oIE = _IECreateEmbedded()
			$small = GUICreate('', 0, 0, 0, 0, $WS_OVERLAPPEDWINDOW + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
			$this = GUICtrlCreateObj($oIE, 0, 0, 0, 0)
			GUISetState(@SW_HIDE)
			_IENavigate($oIE, @ScriptDir & '\temp\' & $aTableData[$i][0] & '.html', 1)
			$oIE.document.location.reload(True)
			;Sleep(500)
			;$oIE=_IEAttach('blank page','instance',1)
			;$oPlayer = $temp
		EndIf
		;_ArrayDisplay($aPlayerTable,$aTableData[$i][0])
		If _HTML_Search(_IEBodyReadHTML($oIE), 'Player ' & $aTableData[$i][0] & ' has no battles on server') Then
			$aTableData[$i][1] = 0
			ConsoleWrite('\n')
		Else
			$aPlayerTable = _IETableGetCollection($oIE, 0)
			$aPlayerTable = _IETableWriteToArray($aPlayerTable, 1)
			While Not (IsArray($aPlayerTable)) Or UBound($aPlayerTable, 0) <> 2 Or UBound($aPlayerTable, 1) < 18 Or UBound($aPlayerTable, 2) < 2
				$temp = InetGet('http://www.noobmeter.com/player/na/' & $aTableData[$i][0] & '/' & $aTableData[$i][1], @ScriptDir & '\temp\' & $aTableData[$i][0] & '.html', 1)
				ConsoleWrite($aTableData[$i][0])
				$temp1 = FileOpen(@ScriptDir & '\temp\' & $aTableData[$i][0] & '.html')
				$temp = FileRead($temp1)
				FileClose($temp1)
				$temp = StringRegExpReplace($temp, '[\r\n\t]', " ")
				$temp = StringRegExpReplace($temp, '(?i)<div.*?searchform.*?</div>', "")
				$temp = StringRegExpReplace($temp, '(?i)<img.*?/>', "")
				$temp = StringRegExpReplace($temp, '(?i)<table.*?>.*?</table>', "", 1)
				$temp = StringRegExpReplace($temp, '(?i)<a href=".*?">', "")
				$temp = StringRegExpReplace($temp, '(?i)</a>', "")
				$temp1 = StringRegExp($temp, '(?i)<table.*?>.*?</table>', 1)
				If IsArray($temp1) Then $temp = StringReplace($temp, $temp1[0], 'Copy then paste here Quang', 1)
				$temp = StringRegExpReplace($temp, '(?i)<table.*?>.*?</table>', "", 1)
				If IsArray($temp1) Then $temp = StringReplace($temp, 'Copy then paste here Quang', $temp1[0], 1)
				$temp = StringRegExpReplace($temp, '(?i)<link.*?/>', "")
				$temp = StringRegExpReplace($temp, '(?i)<form.*?>.*?</form>', "")
				$temp = StringRegExpReplace($temp, '(?i)<script.*?>.*?</script>', "")
				$temp = StringRegExpReplace($temp, '(?i)<div id="pr_chart".*?>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>.*?</div>', "")
				;MsgBox(0,'',@error)
				$temp1 = FileOpen(@ScriptDir & '\temp\' & $aTableData[$i][0] & '.html', 2)
				FileFlush(FileWrite($temp1, $temp))
				FileClose($temp1)
				GUIDelete($small)
				$oIE = _IECreateEmbedded()
				$small = GUICreate('', 0, 0, 0, 0, $WS_OVERLAPPEDWINDOW + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
				$this = GUICtrlCreateObj($oIE, 0, 0, 0, 0)
				GUISetState(@SW_HIDE)
				_IENavigate($oIE, @ScriptDir & '\temp\' & $aTableData[$i][0] & '.html', 1)
				;Sleep(500)
				;GUICtrlSetData($lable1,'Error in file, re-download for '&$aTableData[$i][0])
				$oIE.document.location.reload(True)
				If _HTML_Search(_IEBodyReadHTML($oIE), 'Player ' & $aTableData[$i][0] & ' has no battles on server') Then
					$aTableData[$i][1] = 0
					ConsoleWrite('\n')
					$temp1 = 'exit'
					ExitLoop
				EndIf
				$aPlayerTable = _IETableGetCollection($oIE, 0)
				$aPlayerTable = _IETableWriteToArray($aPlayerTable, 1)
			WEnd
			If $temp1 == 'exit' Then
				$aTableData[$i][1] = -1
			ElseIf $aPlayerTable[18][0] = "Recent replays" Then
				$aTableData[$i][1] = -1
			Else
				$temp = $aPlayerTable[18][1]
				$temp = StringSplit($temp, @CR & @LF & @CR & @LF, 1)
				$temp = StringSplit($temp[1], @CR & @LF, 1)
				$temp = StringSplit($temp[$temp[0]], '-')
				$temp = StringTrimLeft($temp[2], 1)
				$temp = StringReplace($temp, ',', '')
				$temp = StringSplit($temp, ' ')
				$temp[1] = _ArraySearch($aMon, $temp[1], 0, 0, 0, 1, 1)
				;MsgBox(0,'',$aPlayerTable[5][1])
				If $temp[1] < @MON Or $temp[2] < @MDAY Or $temp[3] < @YEAR Then
					$aTableData[$i][1] = 0
				ElseIf $temp[1] >= @MON And $temp[2] >= @MDAY And $temp[3] >= @YEAR Then
					$aTableData[$i][1] = $aPlayerTable[5][2]
				Else
					;_ArrayDisplay($aPlayerTable)
				EndIf
			EndIf
		EndIf
		;Sleep(300)
	Next
	GUIDelete($small)
	_ArrayDeleteCol($aTableData, 6)
	_ArrayDeleteCol($aTableData, 5)
	_ArrayDeleteCol($aTableData, 4)
	_ArrayDeleteCol($aTableData, 3)
	_ArrayDeleteCol($aTableData, 2)



	For $i = 1 To $Data[0][0]
		FileDelete(@ScriptDir & '\temp\' & $aTableData[$i][0] & '.html')
	Next
	OnAutoItExitUnRegister("_Cleanup")

	;MsgBox(0,"",$aTableData[1][0])
	;$oTable = _IETableGetCollection($oIE, 1)
	;$aTableData = _IETableWriteToArray($oTable, True)
	If $CmdLine[0] > 1 Then ProcessWaitClose($cmdline[2])
	Do
		ObjGet("", "Excel.Application")
		$temp = @error
		Sleep(100)
	Until $temp <> 0
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Write Excel ' & $c & '.', 1)
	$oEx = _Excel_Open(False)
	$oBook = _Excel_BookOpen($oEx, $serverfile, False, True)
	ConsoleWrite(@error & ' line:' & @ScriptLineNumber & @CRLF)
	$table1 = _Excel_RangeRead($oBook, 3, 'A3:D' & _Excel_RangeRead($oBook, 3, 'A1') + 3)
	;_ArrayDisplay($table1)
	Local $starttime = _TimeMakeStamp(0, 0, 0)
	For $i = 1 To $Data[0][0]
		$temp1 = StringRegExpReplace($aTableData[$i][0], '(?i)[^a-z0-9 ]', '')
		ConsoleWrite($temp1 & @CRLF)
		$temp1 = StringRegExp($temp1, '[[:alnum:]]{3,4}', 3)[0]
		ConsoleWrite($temp1 & @CRLF)
		$temp = _ArraySearch($table1, $temp1, 1, 0, 0, 3, 1, 2)
		ConsoleWrite($temp & @CRLF)
		If $temp <> -1 Then
			$temp = $table1[$temp][0]
			ConsoleWrite(@error & ' ' & $temp & ' line:' & @ScriptLineNumber & @CRLF)
			$temp1 = _excel_ColumnToLetter(_excel_ColumnToNumber($temp) + 2) & '4'
			ConsoleWrite($temp1)
			$temp1 = _Excel_RangeRead($oBook, 3, $temp & '3:' & $temp1)
			ConsoleWrite(@error & ' line:' & @ScriptLineNumber & @CRLF)
			;_ArrayDisplay($temp1)
			If IsArray($temp1) And UBound($temp1, 0) = 2 And $aTableData[$i][1] > 0 Then
				ConsoleWrite(@error & ' line:' & @ScriptLineNumber & @CRLF)
				If ($temp1[0][0] = '' And $temp1[0][1] = '') Or ($temp1[0][1] >= $starttime And $temp1[0][0] = '') Then
					If (StringInStr($temp1[1][0], @MON & '/' & @MDAY & '/' & @YEAR) <> 0 And StringInStr($temp1[1][0], '|' & $aTableData[$i][1]) <> 0) Then
						__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Write Excel case ' & $temp & ' 1-1 [[' & $temp1[0][0] & ',' & $temp1[0][1] & '],[' & $temp1[1][0] & ',' & $temp1[1][1] & ']].', 1)
						_Excel_RangeWrite($oBook, 3, @WDAY & '|' & @MON & '/' & @MDAY & '/' & @YEAR & '|' & $aTableData[$i][1], $temp & '4')
					ElseIf (StringInStr($temp1[1][0], @MON & '/' & @MDAY & '/' & @YEAR) = 0 And $temp1[1][0] <> '') Or ($temp1[1][0] <> '' And $temp1[1][1] < $starttime) Then
						__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Write Excel case ' & $temp & ' 1-2 [[' & $temp1[0][0] & ',' & $temp1[0][1] & '],[' & $temp1[1][0] & ',' & $temp1[1][1] & ']].', 1)
						_Excel_RangeWrite($oBook, 3, @WDAY & '|' & @MON & '/' & @MDAY & '/' & @YEAR & '|' & $aTableData[$i][1], $temp & '3')
					ElseIf ($temp1[1][1] >= $starttime And $temp1[1][0] = '') Or ($temp1[0][1] >= $starttime And $temp1[0][0] = '') Then
						$k = 4
						__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Write Excel case ' & $temp & ' 1-3 [[' & $temp1[0][0] & ',' & $temp1[0][1] & '],[' & $temp1[1][0] & ',' & $temp1[1][1] & ']].', 1)
						If ($temp1[0][1] >= $starttime And $temp1[0][0] = '') Then $k = 2
						Do
							$k = $k + 1
							$temp1 = _excel_ColumnToLetter(_excel_ColumnToNumber($temp) + 2) & $k
							$temp1 = _Excel_RangeRead($oBook, 3, $temp & $k & ':' & $temp1)
						Until $temp1[0][0] <> '' Or $temp1[0][1] < $starttime Or ($temp1[0][0] = '' And $temp1[0][1] = '')
						If $temp1[0][1] < $starttime Or (StringInStr($temp1[0][0], @MON & '/' & @MDAY & '/' & @YEAR) = 0 And $temp1[0][0] <> '') Then
							__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Write Excel case ' & $temp & ' 1-3-1 [' & $temp1[0][0] & ',' & $temp1[0][1] & '].', 1)
							_Excel_RangeWrite($oBook, 3, @WDAY & '|' & @MON & '/' & @MDAY & '/' & @YEAR & '|' & $aTableData[$i][1], $temp & $k - 1)
						ElseIf (StringInStr($temp1[0][0], @MON & '/' & @MDAY & '/' & @YEAR) <> 0 And StringRegExp($temp1[0][0], '\|' & $aTableData[$i][1] & '$') = 0) Then
							__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Write Excel case ' & $temp & ' 1-3-2 [' & $temp1[0][0] & ',' & $temp1[0][1] & '].', 1)
							_Excel_RangeWrite($oBook, 3, @WDAY & '|' & @MON & '/' & @MDAY & '/' & @YEAR & '|' & $aTableData[$i][1], $temp & $k)
						EndIf

					EndIf
				ElseIf (StringInStr($temp1[0][0], @MON & '/' & @MDAY & '/' & @YEAR) <> 0 And StringRegExp($temp1[1][0], '\|' & $aTableData[$i][1] & '$') = 0) Then
					__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Write Excel case ' & $temp & ' 2 [[' & $temp1[0][0] & ',' & $temp1[0][1] & '],[' & $temp1[1][0] & ',' & $temp1[1][1] & ']].', 1)
					_Excel_RangeWrite($oBook, 3, @WDAY & '|' & @MON & '/' & @MDAY & '/' & @YEAR & '|' & $aTableData[$i][1], $temp & '3')
				ElseIf StringInStr($temp1[0][0], @MON & '/' & @MDAY & '/' & @YEAR) = 0 Or $temp1[0][1] < $starttime Then
					InsertNewLine($oBook, $table1, $temp)
					__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Write Excel case ' & $temp & ' 3 [[' & $temp1[0][0] & ',' & $temp1[0][1] & '],[' & $temp1[1][0] & ',' & $temp1[1][1] & ']].', 1)
					_Excel_RangeWrite($oBook, 3, @WDAY & '|' & @MON & '/' & @MDAY & '/' & @YEAR & '|' & $aTableData[$i][1], $temp & '3')
				EndIf
			EndIf
		EndIf
	Next
	_Excel_BookSave($oBook)
	_Excel_BookClose($oBook)
	_Excel_Close($oEx)
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Write 24h log ' & $c & '.', 1)
	$serverfile = FileOpen(@ScriptDir & '\temp\SNS_' & $c & '\SNS_' & $c & '_' & @YEAR & '-' & @MON & '-' & @MDAY & '.txt', 10)
	FileWriteLine($serverfile, '***Range of error: 5 hours***')
	;_ArrayDisplay($aTableData)
	;$temp = _ArrayFindAll($aTableData, '^[5|6|7|8|9]$|\d{2,}', 1, 0, 0, 3, 1)
	;_ArrayDisplay($temp)
	;$temp = _ArrayFindAll($aTableData, '^[1|2|3|4]$', 1, 0, 0, 3, 1)
	;_ArrayDisplay($temp)
	;$temp = _ArrayFindAll($aTableData, 0, 1, 0, 0, 0, 1)
	;_ArrayDisplay($temp)
	;$temp = _ArrayFindAll($aTableData, -1, 1, 0, 0, 0, 1)
	;_ArrayDisplay($temp)
	FileWriteLine($serverfile, 'Active players in SNS-' & $c & ' on ' & @YEAR & '-' & @MON & '-' & @MDAY & ': (' & UBound(_ArrayFindAll($aTableData, '^[5|6|7|8|9]$|\d{2,}', 1, 0, 0, 3, 1)) & ')')
	For $i = 1 To $Data[0][0]
		If $aTableData[$i][1] >= 5 Then FileWriteLine($serverfile, @TAB & @TAB & $aTableData[$i][0] & ' (' & $aTableData[$i][1] & ')')
	Next
	FileWriteLine($serverfile, 'Less than 5 battles players: (' & UBound(_ArrayFindAll($aTableData, '^[1|2|3|4]$', 1, 0, 0, 3, 1)) & ')')
	For $i = 1 To $Data[0][0]
		If $aTableData[$i][1] < 5 And $aTableData[$i][1] >= 1 Then FileWriteLine($serverfile, @TAB & @TAB & $aTableData[$i][0] & ' (' & $aTableData[$i][1] & ')')
	Next
	FileWriteLine($serverfile, 'Offline players: (' & UBound(_ArrayFindAll($aTableData, 0, 1, 0, 0, 0, 1)) & ')')
	For $i = 1 To $Data[0][0]
		If $aTableData[$i][1] = 0 Then FileWriteLine($serverfile, @TAB & @TAB & $aTableData[$i][0])
	Next
	FileWriteLine($serverfile, 'New players (check later): (' & UBound(_ArrayFindAll($aTableData, -1, 1, 0, 0, 0, 1)) & ')')
	For $i = 1 To $Data[0][0]
		If $aTableData[$i][1] = -1 Then FileWriteLine($serverfile, @TAB & @TAB & $aTableData[$i][0])
	Next
	FileFlush($serverfile)
	FileClose($serverfile)
	;_ArrayDisplay($aTableData)
	;$string=
	;$string=_HTML_GetTable('http://worldoftanks.com/leaderboard/#wot&lb_tr=1&lb_of=battles_count&lb_cd=2014-2-4&lb_sg=clan&lb_nick=quangkieu');1000002504,1000006228,1000001787
	;$string=StringSplit($string,@cr&@lf&@cr&@lf,1)
;~ $string=StringTrimLeft($string[2],3)
;~ $string=StringTrimRight($string,1)
;~ $string=StringReplace($string,@cr&@LF,'')
;~ ;MsgBox(0,"",$string[2])
;~ MsgBox(0,"",$string)
;~ ConsoleWrite($string)
;~ $string=Jsmn_Decode($string)
;~ $string=Jsmn_ObjTo2DArray($string)
	;_ArrayDisplay($string)
	;ConsoleWrite($string)
	Do
		$temp = _TS3Upload(@ScriptDir & '\temp\SNS_' & $c & '\SNS_' & $c & '_' & @YEAR & '-' & @MON & '-' & @MDAY & '.txt', $c)
	Until ($temp = 1)
	;_ArrayDisplay($aTableData)
EndFunc   ;==>Example

Func _CleanUp()
	For $i = 1 To UBound($aTableData) - 1
		FileDelete(@ScriptDir & '\temp\' & $aTableData[$i][0] & '.html')
		FileDelete(@ScriptDir & '\temp\' & $aTableData[$i][1] & '.html')
	Next
EndFunc   ;==>_CleanUp


Func Start2()
	Local $first, $second, $third, $temp = 0, $kk
	If $CmdLine[0] = 0 Then
		$kk = 'lsajdlfkasj'
		If @Compiled = 1 Then
			_AutoItErrorTrap("Error in Main Upload")
			__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Call ExcelIndex.', 1)
			Run(@ScriptDir & '\ExcelIndex.exe')
			Sleep(10000)
			__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Call Last24h A.', 1)
			$first = Run(@AutoItExe & ' A', "", @SW_HIDE, 0x8)
			Sleep(10000)
			ProcessWaitClose($first)
			__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Call Last24h R.', 1)
			$second = Run(@AutoItExe & ' R ', "", @SW_HIDE, 0x8)
			Sleep(10000)
			ProcessWaitClose($second)
			__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Call Last24h C.', 1)
			$third = Run(@AutoItExe & ' C ', "", @SW_HIDE, 0x8)
		Else
			$kk = 'A'
		EndIf
	ElseIf $CmdLine[0] >= 1 Then
		;==============================================
		;==============================================
		;SERVER!! Start Me First !!!!!!!!!!!!!!!
		;==============================================
		;==============================================
		Local $string, $sendSocket
		$kk = $cmdline[1]
		_AutoItErrorTrap("Error in Upload" & $CmdLine[1])
	EndIf
	If _ArraySearch($clan, $kk, 0, 0, 0, 0, 1, 0) = -1 Then
		TCPCloseSocket($MainSocket)
		TCPShutdown()
		Exit
	EndIf
	DirCreate(@ScriptDir & '\temp\')
	Sleep(5000)
	TCPSend($MainSocket, "login client_login_name=" & $QueryUser & " client_login_password=" & $QueryPass & "" & @CRLF)
	;If @error Then MsgBox(0, '', @error)
	TCPSend($MainSocket, "use " & $QueryVituralServer & "" & @CRLF)
	;If @error Then MsgBox(0, '', @error)
	TCPSend($MainSocket, "ftinitupload cid=" & $ChannelUpload & " cpw=" & $ChannelPass & " name=\/0 overwrite=0 resume=0 clientftfid=" & _ArraySearch($clan, $kk, 0, 0, 0, 0, 1, 0) + 1 & " size=1" & @CRLF)
	;If @error Then MsgBox(0, '', @error)
	Sleep(2000)
	$string = TCPRecv($MainSocket, 1024000)
	$string = StringRegExp($string, '(?i)serverftfid=.*? ', 1)
	If IsArray($string) Then
		$string = StringReplace($string[0], 'serverftfid=', '')
		$string = StringReplace($string, ' ', '')
		TCPSend($MainSocket, 'ftstop serverftfid=' & $string & ' delete=0' & @CRLF)
		;If @error Then MsgBox(0, '', @error)
	Else
		$sendSocket = TCPConnect($serverIP, $QueryTransfer)
		Sleep(3000)
		TCPSend($sendSocket, 'close this connectiont cause I do not need' & @CRLF)
		TCPCloseSocket($sendSocket)
	EndIf
	$string = 0
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[Last24] Start 24h ' & $kk & '.', 1)
	Example($kk)
	Sleep(3000)
	TCPSend($MainSocket, "ftinitupload cid=" & $ChannelUpload & " cpw=" & $ChannelPass & " name=\/0 overwrite=0 resume=0 clientftfid=" & _ArraySearch($clan, $kk, 0, 0, 0, 0, 1, 0) + 1 & " size=1" & @CRLF)
	Sleep(3000)
	$string = TCPRecv($MainSocket, 1024000)
	$string = StringRegExp($string, '(?i)serverftfid=.*? ', 1)
	If IsArray($string) Then
		$string = StringReplace($string[0], 'serverftfid=', '')
		$string = StringReplace($string, ' ', '')
		TCPSend($MainSocket, 'ftstop serverftfid=' & $string & ' delete=0' & @CRLF)
	Else
		$sendSocket = TCPConnect($serverIP, $QueryTransfer)
		Sleep(3000)
		TCPSend($sendSocket, 'close this connectiont cause I do not need' & @CRLF)
		TCPCloseSocket($sendSocket)
	EndIf
	TCPSend($MainSocket, 'quit' & @CRLF)
	TCPCloseSocket($MainSocket)
	TCPShutdown()
	ConsoleWrite('End ' & $kk)

EndFunc   ;==>Start2
;Example('A')
;~ ; ===================================================================
;~ ; JSON UDF's
;~ ; v0.3
;~ ;
;~ ; By: BAM5
;~ ; Last Updated: 10/07/2009
;~ ; Tested with AutoIt Version 3.3.0.0
;~ ; Extra requirements: Nothing!
;~ ;
;~ ; A set of functions that allow you to encode and decode JSON Strings
;~ ;
;~ ; Thanks to wraithdu for setting me up with some SRE. It really made
;~ ; the script a lot less messy.
;~ ;
;~ ; Comments:
;~ ;           Unfortunately I have no idea how to encode or even detect
;~ ;       multi-dimensional arrays. I wish multi-dimensional arrays in
;~ ;       Autoit were more like javascript where $array[0] would point
;~ ;       to another array so accessing the array in the array would be
;~ ;       coded like this $array[0][0]. But that's more OOP which AIT
;~ ;       hasn't really gotten into.
;~ ;
;~ ;           In order to access arrays in other arrays you must first
;~ ;       point a variable to the embeded array. Example:
;~ ;
;~ ;               Dim $EmbededArray[1]=["Look I work!"]
;~ ;               Dim $ArrayWithArrayInside[2]=[$EmbededArray, "extra"]
;~ ;               $array = $ArrayWithArrayInside[0]
;~ ;               MsgBox(0, "Hooray!", $array[0])
;~ ;
;~ ;           With the way Javascript works it would be:
;~ ;
;~ ;               Dim $EmbededArray[1]=["Look I work!"]
;~ ;               Dim $ArrayWithArrayInside[2]=[$EmbededArray, "extra"]
;~ ;               MsgBox(0, "Hooray!", $ArrayWithArrayInside[0][0])
;~ ;
;~ ;           Which is why JSON is more sensible in Javascript and other
;~ ;       languages.
;~ ;
;~ ; Main functions:
;~ ; _toJSONString - Encodes a object to a JSON String
;~ ; _fromJSONString - Creates a object from a JSON String
;~ ; ===================================================================
;~ #include <Array.au3>
; ===================================================================
; _ToJSONString($obj)
;
; Goes through an object you give it- being an array or string or
; other- and turns it into a JSON String which you can send to
; servers or save to a text file to recall information later.
;
; Parameters:
;   $obj - IN - Object to be turnd into a JSON String
; Returns:
;   JSON String or false on failure
; Errors:
;   @error = 1 - Unknown type of variable inputed
; ===================================================================
;~ Func _ToJSONString($obj)
;~     If IsArray($obj) Then
;~         $returnString = "["
;~         For $object In $obj
;~             $returnString &= _ToJSONString($object) & ","
;~         Next
;~         $returnString = StringLeft($returnString, StringLen($returnString) - 1)
;~         $returnString &= "]"
;~     ElseIf IsFloat($obj) Or IsInt($obj) Or IsBinary($obj) Then
;~         $returnString = String($obj)
;~     ElseIf IsBool($obj) Then
;~         If $obj Then
;~             $returnString = "true"
;~         Else
;~             $returnString = "false"
;~         EndIf
;~     ElseIf IsString($obj) Then
;~         $returnString = '"' & StringReplace(StringReplace(String($obj), '"', '\"'), ',', '\,') & '"'
;~     Else
;~         SetError(1)
;~         Return (False)
;~     EndIf
;~     Return ($returnString)
;~ EndFunc   ;==>_toJSONString
; ===================================================================
; _FromJSONString($str)
;
; Takes a JSON String and decodes it into program objects such as
; arrays and numbers and strings and bools.
;
; Parameters:
;   $str - IN - The JSON String to decode into objects
; Returns:
;   A object decoded from a JSON String or False on error
; Errors
;   @error = 1 - Syntax error in the JSON String or unknown variable
;   Also, if there is an error in decoding part of the string such as
;       a variable or a array, this function will replace the variable
;       or array with a string explaining the error.
; ===================================================================
;~ Func _FromJSONString($str)
;~     If StringLeft($str, 1) = '"' And StringRight($str, 1) = '"' And Not (StringRight($str, 2) = '\"') Then
;~         $obj = StringReplace(StringReplace(_StringRemoveFirstLastChar($str, '"'), '\"', '"'), '\,', ',')
;~     ElseIf $str = "true" Then
;~         $obj = True
;~     ElseIf $str = "false" Then
;~         $obj = False
;~     ElseIf StringLeft($str, 2) = "0x" Then
;~         $obj = Binary($str)
;~     ElseIf StringIsInt($str) Then
;~         $obj = Int($str)
;~     ElseIf StringIsFloat($str) Then
;~         $obj = Number($str)
;~     ElseIf StringLeft($str, 1) = '[' And StringRight($str, 1) = ']' Then
;~         $str = StringTrimLeft($str, 1)
;~         $str = StringTrimRight($str, 1)
;~         $ufelems = StringRegExp($str, "(\[.*?\]|.*?[^\\])(?:,|\z)", 3)
;~         Dim $obj[1]
;~         For $elem In $ufelems
;~             $insObj = _FromJSONString($elem)
;~             If $insObj = False And @error = 1 Then $insObj = 'Error in syntax of piece of JSONString: "' & $elem & '"'
;~             _ArrayAdd($obj, $insObj)
;~         Next
;~         _ArrayDelete($obj, 0)
;~     Else
;~         SetError(1)
;~         Return (False)
;~     EndIf
;~     Return ($obj)
;~ EndFunc   ;==>_fromJSONString
; ===================================================================
; _StringRemoveFirstLastChar($str, $char[, $firstCount=1[, $lastCount=1]])
;
; Removes characters matching $char from the beginning and end of a string
;
; Parameters:
;   $str - IN - String to search search modify and return
;   $char - IN - Character to find and delete in $str
;   $firstCount - OPTIONAL IN - How many to delete from the beginning
;   $lastCount - OPTIONAL IN - How many to delete from the end
; Returns:
;   Modified string
; Remarks:
;   Could probably be easily turned into a replace function
;                   (But I'm too lazy :P )
; ===================================================================
;~ Func _StringRemoveFirstLastChar($str, $char, $firstCount = 1, $lastCount = 1)
;~     $returnString = ""
;~     $splited = StringSplit($str, '"', 2)
;~     $count = 1
;~     $lastCount = UBound($splited) - $lastCount
;~     For $split In $splited
;~         $returnString &= $split
;~         If $count > $firstCount And $count < $lastCount Then $returnString &= '"'
;~         $count += 1
;~     Next
;~     Return ($returnString)
;~ EndFunc   ;==>_StringRemoveFirstLastChar
