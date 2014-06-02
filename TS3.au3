#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Work\Icon1.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_AU3Check_Parameters=-w 1
#AutoIt3Wrapper_Run_Before=".\software\new folder (2)\ShowOriginalLine.exe" %in%
#AutoIt3Wrapper_Run_After=".\software\new folder (2)\ShowOriginalLine.exe" %in%
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/sci 1
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/striponly /sci 1 /mo
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#Region ****#~ endfunc_comment=1
;Directives created by AutoIt3Wrapper_GUI * * * *
#~ endfunc_comment=1
#AutoIt3Wrapper_run_debug_mode=Y
#EndRegion ****#~ endfunc_comment=1
#AutoIt3Wrapper_Run_Debug=N
;#RequireAdmin
#include <WinAPIShPath.au3>
;#include <array.au3>
If $CmdLine[0] = 3 Then
	Opt("TrayIconHide", 1)
	If StringRegExp(_WinAPI_CommandLineToArgv($CmdLineRaw)[1], 'http:\/\/api.worldoftanks.*?\/wot\/clan\/battles\/\?application_id=8b379c513ce91c4e9ddfe538a12b0e67&clan_id=') = 0 Then Exit
	Local $temp = FileOpen(_WinAPI_CommandLineToArgv($CmdLineRaw)[3] & '\inet_temp.html', 10)
	If @error Or $temp = -1 Then Exit
	FileWrite($temp, InetRead(_WinAPI_CommandLineToArgv($CmdLineRaw)[1], _WinAPI_CommandLineToArgv($CmdLineRaw)[2]))
	FileClose($temp)
	;MsgBox(0, '', FileRead(_WinAPI_CommandLineToArgv($CmdLineRaw)[3] & '\inet_temp.html'))
	$temp = 0
	Opt("TrayIconHide", 1)
	Exit
EndIf

#include <Process.au3>
; Quit with message box if app is already running
$numberOfAppInstances = ProcessList(_ProcessGetName(@AutoItPID))
$numberOfAppInstances = $numberOfAppInstances[0][0]
If $numberOfAppInstances > 1 Then
	If $CmdLine[0] <> 3 Then MsgBox(0, "Exiting", "Application already running!", 180)
	Opt("TrayIconHide", 1)
	Exit
EndIf
#include <String.au3>
#include <File.au3>
#include <Inet.au3>
#include <_Array2D.au3>
#include <Date.au3>
#include <_HTML_Ex.au3>
#include <UnixTime.au3>
#include <_Excel_Rewrite.au3>
#include <Timers.au3>
#include ".\software\new folder (2)\_AutoItErrorTrap.au3"
;Opt("TrayIconDebug", 1)
;$file="D:/quang1.txt"
$string = 0
;FileCopy("D:/quang1.txt","D:/quang2.txt",1)

Global $PassUser[1][3] = [[0, 0, 0]], $StampTime

_AutoItErrorTrap("Error in Main TS3")

Do
	$temp = @ScriptDir & '\temp\INFO.ini'
	$temp = IniReadSection($temp, 'Server')
	If @error Then
		FileInstall('.\temp\INFO - Copy.ini', @ScriptDir & '\temp\INFO.ini', 1)
	EndIf
Until IsArray($temp)
Global $Start = TimerInit(), $CWtime = 0, $ManualUpdate = 0, $CheckINI = 0;$clan = [['A', 1000002504],['R', 1000001787],['C', 1000006228]],
CheckINI()
Func CheckINI()
	TrayItemSetState($CheckINI, 68)
	Global $iniA = @ScriptDir & '\temp\INFO.ini', $StampTime = _TimeMakeStamp(), $serverfile = SetServer('xlsxfile'), $serverIP = SetServer('ServerIP'), $QueryPort = SetServer('queryport'), $QueryTransfer = SetServer('querytransfer')
	Global $QueryVituralServer = SetServer('queryvituralserver'), $QueryUser = SetServer('queryuser'), $QueryPass = SetServer('querypass'), $CopyTo = SetServer('copyto'), $SleepPeriod = SetServer('sleepperiod')


	$temp = IniReadSection(@ScriptDir & '\temp\INFO.ini', 'PassUsers')
	If @error Then
		Global $PassUser[1][3] = [[0, 0, 0]]
	Else
		Local $temp1[$temp[0][0]][3]
		$temp1[0][0] = $temp[0][0] - 1
		$temp1[0][1] = $temp[1][1]
		For $i = 2 To $temp[0][0]
			$temp1[$i - 1][2] = $temp[$i][0]
			$temp1[$i - 1][0] = StringSplit($temp[$i][1], ':')[1]
			$temp1[$i - 1][1] = StringSplit($temp[$i][1], ':')[2]
		Next
		$PassUser = $temp1
		$temp1 = 0
	EndIf
	;_ArrayDisplay($PassUser)
	$temp = 0
EndFunc   ;==>CheckINI
;test()

Opt("TrayOnEventMode", 1)
Opt("TrayAutoPause", 0)
start1(0)
Opt("TrayIconHide", 1)

Func SetServer($c, $mode = 0)
	While 1
		If IsArray($iniA) Then ExitLoop
		Global $iniA = @ScriptDir & '\temp\INFO.ini'
		$iniA = IniReadSection($iniA, 'Server')
		If @error Then
			__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Install INFO.ini.', 1)
			FileInstall('.\temp\INFO - Copy.ini', @ScriptDir & '\temp\INFO.ini', 1)
		EndIf
	WEnd
	Local $this, $a[1][1]
	$c = StringLower($c)
	$this = _ArraySearch($iniA, $c, 0, 0, 0, 1, 0)
	If $this <> -1 And $mode = 0 Then
		$this = $iniA[$this][1]
		If $this = '' Then
			$this = SetServer($c, 1)
		ElseIf $c = 'ServerIP' Then
			TCPStartup()
			(@error = 1) ? ($this = TCPNameToIP($this)) : ($this = TCPNameToIP('localhost'))
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
			Case 'copyto'
				$this = ''
			Case 'sleepperiod'
				$this = 5
		EndSwitch
	EndIf
	Return $this
EndFunc   ;==>SetServer

Func __LogWrite($file, $text, $mode = 1)
	Local $Ram = ProcessGetStats()
	If @Compiled = 1 Then
		TraySetToolTip(Int($Ram[0] / 1048576) & "MB " & $text)
	Else
		ConsoleWrite(@CRLF & Int($Ram[0] / 1048576) & "MB " & $text & @CRLF)
	EndIf
	_FileWriteLog($file, Int($Ram[0] / 1048576) & "MB | " & Int($Ram[1] / 1048576) & 'MB ' & TimerDiff($Start) & ' ' & $text, $mode)
	FileClose($file)
	Global $Start = TimerInit()
EndFunc   ;==>__LogWrite

Func ManualUpdate()
	If @TRAY_ID <> $ManualUpdate Then Return
	$SleepPeriod = 0
	Return
EndFunc   ;==>ManualUpdate

Func start1($arg = 1)
	Local $k = 1, $value = 0, $error = 0
	If @Compiled = 1 Then
		;Opt("TrayIconDebug", 1)
		If $arg = 0 Then
			;#RequireAdmin
			__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call Last24h.', 1)
			Local $begin = TimerInit()
			Run(@ScriptDir & '\Last24h.exe', "", @SW_HIDE, 0x8)
			__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call Main.', 1)
			main()
			If $CopyTo <> '' Then
				$temp = @ScriptDir & '\temp\INFO.ini'
				$temp = IniReadSection($temp, 'Server')
				$serverfile = SetServer('xlsxfile')
				FileCopy($serverfile, $CopyTo & '\TS3alpha.xlsx', 9)
			EndIf
			TraySetToolTip("Checked at: " & _StringFormatTime('%c', $StampTime) & ". Sleeping")
			$ManualUpdate = TrayCreateItem('Manual Update', -1, 1)
			$CheckINI = TrayCreateItem('Check INI file', -1, 1)
			TrayItemSetOnEvent($ManualUpdate, "ManualUpdate")
			TrayItemSetOnEvent($CheckINI, "CheckINI")
		EndIf
		Do
			Sleep(500)
			If TimerDiff($begin) >= $SleepPeriod * 60000 Then
				If $SleepPeriod = 0 Then $SleepPeriod = SetServer('sleepperiod')
				TrayItemSetState($ManualUpdate, 128 + 4)
				TrayItemSetState($CheckINI, 128 + 4)
				If @HOUR = 23 And $k = 0 And ProcessExists('Last24h.exe') = 0 Then
					$k = 1
					__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call Last24h.', 1)
					Run(@ScriptDir & '\Last24h.exe', "", @SW_HIDE, 0x8)
				ElseIf @HOUR <> 23 And $k = 1 Then
					$k = 0
				EndIf
				$StampTime = _TimeMakeStamp()
				$temp = @ScriptDir & '\temp\INFO.ini'
				$temp = IniReadSection($temp, 'Server')
				$serverfile = SetServer('xlsxfile')
				__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call Main.', 1)
				$begin = TimerInit()
				$value = main()
				$error = @error
				If $CopyTo <> '' Then
					$temp = @ScriptDir & '\temp\INFO.ini'
					$temp = IniReadSection($temp, 'Server')
					$serverfile = SetServer('xlsxfile')
					__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Copy to ' & $CopyTo & ' error(' & FileCopy($serverfile, $CopyTo & '\TS3alpha.xlsx', 9) & ').', 1)
				EndIf
				TraySetToolTip("Checked at: " & _StringFormatTime('%c', $StampTime) & ". Sleeping")
				If $error = 2 Then TraySetToolTip("TS3 Server was down at: " & _StringFormatTime('%c', $StampTime) & ". Sleeping")
				TrayItemSetState($ManualUpdate, 68)
				TrayItemSetState($CheckINI, 68)
			EndIf
		Until 0
	ElseIf @Compiled = 0 Then
		Opt("TrayIconDebug", 1)
		Local $begin = TimerInit(), $non = 0
		If $CopyTo <> '' Then
			FileCopy($serverfile, $CopyTo & '\TS3alpha.xlsx', 9)
		EndIf
		main()
		Do
			Sleep(500)
			If TimerDiff($begin) >= $SleepPeriod * 1000 * 60 Or ($CWtime <> 0) ? (_TimeGetStamp >= $CWtime) : (False) Then
				$StampTime = _TimeMakeStamp()
				$temp = @ScriptDir & '\temp\INFO.ini'
				$temp = IniReadSection($temp, 'Server')
				$serverfile = SetServer('xlsxfile')
				$temp = main()
				$begin = TimerInit()
				$non = $non + 1
			EndIf
			;Exit
		Until ($non > 100)
	EndIf
EndFunc   ;==>start1

Func main()
	#cs $file=FileOpen("d:/quang2.txt")
		
		for $i=1 to _FileCountLines("d:/quang2.txt")
		$temp=FileReadLine($file,$i)
		if StringLen($temp)>78 and StringLen($temp)<80 Then
		_FileWriteToLine("d:/quang2.txt",$i,$temp&" ",1)
		EndIf
		if @error<>0 Then
		MsgBox(0,"",@error)
		Exit
		EndIf
		Next
	#ce FileClose($file)
	Local $MainSocket, $ConnectedSocket, $temp1 = 0
	; Start The TCP Services
	;==============================================
	; Create a Listening "SOCKET".
	;   Using your IP Address and Port 33891.
	;==============================================
	Do
		TCPStartup()
		$MainSocket = TCPConnect($serverIP, $QueryPort)
		If @error Then
			@error = 0
			$array = 0
			$temp1 = $temp1 + 1
			If $temp1 >= 2 Then
				TrayTip('TS3 server is down!', 'TS3 server down!', 100)
				TraySetIcon('stop')
				Local $temp1 = [[0, 0, 0, 0, 0, 0]]
				LogUsers($temp1)
				IniWriteSection(@ScriptDir & '\temp\INFO.ini', 'PassUsers', '')
				__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] TS3 server is down at ' & _StringFormatTime('%c', $StampTime) & '.', 1)
				Return SetError(2, 0, 0)
			EndIf
			ContinueLoop
		EndIf
		TraySetIcon()
		; If the Socket creation fails, exit.
		TCPSend($MainSocket, "login client_login_name=" & $QueryUser & " client_login_password=" & $QueryPass & "" & @CRLF)
		Sleep(300)
		TCPSend($MainSocket, "use " & $QueryVituralServer & "" & @CRLF)
		;Sleep(300)
		TCPSend($MainSocket, "channellist" & @CRLF)
		;Sleep(300)
		TCPSend($MainSocket, "clientlist" & @CRLF)
		Sleep(3000)
		$string = TCPRecv($MainSocket, 1024000)
		;MsgBox(0,"",$string)
		TCPSend($MainSocket, 'quit' & @CRLF & @CRLF)
		TCPCloseSocket($MainSocket)
		TCPShutdown()
		;$file=FileOpen("d:/quang2.txt")
		;$string=FileRead($file)
		;FileClose($file)
		;MsgBox(0,"test",$string)
		$temp = StringSplit($string, 'error id=', 1)
		For $i = 1 To $temp[0]
			If StringIsDigit(StringLeft($temp[$i], 1)) Then
				If StringLeft($temp[$i], 1) <> 0 Then
					MsgBox(0, 'Error', $temp[$i])
					Exit
				EndIf
			EndIf
		Next

		$temp = 0
		$i = 0

		$string = StringReplace($string, @LF & @CR, "")
		$string = StringReplace($string, 'TS3Welcome to the TeamSpeak 3 ServerQuery interface, type "help" for a list of commands and "help <command>" ' & _
				'for information on a specific command.error id=0 msg=ok', "")
		$string = StringReplace($string, 'error id=0 msg=ok', "channel" & @CRLF & "::", 1)
		$string = StringReplace($string, 'error id=0 msg=ok', "::" & @CRLF & "client" & @CRLF & "::", 1)
		$string = StringReplace($string, 'error id=0 msg=ok', "")
		$string = StringReplace($string, '0quit', '0' & @CRLF)
;~ $file=FileOpen("D:/quang2.txt",2)
;~ FileWrite($file,$string)
;~ FileWrite($file,@cr&@lf)
;~ ConsoleWrite($string)
;~ FileFlush($file)
;~ FileClose($file)

		$array = StringSplit($string, '::', 1)
		$temp1 = $temp1 + 1
		If $temp1 >= 10 Then
			Local $temp1 = [[0, 0, 0, 0, 0, 0]]
			LogUsers($temp1)
			IniWriteSection(@ScriptDir & '\temp\INFO.ini', 'PassUsers', '')
			Return
		EndIf
	Until IsArray($array) = 1 And UBound($array) = 5
	$array[2] = StringReplace($array[2], 'channel_name=', "")
	$array[2] = StringReplace($array[2], 'cid=', '')
	$array[2] = StringReplace($array[2], 'pid=', '')
	$array[2] = StringReplace($array[2], 'total_clients=', '')
	$array[4] = StringReplace($array[4], @CRLF, '')
	$array[4] = StringReplace($array[4], 'clid=', '')
	$array[4] = StringReplace($array[4], 'cid=', '')
	$array[4] = StringReplace($array[4], 'client_database_id=', '')
	$array[4] = StringReplace($array[4], 'client_nickname=', '')
	$temp = StringSplit($array[2], '|')

	Local $array1[$temp[0] + 1][6]
	For $i = 0 To $temp[0]
		$array1[$i][0] = $temp[$i]
	Next
	For $i = 1 To $array1[0][0]
		$temp = 0
		$temp = StringSplit($array1[$i][0], ' ')
		For $j = 0 To ($temp[0] - 1)
			$array1[$i][$j] = $temp[$j + 1]
		Next
	Next

	For $i = 1 To $array1[0][0]
		_ArraySwap($array1[$i][0], $array1[$i][3])
		_ArraySwap($array1[$i][1], $array1[$i][3])
		_ArraySwap($array1[$i][2], $array1[$i][3])
		_ArraySwap($array1[$i][4], $array1[$i][3])
		$array1[$i][0] = StringReplace($array1[$i][0], '\s', ' ')
		$array1[$i][0] = StringReplace($array1[$i][0], '\/', '/')
		$array1[$i][0] = StringReplace($array1[$i][0], '\p', '|')
		$array1[$i][0] = StringReplace($array1[$i][0], '\\', '\')
	Next

	_ArrayDeleteCol($array1, 5)
	_ArrayDeleteCol($array1, 4)
	$array1[0][1] = "cid"
	$array1[0][2] = "pid"
	$array1[0][3] = 'on'
	$temp = StringSplit($array[4], '|')

	Local $array2[$temp[0] + 1][6]
	For $i = 0 To $temp[0]
		$array2[$i][0] = $temp[$i]
	Next
	For $i = 1 To $array2[0][0]
		$temp = 0
		$temp = StringSplit($array2[$i][0], ' ')
		For $j = 0 To ($temp[0] - 1)
			$array2[$i][$j] = $temp[$j + 1]
		Next
	Next
	For $i = 1 To $array2[0][0]
		_ArraySwap($array2[$i][0], $array2[$i][3])
		_ArraySwap($array2[$i][2], $array2[$i][3])
		$array2[$i][0] = StringReplace($array2[$i][0], '\s', ' ')
		$array2[$i][0] = StringReplace($array2[$i][0], '\/', '/')
		$array2[$i][0] = StringReplace($array2[$i][0], '\p', '|')
		$array2[$i][0] = StringReplace($array2[$i][0], '\\', '\')
	Next

	;_ArrayDeleteCol($array2,4)
	If _ArraySearch($array2, 'Unknow', 0, 0, 0, 1, 1, 0) <> -1 Then
		_ArrayDelete($array2, _ArraySearch($array2, 'Unknow', 0, 0, 0, 1, 1, 0))
		$array2[0][0] = $array2[0][0] - 1
	EndIf

	$array2[0][1] = 'cid'
	$array2[0][2] = 'clid'
	$array2[0][3] = 'cdid'
	$array2[0][4] = 'last'
	$array2[0][5] = 'clan'

;~ 	For $j = 9 To 12
;~ 		$temp1 = 0
;~ 		Do
;~ 			TCPStartup()
;~ 			$MainSocket = TCPConnect($serverIP, $QueryPort)
;~ 			TCPSend($MainSocket, "login client_login_name=" & $QueryUser & " client_login_password=" & $QueryPass & "" & @CRLF)
;~ 			Sleep(300)
;~ 			TCPSend($MainSocket, "use " & $QueryVituralServer & "" & @CRLF)
;~ 			TCPSend($MainSocket, 'servergroupclientlist sgid=' & $j & @CRLF)
;~ 			$temp = ''
;~ 			While StringRegExp($temp, 'error id=.*?' & @LF) = 0
;~ 				Sleep(500)
;~ 				$temp = $temp & TCPRecv($MainSocket, 1000000)
;~ 			WEnd
;~ 			TCPSend($MainSocket, 'quit' & @CRLF & @CRLF)
;~ 			TCPCloseSocket($MainSocket)
;~ 			$temp = StringRegExpReplace($temp, '(*ANYCRLF)', '')
;~ 			$temp = StringReplace($temp, 'error id=0 msg=ok', '')
;~ 			$temp = StringReplace($temp, 'cldbid=', '')
;~ 			$temp = StringSplit($temp, '|')
;~ 			$temp1 = $temp1 + 1
;~ 			If $temp1 >= 10 Then
;~ 				Local $temp=[0]
;~ 				ExitLoop 2
;~ 			EndIf
;~ 		Until IsArray($temp)
;~ 		For $i = 1 To $array2[0][0]
;~ 			;ConsoleWrite($i)
;~ 			If _ArraySearch($temp, $array2[$i][3], 1, 0, 0, 0, 1, 0) <> -1 Then
;~ 				$array2[$i][5] = $array2[$i][5] & '1'
;~ 			Else
;~ 				$array2[$i][5] = $array2[$i][5] & '0'
;~ 			EndIf
;~ 		Next
;~ 	Next
	#cs	If _ArraySearch($array2, 'quangkieu', 0, 0, 0, 1, 1, 0) <> -1 Then
		$temp = _ArraySearch($array2, 'quangkieu', 0, 0, 0, 1, 1, 0)
		$array2[$temp][0] = 'Quang Kieu KTAQ'
	#ce EndIf



	Do
		TCPStartup()
		$MainSocket = TCPConnect($serverIP, $QueryPort)
		Sleep(500)
		TCPSend($MainSocket, "login client_login_name=" & $QueryUser & " client_login_password=" & $QueryPass & "" & @CRLF)
		TCPSend($MainSocket, "use " & $QueryVituralServer & "" & @CRLF)

		For $i = 1 To $array2[0][0]
			TCPSend($MainSocket, "clientdbinfo cldbid=" & $array2[$i][3] & @CRLF)
		Next
		Sleep(3000)
		$string = TCPRecv($MainSocket, 1024000)
		TCPSend($MainSocket, 'quit' & @CRLF & @CRLF)
		TCPCloseSocket($MainSocket)
		TCPShutdown()
		$string = StringReplace($string, @LF & @CR, "")
		$string = StringReplace($string, 'TS3Welcome to the TeamSpeak 3 ServerQuery interface, type "help" for a list of commands and "help <command>" ' & _
				'for information on a specific command.error id=0 msg=ok', "")
		$string = StringReplace($string, 'error id=0 msg=ok', "", 1)
		$string = StringReplace($string, "error id=512 msg=invalid\sclientID", '' & _
				"client_unique_identifier=0= client_nickname=" & $QueryUser & " client_database_id=0 client_created=0 client_lastconnected=" & $StampTime & _
				" client_totalconnections=1 client_flag_avatar client_description client_month_bytes_uploaded=0 client_month_bytes_downloaded=0 client_total_bytes_uploaded=0" & _
				" client_total_bytes_downloaded=0 client_icon_id=0 client_base64HashClientUID=0 client_lastip=0error id=0 msg=ok")
		$string = StringReplace($string, 'error id=0 msg=ok', "::", $array2[0][0] - 1)
		$string = StringReplace($string, 'error id=0 msg=ok', '')
		$string = StringReplace($string, '0quit', '0' & @CRLF)
		$array = StringSplit($string, "::", 1)
		$temp1 = $temp1 + 1
		If $temp1 >= 4 Then
			Local $temp1 = [[0, 0, 0, 0, 0, 0]]
			LogUsers($temp1)
			Return
		EndIf
	Until IsArray($array) = 1 And UBound($array) = $array2[0][0] + 1 And (UBound($array) = 1) ? (True) : (StringSplit($array[1], ' ')[0] > 5)
	For $i = 1 To $array[0]
		$array2[$i][4] = StringReplace(StringSplit($array[$i], ' ')[5], "client_lastconnected=", '', 1)
	Next

	LogUsers($array2)
	__ImageServer($array1, $array2)




;~ _TS3Upload(@ScriptDir&'\SNS_A_'&@YEAR&'-'&@MON&'-'&@MDAY&'.txt','A')
;~ _TS3Upload(@ScriptDir&'\SNS_R_'&@YEAR&'-'&@MON&'-'&@MDAY&'.txt','R')
;~ _TS3Upload(@ScriptDir&'\SNS_C_'&@YEAR&'-'&@MON&'-'&@MDAY&'.txt','C')
;~ _TS3Upload('.txt','C')
;~ TCPStartup()
;~ $MainSocket = TCPConnect('185.4.75.32', '80')
;~ sleep(500)
;~ TCPSend($MainSocket,"GET /wot/clan/battles/?application_id=demo&clan_id=1000002504,1000006228,1000001787 HTTP/1.1"&@CRLF)
;~ TCPSend($MainSocket,"Host: api.worldoftanks.com"&@CRLF)
;~ TCPSend($MainSocket,"User-Agent: Mozilla/4.0"&@CRLF&@CRLF)
;~ Sleep(3000)
;~ $string=TCPRecv($MainSocket, 1024000)
;~ TCPCloseSocket($MainSocket)
;~ TCPShutDown()
;~ MsgBox(0,"",$string)


	;_ArrayDisplay($array)
	;_ArrayDisplay($array1)
	;_ArrayDisplay($array2)
	Dim $temp[$array2[0][0] + 1][2]
	For $i = 0 To $array2[0][0]
		_ArraySwap($array2[$i][2], $array2[$i][3])
		$temp[$i][0] = $array2[$i][2]
		$temp[$i][1] = $array2[$i][0] & ':' & $array2[$i][1]
	Next

	ReDim $array2[$array2[0][0] + 1][3]
	$PassUser = $array2
	$PassUser[0][1] = $StampTime
	$temp[0][0] = 'Stamp'
	$temp[0][1] = $StampTime
	;_ArrayDisplay($PassUser)
	;_ArrayDisplay($temp)
	IniWriteSection(@ScriptDir & '\temp\INFO.ini', 'PassUsers', '')
	IniWriteSection(@ScriptDir & '\temp\INFO.ini', 'PassUsers', $temp, 0)
	$temp = 0
	Return SetError(0, UBound($PassUser), 0)
EndFunc   ;==>main

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

Func InsertNewLine(ByRef $a, ByRef $b, ByRef $c)
	_Excel_RangeWrite($a, 3, _Excel_RangeRead($a, 3, _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0]) + 2) & _Excel_RangeRead($a, 3, _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0]) + 3) & '1')) + 1, _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0]) + 3) & '1')
	_Excel_RangeInsert(_Excel_SheetList($a)[2][1], _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0])) & '3:' & _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0]) + 3) & '3', $xlShiftDown)
	;If @error Then MsgBox(0, '', @error & ' ' & _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0])) & '3:' & _Excel_ColumnToLetter(_Excel_ColumnToNumber($b[$c][0]) + 3) & '3')
EndFunc   ;==>InsertNewLine

Func WriteNewLine(ByRef $Book, ByRef $b, ByRef $c, ByRef $d)
	_Excel_RangeWrite($Book, 3, $d, $b[$c][0] & '3')
EndFunc   ;==>WriteNewLine

Func _CloseExcel()
	If @exitCode < 3 Then
		If Not ProcessExists('EXCEL.EXE') Or Not ProcessExists('excel.exe') Then Return
		Local $oBook = _Excel_BookAttach($serverfile)
		If Not @error Then
			_Excel_BookClose($oBook)
		EndIf
	EndIf
EndFunc   ;==>_CloseExcel


Func LogUsers(ByRef $c)
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call LogUsers.', 1)
	Local $array2 = $c, $temp, $temp1, $starttime = _TimeMakeStamp(0, 0, 0)
;~ 	ReDim $array2[$array2[0][0]][5]
;~ 	$array2[0][0] = $array2[0][0] - 1
;~ 	;_ArrayDisplay($PassUser)
;~ 	If $PassUser[0][0] >= 2 Then
;~ 		_ArrayDelete($PassUser, 3)
;~ 		$PassUser[0][0] = $PassUser[0][0] - 1
;~ 	EndIf
	$oEx = 0
	OnAutoItExitRegister('_CloseExcel')
	$oBook = _Excel_BookAttach($serverfile)
	If @error Then
		$oEx = _Excel_Open(False, Default, Default, Default, True)
		Sleep(500)
		$oBook = _Excel_BookOpen($oEx, $serverfile, False, True)
	ElseIf $oBook.Application.Visible = False Then
		If @Compiled = 0 Then $oBook.Application.Visible = True
		$oBook = 0
		TraySetToolTip('[TS3Main] Wait for other process to finish!')
		TraySetIcon('warning')
		ProcessWaitClose(ProcessExists('EXCEL.EXE'), 60 * 30)
		$temp = ObjGet("", "Excel.Application")
		If IsObj($temp) And ObjName($temp, 1) <> "_Application" Then
			For $oWorkbook In $temp.Workbooks
				If Not $temp.Saved Then
					$temp.Save()
				EndIf
			Next
			$temp.Quit()
			$temp = 0
		EndIf
		TraySetIcon()
		TraySetToolTip()
		$oEx = _Excel_Open(False)
		Sleep(500)
		$oBook = _Excel_BookOpen($oEx, $serverfile, False, True)
	EndIf
;~ 	$temp = ObjGet("", "Excel.Application")
;~ 	If Not @error Then
;~ 		If $temp.Visible Then
;~ 			$oEx=
;~ 			If @error Then
;~ 				$temp = 0
;~
;~ 				TrayTip("Please close Excel!", "[ExcelIndex] Please close Excel!", 5)
;~ 				TraySetToolTip("Please close Excel!")
;~ 				TraySetIcon("warning")
;~ 				ProcessWaitClose(ProcessExists('EXCEL.EXE'))
;~ 			Else
;~ 			EndIf
;~ 		Else
;~ 			$temp=0
;~ 			TraySetToolTip("Wait for other process to finish!")
;~ 			TraySetIcon("warning")
;~ 			ProcessWaitClose(ProcessExists('EXCEL.EXE'),60*15)
;~ 			$temp=ObjGet("", "Excel.Application")
;~ 			If IsObj($temp) and ObjName($temp, 1) <> "_Application" Then
;~ 				For $oWorkbook In $temp.Workbooks
;~ 					If Not $temp.Saved Then
;~ 						$temp.Save()
;~ 					EndIf
;~ 				Next
;~ 				$temp.Quit()
;~ 				$temp=0
;~ 			EndIf
;~ 		EndIf
;~ 		TraySetToolTip()
;~ 		TraySetIcon()
;~ 	EndIf
	ConsoleWrite(@error & ' ' & @ScriptLineNumber)
	$table1 = _Excel_RangeRead($oBook, 3, 'A3:D' & _Excel_RangeRead($oBook, 3, 'A1') + 3)
	;_ArrayDisplay($PassUser)
	;_ArrayDisplay($table1)
	;ConsoleWrite(@error & ' ' & $PassUser[0][0] & @ScriptLineNumber )
	If $array2[0][0] >= 2 Then
		For $i = 1 To $array2[0][0]
			$temp = _ArraySearch($table1, $array2[$i][3], 0, 0, 0, 0, 1, 1)
			$temp1 = _ArraySearch($PassUser, $array2[$i][3], 1, 0, 0, 0, 1, 2)
			If $temp <> -1 Then
				ConsoleWrite('this')
				$table2 = $table1[$temp][0] & '3:' & _Excel_ColumnToLetter(_Excel_ColumnToNumber($table1[$temp][0]) + 3) & '3'
				$table2 = _Excel_RangeRead($oBook, 3, $table2)
				If $table2[0][1] = '' And $table2[0][2] <> '' Then
					ConsoleWrite('1st' & $temp)
					Dim $table2 = [['', '', '', '']]
					WriteNewLine($oBook, $table1, $temp, $table2)
				ElseIf $table2[0][0] = '' And $table2[0][1] = '' And $table2[0][2] = '' Then
					ConsoleWrite('2nd' & $temp)
					Dim $table2 = [['', $StampTime, '', $array2[$i][1]]]
					If $temp1 = -1 Then Dim $table2 = [['', $StampTime, '', $array2[$i][1]]]
					WriteNewLine($oBook, $table1, $temp, $table2)
				ElseIf StringInStr($table2[0][0], @MON & '/' & @MDAY & '/' & @YEAR) <> 0 Or $table2[0][1] >= $starttime Then
					If $table2[0][1] = '' Then
						ConsoleWrite('3r1d')
						$table2[0][1] = $array2[$i][4]
						$table2[0][3] = $array2[$i][1]
						WriteNewLine($oBook, $table1, $temp, $table2)
					ElseIf $table2[0][1] <> '' And $table2[0][3] <> '' And $table2[0][3] <> $array2[$i][1] Then
						ConsoleWrite('3r2d' & $temp)
						$table2[0][2] = $StampTime - 1
						WriteNewLine($oBook, $table1, $temp, $table2)
						InsertNewLine($oBook, $table1, $temp)
						Dim $table2 = [['', $StampTime, '', $array2[$i][1]]]
						WriteNewLine($oBook, $table1, $temp, $table2)
					ElseIf $table2[0][1] <> '' And $table2[0][3] = $array2[$i][4] And $temp1 = -1 Then
						ConsoleWrite('3r3d' & $temp)
						$table2[0][2] = $StampTime - 1
						WriteNewLine($oBook, $table1, $temp, $table2)
						InsertNewLine($oBook, $table1, $temp)
						Dim $table2 = [['', $StampTime, '', $array2[$i][1]]]
						WriteNewLine($oBook, $table1, $temp, $table2)
					EndIf
				ElseIf StringInStr($table2[0][0], @MON & '/' & @MDAY & '/' & @YEAR) = 0 Or $table2[0][1] < $starttime Then
					If $table2[0][1] = '' Then
						ConsoleWrite('4t1h' & $temp)
						InsertNewLine($oBook, $table1, $temp)
						Dim $table2 = [['', $StampTime, '', $array2[$i][1]]]
						WriteNewLine($oBook, $table1, $temp, $table2)
					ElseIf $table2[0][1] <> '' And $table2[0][2] = '' Then
						ConsoleWrite('4th' & $temp)
						If $temp1 <> -1 Then
							ConsoleWrite('4t2h' & $temp)
							$table2[0][2] = ($starttime - 1)
							WriteNewLine($oBook, $table1, $temp, $table2)
						EndIf
						InsertNewLine($oBook, $table1, $temp)
						Dim $table2 = [['', $StampTime, '', $array2[$i][1]]]
						WriteNewLine($oBook, $table1, $temp, $table2)

					ElseIf $table2[0][1] <> '' And $table2[0][2] <> '' Then
						ConsoleWrite('5th' & $temp)
						InsertNewLine($oBook, $table1, $temp)
						Dim $table2 = [['', $StampTime, '', $array2[$i][1]]]
						WriteNewLine($oBook, $table1, $temp, $table2)
					EndIf
				EndIf
			EndIf
		Next
	EndIf

	If $PassUser[0][0] >= 1 Then
		For $i = 1 To $PassUser[0][0]
			$temp = _ArraySearch($table1, $PassUser[$i][2], 0, 0, 0, 0, 1, 1)
			ConsoleWrite(@error & ' ' & @ScriptLineNumber)
			$temp1 = _ArraySearch($array2, $PassUser[$i][2], 1, 0, 0, 0, 1, 3)
			ConsoleWrite(@error & ' ' & @ScriptLineNumber)
			If $temp1 = -1 And $temp <> -1 Then
				$table2 = $table1[$temp][0] & '3:' & _Excel_ColumnToLetter(_Excel_ColumnToNumber($table1[$temp][0]) + 3) & '3'
				;MsgBox(0,'',$table2)
				$table2 = _Excel_RangeRead($oBook, 3, $table2)
				;_ArrayDisplay($table2)
				;MsgBox(0,'',StringInStr($table2[0][0],@MON & '/' & @MDAY & '/' & @YEAR)&' '&$table2[0][1]&' '&$starttime)
				If $table2[0][1] = '' And $table2[0][2] <> '' Then
					ConsoleWrite('pass1')
					Dim $table2 = [['', '', '', '']]
					WriteNewLine($oBook, $table1, $temp, $table2)
				ElseIf StringInStr($table2[0][0], @MON & '/' & @MDAY & '/' & @YEAR) <> 0 Or $table2[0][1] >= $starttime Then
					If $table2[0][1] <> '' Then
						ConsoleWrite('off' & $temp)
						If $table2[0][3] <> '' And $table2[0][3] <> $PassUser[$i][1] Then

							$table2[0][2] = $PassUser[0][1] + 1
							WriteNewLine($oBook, $table1, $temp, $table2)
							InsertNewLine($oBook, $table1, $temp)
							Dim $table2 = [['', $PassUser[0][1] + 2, $StampTime, $PassUser[$i][1]]]
							WriteNewLine($oBook, $table1, $temp, $table2)
						ElseIf $table2[0][3] = $PassUser[$i][1] Then
							$table2[0][2] = $StampTime
							WriteNewLine($oBook, $table1, $temp, $table2)
							InsertNewLine($oBook, $table1, $temp)
						EndIf
					EndIf
				ElseIf StringInStr($table2[0][0], @MON & '/' & @MDAY & '/' & @YEAR) = 0 Or $table2[0][1] < $starttime Then
					If $table2[0][1] <> '' Then
						ConsoleWrite('that' & $temp)
						$table2[0][2] = $starttime - 1
						WriteNewLine($oBook, $table1, $temp, $table2)
						InsertNewLine($oBook, $table1, $temp)
						Dim $table2 = [['', $starttime, $StampTime, $PassUser[$i][1]]]
						WriteNewLine($oBook, $table1, $temp, $table2)
					EndIf
				EndIf
			EndIf
		Next
	EndIf
	_Excel_BookSave($oBook)
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call CW log.', 1)
	$CWtime = _ClanWar($c, $oBook)
	Local $error = @error, $ext = @extended
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Return from CW log (' & $error & ',' & $ext & ',' & $CWtime & ').', 1)
	_Excel_BookSave($oBook)
	If Not ($oBook.Application.Visible) Then
		_Excel_BookClose($oBook)
		_Excel_Close($oEx)
	EndIf
	OnAutoItExitUnRegister('_CloseExcel')
	$oEx = 0
EndFunc   ;==>LogUsers

Func __ImageServer(ByRef $array1, ByRef $array2)
	Local $temp, $temp1, $imagefile, $temp2
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call ServerImage.', 1)
	Local $imagefile = FileOpen(@ScriptDir & '\temp\TS3_Image\' & @YEAR & '-' & @MON & '-' & @MDAY & '_' & @HOUR & @MIN & @SEC & @MSEC & '.txt', 10)
	;$imagefile = FileOpen('D:/quang4.txt', 10)
	FileWriteLine($imagefile, 'Server image at ' & @MON & '/' & @MDAY & '/' & @YEAR & ' ' & @HOUR & ':' & @MIN & ':' & @SEC & ':' & @CRLF & @CRLF & '(...):channel' & @CRLF & _
			StringFormat('[SNS-A] %5d\n[SNS-R] %5d\n[SNS-C] %5d', UBound(_ArrayFindAll($array2, '.1..', 1, 0, 0, 3, 5)), UBound(_ArrayFindAll($array2, '..1.', 1, 0, 0, 3, 5)), UBound(_ArrayFindAll($array2, '...1', 1, 0, 0, 3, 5))) & @CRLF)
	$temp = 0
	For $i = 1 To $array1[0][0]
		If $array1[$i][2] = 0 Then
			$temp = 0
		ElseIf $array1[$i][2] = $array1[$i - 1][1] Then
			$temp += 1
		ElseIf $array1[$i][2] = $array1[$i - 1][2] Then
			$temp = $temp
		ElseIf $array1[$i][2] <> $array1[$i - 1][1] Then
			$temp1 = 0
			$temp2 = $i
			ConsoleWrite($temp2 & ' ' & $array1[$temp2][2] & @CRLF)
			Do
				$temp2 = _ArraySearch($array1, $array1[$temp2][2], 0, 0, 0, 0, 1, 1)
				If $temp2 <> -1 Then
					If $array1[$temp2][2] = 0 Then
						$temp = $temp1 + 1
					Else
						$temp1 = 1 + $temp1
					EndIf
				EndIf
			Until $array1[$temp2][2] = 0
		EndIf
		$temp1 = 0
		$temp2 = 0
		If $temp <> 0 Then
			For $j = 1 To $temp
				FileWrite($imagefile, @TAB)
			Next
		EndIf
		FileWrite($imagefile, '(' & $array1[$i][0] & ')  ' & $array1[$i][3] & @CRLF)
		$temp1 = 1
		Do
			$temp2 = _ArraySearch($array2, $array1[$i][1], $temp1, 0, 0, 0, 1, 1)
			If $temp2 <> -1 Then
				For $j = 1 To $temp + 2
					FileWrite($imagefile, @TAB)
				Next
				FileWrite($imagefile, '(' & _StringFormatTime("%c", $array2[$temp2][4]) & ')' & @TAB & StringFormat('[%5d]    ', $array2[$temp2][3]) & $array2[$temp2][0] & @CRLF)
				$temp1 = $temp2 + 1
			EndIf
		Until $temp2 = -1
		$temp1 = 0
		$temp2 = 0
	Next
	FileFlush($imagefile)
	FileClose($imagefile)
EndFunc   ;==>__ImageServer

Func _ClanWar(ByRef $array2, ByRef $oBook)
	Local $temp, $temp1, $temp2, $table, $table1
	Do
		Local $clan = IniReadSection(@ScriptDir & '\temp\INFO.ini', 'Clans') ;[[4, 0],['Region', 'na'],['SNS-A', '1000002504:10:14,15,16'],['SNS-R', '1000001787:11:17,18,19'],['SNS-C', '1000006228:12:26,27,28']]
		If @error Or Not (IsArray($clan)) Then
			Local $temp = [['Region', 'na'],['SNS-A', '1000002504:10:14,15,16'],['SNS-R', '1000001787:11:17,18,19'],['SNS-C', '1000006228:12:26,27,28']]
			IniWriteSection(@ScriptDir & '\temp\INFO.ini', 'Clans', $temp, 0)
		EndIf
	Until IsArray($clan)
	Local $temp = [['na', '.com'],['eu', '.eu'],['ru', '.ru'],['sea', '-sea.com'],['kr', '.kr']]
	Local $region = $temp[_ArraySearch($temp, $clan[1][1])][1]
	;$temp=_Date_Time_GetTimeZoneInformation()
	If UBound($clan, 0) <> 2 Or $clan[0][0] < 1 Then Return SetError(1, 0, 0)
	If $clan[1][0] <> 'Region' Then Return SetError(1, 1, 0)
	For $i = 2 To $clan[0][0]
		$clan[$i][1] &= ':'
	Next
	If UBound($array2) > 1 And $array2[0][5] <> 1 Then $array2[0][5] = 1
	For $c = 2 To $clan[0][0]
		If UBound($array2) > 1 And $array2[0][5] <> 1 Then
			If StringSplit($clan[$c][1], ':')[0] <> 4 Then Return SetError(2, $c - 1, 0)
			$temp1 = 0
			Do
				TCPStartup()
				$MainSocket = TCPConnect($serverIP, $QueryPort)
				TCPSend($MainSocket, "login client_login_name=" & $QueryUser & " client_login_password=" & $QueryPass & "" & @CRLF)
				Sleep(300)
				TCPSend($MainSocket, "use " & $QueryVituralServer & "" & @CRLF)
				TCPSend($MainSocket, 'servergroupclientlist sgid=' & StringSplit($clan[$c][1], ':')[2] & @CRLF)
				$temp = ''
				While StringRegExp($temp, 'error id=.*?' & @LF) = 0
					Sleep(500)
					$temp = $temp & TCPRecv($MainSocket, 1000000)
				WEnd
				TCPSend($MainSocket, 'quit' & @CRLF & @CRLF)
				TCPCloseSocket($MainSocket)
				$temp = StringRegExpReplace($temp, '(*ANYCRLF)', '')
				$temp = StringReplace($temp, 'error id=0 msg=ok', '')
				$temp = StringReplace($temp, 'cldbid=', '')
				$temp = StringSplit($temp, '|')
				$temp1 = $temp1 + 1
				If $temp1 >= 10 Then
					Local $temp = [0]
					ExitLoop 2
				EndIf
			Until IsArray($temp)
			For $i = 1 To $array2[0][0]
				;ConsoleWrite($i)
				If _ArraySearch($temp, $array2[$i][3], 1, 0, 0, 0, 1, 0) <> -1 Then
					$array2[$i][5] = $array2[$i][5] & '1'
				Else
					$array2[$i][5] = $array2[$i][5] & '0'
				EndIf
			Next
		EndIf
		For $i = 2 To $clan[0][0]
			If $i = $c Then $clan[$c][1] &= '1'
			If $i <> $c Then $clan[$c][1] &= '.'
		Next
	Next
	If @Compiled = 0 And _ArraySearch($array2, 'quangkieu', 1, 0, 0, 1, 0, 0) > -1 Then $array2[_ArraySearch($array2, 'quangkieu', 1, 0, 0, 1, 0, 0)][5] = 111
	For $c = 2 To $clan[0][0]
		__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Get CW battles from WG for ' & $clan[$c][0] & '.', 1)
		__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '===========================================' & @CRLF & '                   [TS3Main] Get CW battles from WG for ' & $clan[$c][0] & '.', 1)
		If StringSplit($clan[$c][1], ':')[0] <> 4 Then Return SetError(2, $c - 1, 0)
		For $i = 0 To 2
			;__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] WG:' & $temp2, 1)
			$temp1 = inetread_timeout('http://api.worldoftanks' & $region & '/wot/clan/battles/?application_id=8b379c513ce91c4e9ddfe538a12b0e67&clan_id=' & StringSplit($clan[$c][1], ':')[1], 9)
			$temp1 = BinaryToString($temp1)
			;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
			Sleep(500)
			$temp2 = inetread_timeout('http://api.worldoftanks' & $region & '/wot/clan/battles/?application_id=8b379c513ce91c4e9ddfe538a12b0e67&clan_id=' & StringSplit($clan[$c][1], ':')[1], 9)
			$temp2 = BinaryToString($temp2)
			;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
			If $temp1 <> $temp2 Then ExitLoop
		Next
;~ 		$temp2 = '{"status":"ok","count":1,"data":{"1000002504":[{"provinces":["CA_86"],"started":true,"private":null,"time":' & $StampTime + 650 & ',"arenas":[{"name_i18n":"Severogorsk","name":"85_winter"}],"type":"landing"},{"provinces":["CA_84"],"started":true,"private":null,"time":' & $StampTime + 600 & ',"arenas":[{"name_i18n":"Swamp","name":"22_slough"}],"type":"for_province"},{"provinces":["CA_87"],"started":true,"private":null,"time":' & $StampTime & ',"arenas":[{"name_i18n":"Tundra","name":"63_tundra"}],"type":"for_province"},{"provinces":["CA_87"],"started":true,"private":null,"time":' & $StampTime - 600 & ',"arenas":[{"name_i18n":"Tundra","name":"63_tundra"}],"type":"for_province"}]}'
		If @error Or StringInStr($temp2, '"error"') <> 0 Then Return SetError(3, $c - 1, 0);no connection or wrong ClanID,1396577583,
		__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] WG:' & $temp2, 1)
		;MsgBox(0, '', $temp2)
		$temp1 = StringRegExp($temp2, '(?!"provinces":\[)(,*"\w+_\d+")+(?=],"started")', 3)
		$temp = UBound($temp1)
		__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] ' & $clan[$c][0] & ': ' & $temp & ' battle(s)', 1)
		If $temp = 0 Then ContinueLoop
		$table = $temp2
		__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] Get CW battles from wotapi.ru for ' & $clan[$c][0] & '.', 1)
;~ 	For $i = 0 To 2
;~ 		$temp1 = inetread_timeout('http://cw.wotapi.ru/na/1000002504?type=json', 9)
;~ 		$temp1 = BinaryToString($temp1)
;~ 		;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
;~ 		Sleep(500)
;~ 		$temp2 = inetread_timeout('http://cw.wotapi.ru/na/1000002504?type=json', 9)
;~ 		$temp2 = BinaryToString($temp2)
;~ 		;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
;~ 		If $temp1 <> $temp2 Then ExitLoop
;~ 	Next
;~ 	If @error Or StringInStr($temp2, '"error"') <> 0 Then Return SetError(2, 0, 0); no connection or wrong ClanID
		$temp2 = '{"data":{"battles_count":0,"last_update":';"1396523110","status":"ok","battles":[{"type":"for_province","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: -32px;","time":"01:30","primetime":"00:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_84"},"province":{"name":"Eastern Kativik","id":"CA_84","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_84"},"maps":"Swamp","server":"US301"},{"type":"for_province","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: -32px;","time":"01:30","primetime":"00:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_87"},"province":{"name":"Southern Labrador","id":"CA_87","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_87"},"maps":"Sacred Valley","server":"US301"}]}}';problem in sync between 2 servers
		__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] wotapi:' & $temp2, 1)
		;MsgBox(0, '', $temp2)
		If StringRegExp($temp2, '"battles_count":\K\d+(?=,"last_update":)', 3)[0] = 0 Then $temp2 = '{"data":{"battles_count":0,"last_update":"1396523110","status":"ok","battles":[{"type":"for_province","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: -32px;","time":"01:30","primetime":"00:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_84"},"province":{"name":"Eastern Kativik","id":"CA_84","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_84"},"maps":"Swamp","server":"US301"},{"type":"for_province","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: -32px;","time":"01:30","primetime":"00:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_87"},"province":{"name":"Southern Labrador","id":"CA_87","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_87"},"maps":"Sacred Valley","server":"US301"}]}}';problem in sync between 2 servers
		$temp1 = $temp2
		$temp2 = $table
		;started | time | region | province | map | code | type | server |
		;	/		/							/	/		/
		#Region
		If $temp <> 0 Then
			Local $table[$temp + 1][8]
			$table[0][0] = $temp
			$temp = $temp1
			$temp1 = StringRegExp($temp2, '"provinces":\[\K.*?(?=],"started")', 3)
			For $i = 1 To $table[0][0]
				$table[$i][5] = StringReplace($temp1[$i - 1], '"', '')
			Next
			$temp1 = StringRegExp($temp2, '"started":\K.*?(?=,"private")', 3)
			For $i = 1 To $table[0][0]
				$table[$i][0] = $temp1[$i - 1]
			Next
			$temp1 = StringRegExp($temp2, '"time":\K.*?(?=,"arenas")', 3)
			For $i = 1 To $table[0][0]
				$table[$i][1] = $temp1[$i - 1]
			Next
			$temp1 = StringRegExp($temp2, '"name_i18n":"\K.*?(?=","name")', 3)
			For $i = 1 To $table[0][0]
				$table[$i][4] = $temp1[$i - 1]
			Next
			$temp1 = StringRegExp($temp2, '"type":"\K.*?(?="})', 3)
			For $i = 1 To $table[0][0]
				$table[$i][6] = $temp1[$i - 1]
			Next
		EndIf
		;_ArrayDisplay($table)
		$temp2 = $temp
		;code | type | region | province | server |
		Local $temp[StringRegExp($temp2, '_count":\K.*?(?=,")', 3)[0] + 1][5]
		$temp[0][0] = UBound($temp) - 1
		If $table[0][0] <> 0 And $temp[0][0] <> 0 Then
			$temp1 = StringRegExp($temp2, 'e":{"name".*?"id":"\K.*?(?=",)', 3)
			For $i = 1 To $temp[0][0]
				$temp[$i][0] = $temp1[$i - 1]
			Next
			$temp1 = StringRegExp($temp2, '"type":"\K.*?(?=",)', 3)
			For $i = 1 To $temp[0][0]
				$temp[$i][1] = $temp1[$i - 1]
			Next
			$temp1 = StringRegExp($temp2, 'n":{"name":"\K.*?(?=","id")', 3)
			For $i = 1 To $temp[0][0]
				$temp[$i][2] = $temp1[$i - 1]
			Next
			$temp1 = StringRegExp($temp2, 'e":{"name":"\K.*?(?=",)', 3)
			For $i = 1 To $temp[0][0]
				$temp[$i][3] = $temp1[$i - 1]
			Next
			$temp1 = StringRegExp($temp2, 'r":"\K.*?(?="})', 3)
			For $i = 1 To $temp[0][0]
				$temp[$i][4] = $temp1[$i - 1]
			Next
		EndIf
		;started | time | region | province | map | code | type | server |
		;	0		1		2			3		4	5		6		7
		;code | type | region | province | server |
		;_ArrayDisplay($temp)
		For $i = $table[0][0] To 1 Step -1
			For $j = 1 To $temp[0][0]
				If $table[$i][6] <> 'meeting_engagement' And $temp[$j][1] <> 'meeting_engagement' And $table[$i][5] = $temp[$j][0] Then
					$table[$i][2] = $temp[$j][2]
					$table[$i][3] = $temp[$j][3]
					$table[$i][7] = $temp[$j][4]
					ExitLoop
				ElseIf $table[$i][6] = 'meeting_engagement' Then
					_ArrayDelete($table, $i)
					$table[0][0] = UBound($table) - 1
				EndIf
			Next
			$table[$i][0] = $table[$i][5]
			$table[$i][2] = $table[$i][4] & ',' & $table[$i][2] & ',' & $table[$i][3]
			$table[$i][3] = $table[$i][7]
		Next
		;code | time | map,region,province | server |
		;_ArrayDisplay($table)
		ReDim $table[$table[0][0] + 1][4]
		;Local $temp2=[14,15,16]
		Local $tt = 0
		Do
			TCPStartup()
			$MainSocket = TCPConnect($serverIP, $QueryPort)
			$tt = $tt + 1
			If $tt >= 5 Then Return SetError(5, $c - 1, 0)
		Until @error = 0
		Sleep(500)
		TCPSend($MainSocket, "login client_login_name=" & $QueryUser & " client_login_password=" & $QueryPass & "" & @CRLF)
		TCPSend($MainSocket, "use " & $QueryVituralServer & "" & @CRLF)
		Sleep(500)
		TCPRecv($MainSocket, 10000000)
		Sleep(300)
		;MsgBox(0, '', TCPRecv($MainSocket, 100000))
		If _Excel_RangeRead($oBook, 2, '' & _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '1') = 0 Or _Excel_RangeRead($oBook, 2, '' & _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '1') = '' Then
			__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] Write new CW report.', 1)
			$table[0][1] = StringSplit($clan[$c][1], ':')[4] & $clan[$c][0]
			$table[0][2] = @WDAY & '|' & @MON & '\' & @MDAY & '\' & @YEAR
			_Excel_RangeWrite($oBook, 2, $table, '' & _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '1')
			For $i = $table[0][0] To 1 Step -1
				If ($StampTime < $table[$i][1] - (10 * 60)) And $StampTime >= $table[$i][1] - 15 * 60 Then
					__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] CW code 1:' & _StringFormatTime('%c', $table[$i][1]) & ' ' & $table[$i][2])
					For $temp1 = 1 To UBound($array2) - 1
						If StringRegExp($array2[$temp1][5], StringSplit($clan[$c][1], ':')[4]) = 1 And StringRegExp($array2[$temp1][1], '' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & '') = 0 Then
							TCPSend($MainSocket, 'clientpoke clid=' & $array2[$temp1][2] & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
									']\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] - $StampTime) / 60) & '\sminutes.\sBe\sready!' & @CRLF)
						EndIf
					Next
					Do
						TCPSend($MainSocket, 'whoami' & @CRLF)
						Sleep(500)
						$temp2 = TCPRecv($MainSocket, 1000000)
						$temp2 = StringRegExp($temp2, 'client_id=\K\d+(?= )', 3)
					Until IsArray($temp2) And UBound($temp2, 0) = 1
					$temp2 = $temp2[0]
					For $cwc = 1 To StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[0]
						$temp1 = StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[$cwc]
						TCPSend($MainSocket, 'clientmove clid=' & $temp2 & ' cid=' & $temp1 & @CRLF)
						TCPSend($MainSocket, 'sendtextmessage targetmode=2 target=' & $temp1 & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
								']\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] - $StampTime) / 60) & '\sminutes.\sThere\sare\s' & UBound(_ArrayFindAll($array2, StringSplit($clan[$c][1], ':')[4], 1, 0, 0, 3, 5)) & '\s' & $clan[$c][0] & '\splayers\son\sTS3' & @CRLF)
					Next
				EndIf
				If ($StampTime < $table[$i][1]) And $StampTime >= $table[$i][1] - 10 * 60 Then
					__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] CW code 2:' & _StringFormatTime('%c', $table[$i][1]) & ' ' & $table[$i][2])
					$temp = ''
					For $temp1 = 1 To UBound($array2) - 1
						If StringRegExp($array2[$temp1][1], '' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & '') = 1 And StringRegExp($array2[$temp1][5], StringSplit($clan[$c][1], ':')[4]) = 1 Then
							$temp &= $array2[$temp1][3] & ','
						EndIf
						If StringRegExp($array2[$temp1][5], StringSplit($clan[$c][1], ':')[4]) = 1 And StringRegExp($array2[$temp1][1], '' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & '') = 0 Then
							TCPSend($MainSocket, 'clientpoke clid=' & $array2[$temp1][2] & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
									']\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] - $StampTime) / 60) & '\sminutes.\sPlease\sjoin\sCW\sclannels!' & @CRLF)
						EndIf
					Next
					Do
						TCPSend($MainSocket, 'whoami' & @CRLF)
						Sleep(500)
						$temp2 = TCPRecv($MainSocket, 1000000)
						$temp2 = StringRegExp($temp2, 'client_id=\K\d+(?= )', 3)
					Until IsArray($temp2) And UBound($temp2, 0) = 1
					$temp2 = $temp2[0]
					For $cwc = 1 To StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[0]
						$temp1 = StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[$cwc]
						TCPSend($MainSocket, 'clientmove clid=' & $temp2 & ' cid=' & $temp1 & @CRLF)
						TCPSend($MainSocket, 'sendtextmessage targetmode=2 target=' & $temp1 & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
								']\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] - $StampTime) / 60) & '\sminutes.\sThere\sare\s' & UBound(_ArrayFindAll($array2, StringSplit($clan[$c][1], ':')[4], 1, 0, 0, 3, 5)) & '\s' & $clan[$c][0] & '\splayers\son\sTS3' & @CRLF)
					Next
					Local $temp1 = [[$temp, '', '|', '']]
					_Excel_RangeInsert(_Excel_SheetList($oBook)[1][1], _Excel_ColumnToLetter(($c - 2) * 4 + 1) & $i + 2 & ':' & _Excel_ColumnToLetter(($c - 2) * 4 + 4) & $i + 2, $xlShiftDown)
					_Excel_RangeWrite($oBook, 2, $temp1, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & $i + 2)
					Local $temp = 0, $temp1 = 0, $temp2 = 0
				ElseIf $StampTime >= $table[$i][1] And $StampTime <= $table[$i][1] + 10 * 60 Then
					__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] CW code 3:' & _StringFormatTime('%c', $table[$i][1]) & ' ' & $table[$i][2])
					Do
						TCPSend($MainSocket, 'whoami' & @CRLF)
						Sleep(500)
						$temp2 = TCPRecv($MainSocket, 1000000)
						$temp2 = StringRegExp($temp2, 'client_id=\K\d+(?= )', 3)
					Until IsArray($temp2) And UBound($temp2, 0) = 1
					$temp2 = $temp2[0]
					For $cwc = 1 To StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[0]
						$temp1 = StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[$cwc]
						TCPSend($MainSocket, 'clientmove clid=' & $temp2 & ' cid=' & $temp1 & @CRLF)
						TCPSend($MainSocket, 'sendtextmessage targetmode=2 target=' & $temp1 & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
								']\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] + 60 * 15 - $StampTime) / 60) & '\sminutes.\sThere\sare\s' & UBound(_ArrayFindAll($array2, StringSplit($clan[$c][1], ':')[4], 1, 0, 0, 3, 5)) & '\s' & $clan[$c][0] & '\splayers\son\sTS3' & @CRLF)
					Next
					$temp = ''
					Local $temp1 = [[',', ',', '|', ',']]
					For $temp2 = 1 To $array2[0][0]
						If StringRegExp($array2[$temp2][1], '^(4|' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & ')$') = 0 And StringRegExp($array2[$temp2][5], StringSplit($clan[$c][1], ':')[4]) = 1 Then
							$temp1[0][2] = StringReplace($temp1[0][2], '|', $array2[$temp2][3] & ',|')
							TCPSend($MainSocket, 'clientpoke clid=' & $array2[$temp2][2] & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
									']\sshowed\sand\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] + 60 * 15 - $StampTime) / 60) & '\sminutes.\sPlease\sjoin\sCW\sclannels!' & @CRLF)
						EndIf
						If StringRegExp($array2[$temp2][1], '^(' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & ')$') = 1 And StringRegExp($array2[$temp2][5], StringSplit($clan[$c][1], ':')[4]) = 0 Then
							$temp1[0][3] &= $array2[$temp2][3] & ','
						EndIf
						If StringRegExp($array2[$temp2][1], '^4$') = 1 And StringRegExp($array2[$temp2][5], StringSplit($clan[$c][1], ':')[4]) = 1 Then
							$temp1[0][2] &= $array2[$temp2][3] & ','
						EndIf
						If StringRegExp($array2[$temp2][1], '^(' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & ')$') = 1 And StringRegExp($array2[$temp2][5], StringSplit($clan[$c][1], ':')[4]) = 1 Then
							$temp1[0][1] &= $array2[$temp2][3] & ','
						EndIf
					Next
					_Excel_RangeInsert(_Excel_SheetList($oBook)[1][1], _Excel_ColumnToLetter(($c - 2) * 4 + 1) & $i + 2 & ':' & _Excel_ColumnToLetter(($c - 2) * 4 + 4) & $i + 2, $xlShiftDown)
					_Excel_RangeWrite($oBook, 2, $temp1, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & $i + 2)
					Local $temp = 0, $temp1 = 0, $temp2 = 0
				Else
					Local $temp1 = [[',', ',', ',|,', ',']]
					_Excel_RangeInsert(_Excel_SheetList($oBook)[1][1], _Excel_ColumnToLetter(($c - 2) * 4 + 1) & $i + 2 & ':' & _Excel_ColumnToLetter(($c - 2) * 4 + 4) & $i + 2, $xlShiftDown)
					_Excel_RangeWrite($oBook, 2, $temp1, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & $i + 2)
				EndIf
			Next
		ElseIf _Excel_RangeRead($oBook, 2, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '1') <> 0 And _Excel_RangeRead($oBook, 2, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '1') <> '' And StringInStr(_Excel_RangeRead($oBook, 2, _Excel_ColumnToLetter(($c - 2) * 4 + 3) & '1'), @MON & '\' & @MDAY & '\' & @YEAR) = 0 Then
			__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] Write new day CW report.', 1)
			_Excel_RangeInsert(_Excel_SheetList($oBook)[1][1], 'A' & ':' & _Excel_ColumnToLetter(($clan[0][0] - 1) * 4), $xlShiftToRight)
			TCPSend($MainSocket, 'quit' & @CRLF & @CRLF)
			TCPShutdown()
			$temp = _ClanWar($array2, $oBook)
			Return SetError(@error, @extended, $temp)
		ElseIf _Excel_RangeRead($oBook, 2, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '1') <> 0 And _Excel_RangeRead($oBook, 2, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '1') <> '' And StringInStr(_Excel_RangeRead($oBook, 2, _Excel_ColumnToLetter(($c - 2) * 4 + 3) & '1'), @MON & '\' & @MDAY & '\' & @YEAR) <> 0 Then
			__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] Update CW report.', 1)
			$table1 = _Excel_RangeRead($oBook, 2, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '1')
			$table1 = _Excel_RangeRead($oBook, 2, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '2:' & _Excel_ColumnToLetter(($c - 2) * 4 + 4) & $table1 * 2 + 1)
			;_ArrayDisplay($table1)
			For $i = 1 To $table[0][0]
				Local $j = UBound($table1)
				For $k = 0 To $j - 1
					If Mod($k, 2) = 0 And $table[$i][0] = $table1[$k][0] And $table1[$k][1] > $table[$i][1] - 60 * 10 And $table1[$k][1] < $table[$i][1] + 60 * 10 Then
						If $table1[$k][1] <> $table[$i][1] Then $table1[$k][1] = $table[$i][1]
						If ($StampTime < $table[$i][1] - (10 * 60)) And ($StampTime) > $table[$i][1] - 15 * 60 Then
							__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] CW code 1:' & _StringFormatTime('%c', $table[$i][1]) & ' ' & $table[$i][2])
							For $temp1 = 1 To UBound($array2) - 1
								If StringRegExp($array2[$temp1][5], StringSplit($clan[$c][1], ':')[4]) = 1 And StringRegExp($array2[$temp1][1], '' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & '') = 0 Then
									TCPSend($MainSocket, 'clientpoke clid=' & $array2[$temp1][2] & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
											']\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] - $StampTime) / 60) & '\sminutes.\sBe\sready!' & @CRLF)
									ConsoleWrite('poke ' & $array2[$temp1][0] & ' to ready')
								EndIf
							Next
							Do
								TCPSend($MainSocket, 'whoami' & @CRLF)
								Sleep(500)
								$temp2 = TCPRecv($MainSocket, 1000000)
								$temp2 = StringRegExp($temp2, 'client_id=\K\d+(?= )', 3)
							Until IsArray($temp2) And UBound($temp2, 0) = 1
							$temp2 = $temp2[0]
							For $cwc = 1 To StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[0]
								$temp1 = StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[$cwc]
								TCPSend($MainSocket, 'clientmove clid=' & $temp2 & ' cid=' & $temp1 & @CRLF)
								TCPSend($MainSocket, 'sendtextmessage targetmode=2 target=' & $temp1 & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
										']\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] - $StampTime) / 60) & '\sminutes.\sThere\sare\s' & UBound(_ArrayFindAll($array2, StringSplit($clan[$c][1], ':')[4], 1, 0, 0, 3, 5)) & '\s' & $clan[$c][0] & '\splayers\son\sTS3' & @CRLF)
							Next
						EndIf
						If $StampTime < $table1[$k][1] And $StampTime >= $table1[$k][1] - 60 * 10 Then
							__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] CW code 2:' & _StringFormatTime('%c', $table[$i][1]) & ' ' & $table[$i][2])
							$temp = $table1[$k + 1][0]
							For $temp1 = 1 To $array2[0][0]
								If StringRegExp($array2[$temp1][1], '' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & '') = 1 And StringInStr($temp, $array2[$temp1][3] & ',') = 0 And StringRegExp($array2[$temp1][5], StringSplit($clan[$c][1], ':')[4]) = 1 Then
									$temp &= $array2[$temp1][3] & ','
								EndIf
								If StringRegExp($array2[$temp1][5], StringSplit($clan[$c][1], ':')[4]) = 1 And StringRegExp($array2[$temp1][1], '' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & '') = 0 Then
									TCPSend($MainSocket, 'clientpoke clid=' & $array2[$temp1][2] & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
											']\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] - $StampTime) / 60) & '\sminutes.\sPlease\sjoin\sCW\sclannels!' & @CRLF)
									ConsoleWrite('poke ' & $array2[$temp1][0] & ' to join')
								EndIf
							Next
							Do
								TCPSend($MainSocket, 'whoami' & @CRLF)
								Sleep(500)
								$temp2 = TCPRecv($MainSocket, 1000000)
								$temp2 = StringRegExp($temp2, 'client_id=\K\d+(?= )', 3)
							Until IsArray($temp2) And UBound($temp2, 0) = 1
							$temp2 = $temp2[0]
							For $cwc = 1 To StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[0]
								$temp1 = StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[$cwc]
								TCPSend($MainSocket, 'clientmove clid=' & $temp2 & ' cid=' & $temp1 & @CRLF)
								TCPSend($MainSocket, 'sendtextmessage targetmode=2 target=' & $temp1 & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
										']\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] - $StampTime) / 60) & '\sminutes.\sThere\sare\s' & UBound(_ArrayFindAll($array2, StringSplit($clan[$c][1], ':')[4], 1, 0, 0, 3, 5)) & '\s' & $clan[$c][0] & '\splayers\son\sTS3' & @CRLF)
							Next
							$table1[$k + 1][0] = $temp
							Local $temp = 0, $temp1 = 0, $temp2 = 0
						ElseIf $StampTime >= $table1[$k][1] And $StampTime <= $table1[$k][1] + 60 * 10 Then
							__LogWrite(FileOpen(@ScriptDir & "\temp\CWLog.log", 9), '[TS3Main] CW code 3:' & _StringFormatTime('%c', $table[$i][1]) & ' ' & $table[$i][2])
							Do
								TCPSend($MainSocket, 'whoami' & @CRLF)
								Sleep(500)
								$temp2 = TCPRecv($MainSocket, 1000000)
								$temp2 = StringRegExp($temp2, 'client_id=\K\d+(?= )', 3)
							Until IsArray($temp2) And UBound($temp2, 0) = 1
							$temp2 = $temp2[0]
							For $cwc = 1 To StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[0]
								$temp1 = StringSplit(StringSplit($clan[$c][1], ':')[3], ',')[$cwc]
								TCPSend($MainSocket, 'clientmove clid=' & $temp2 & ' cid=' & $temp1 & @CRLF)
								TCPSend($MainSocket, 'sendtextmessage targetmode=2 target=' & $temp1 & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
										']\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] + 60 * 15 - $StampTime) / 60) & '\sminutes.\sThere\sare\s' & UBound(_ArrayFindAll($array2, StringSplit($clan[$c][1], ':')[4], 1, 0, 0, 3, 5)) & '\s' & $clan[$c][0] & '\splayers\son\sTS3' & @CRLF)
							Next
							$temp = ''
							Local $temp1 = [[',', ',', '|', ',']]
							For $temp2 = 1 To $array2[0][0]
								If StringInStr($table1[$k + 1][2], $array2[$temp2][3] & ',') = 0 And StringRegExp($array2[$temp2][1], '^(4|' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & ')$') = 0 And StringRegExp($array2[$temp2][5], StringSplit($clan[$c][1], ':')[4]) = 1 Then
									$table1[$k + 1][2] = StringReplace($table1[$k + 1][2], '|', $array2[$temp2][3] & ',|')
									TCPSend($MainSocket, 'clientpoke clid=' & $array2[$temp2][2] & ' msg=(' & $clan[$c][0] & ')\sClan\sWar\son\s' & StringRegExpReplace(StringRegExpReplace($table[$i][2], '\s', '\\s'), ',.*?,', '\\s\\p\\s\[') & _
											']\sshowed\sand\swill\sstart\sin\s' & StringFormat('%.0f', ($table[$i][1] + 60 * 15 - $StampTime) / 60) & '\sminutes.\sPlease\sjoin\sCW\sclannels!' & @CRLF)
								ElseIf StringInStr($table1[$k + 1][3], $array2[$temp2][3] & ',') = 0 And StringRegExp($array2[$temp2][1], '^(' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & ')$') = 1 And StringRegExp($array2[$temp2][5], StringSplit($clan[$c][1], ':')[4]) = 0 Then
									$table1[$k + 1][3] &= $array2[$temp2][3] & ','
								ElseIf StringInStr($table1[$k + 1][2], $array2[$temp2][3] & ',') = 0 And StringRegExp($array2[$temp2][1], '^4$') = 0 And StringRegExp($array2[$temp2][5], StringSplit($clan[$c][1], ':')[4]) = 1 Then
									$table1[$k + 1][2] &= $array2[$temp2][3] & ','
								ElseIf StringInStr($table1[$k + 1][1], $array2[$temp2][3] & ',') = 0 And StringRegExp($array2[$temp2][1], '^(' & StringReplace(StringSplit($clan[$c][1], ':')[3], ',', '|') & ')$') = 1 And StringRegExp($array2[$temp2][5], StringSplit($clan[$c][1], ':')[4]) = 1 Then
									$table1[$k + 1][1] &= $array2[$temp2][3] & ','
								EndIf
							Next
						EndIf
						ExitLoop
						Local $temp = 0, $temp1 = 0, $temp2 = 0
					ElseIf $k = $j - 1 Then
						Local $temp1 = [$table[$i][0], $table[$i][1], $table[$i][2], $table[$i][3]]
						_ArrayAdd2D($table1, $temp1)
						Local $temp1 = [',', ',', ',|,', ',']
						_ArrayAdd2D($table1, $temp1)
						$j = UBound($table1)
						;_ArrayDisplay($table1)
						_Excel_RangeWrite($oBook, 2, ($j) / 2, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '1')
						$k = -1
					EndIf
					Local $temp = 0, $temp1 = 0, $temp2 = 0
				Next
			Next
			_Excel_RangeWrite($oBook, 2, $table1, _Excel_ColumnToLetter(($c - 2) * 4 + 1) & '2')
		EndIf
		;MsgBox(0, '', TCPRecv($MainSocket, 10000000))

		TCPSend($MainSocket, 'quit' & @CRLF & @CRLF)
		TCPShutdown()
	Next
;~ 	For $i = 1 To $table[0][0]
;~ 		If $table[$i][1] < $StampTime And $table[$i][1] + $SleepPeriod > $StampTime Then $table[$i][1] = $table[$i][1] + $SleepPeriod
;~ 	Next
;~ 	_ArraySort($table, 0, 1, 0, 1)
;~ 	If $table[1][1] > $StampTime Then Return SetError(0, UBound($table), $table[1][1])
	Return SetError(0, UBound($table), 0)
	#EndRegion
	#cs For $i=1 To $table[0][0]
		If $table[$i][6]<>'meeting_engagement' Then
		If  $table[$i][1]-$temp1>=(10*60) Then
		If _Excel_RangeRead($oBook,2,'A1') ='' or _Excel_RangeRead($oBook,2,'A1') ='0' Or _Excel_RangeRead($oBook,2,'A2') ='SNS_A' or _Excel_RangeRead($oBook,2,'A3')<>@MON&'\'&@MDAY&'\'&@YEAR Then
		If $table[$i][1]-$temp1>=60*10 Then
		Local $temp2[$table[0][0]+1][4]
		[[$table[0][0],'SNS_A',@MON&'\'&@MDAY&'\'&@YEAR]]
		For $t=
		_Excel_RangeWrite($oBook,2,$temp2,'A1')
		
		ElseIf _Excel_RangeRead($oBook,2,'A1') <>'' And _Excel_RangeRead($oBook,2,'A1') <>'0' and _Excel_RangeRead($oBook,2,'A2') ='SNS_A' And _Excel_RangeRead($oBook,2,'A3') =@MON&'\'&@MDAY&'\'&@YEAR then
		For $k=1 To _Excel_RangeRead($oBook,3,'A1')
		
		For $j=1 To $array2[0][0]
		
		Next
		EndIf
		EndIf
		EndIf
	#ce Next

	;_ArrayDisplay($temp)
EndFunc   ;==>_ClanWar

Func inetread_timeout(Const $link, $inet_op = 1, $retry = 0, $tempfile = @TempDir, $timeout = 10)
	Local $Timer, $PID
	FileDelete($tempfile & '\inet_temp.html')
	$Timer = TimerInit()
	If @Compiled = 1 Then
		$PID = Run(@AutoItExe & ' ' & $link & ' ' & $inet_op & ' ' & $tempfile)
	Else
		$PID = Run('c:\users\kieuq\desktop\ts3.exe ' & $link & ' ' & $inet_op & ' ' & $tempfile)
	EndIf
	While 1
		Sleep(500)
		If $PID = 0 Or TimerDiff($Timer) / 1000 >= $timeout Then
			;If $PID=0 Then MsgBox(0,'error',@error)
			;MsgBox(0,'1','1')
			ProcessClose($PID)
			FileDelete($tempfile & '\inet_temp.html')
			If $retry >= 4 Then Return 0
			Return inetread_timeout($link, $inet_op, $retry + 1)
		ElseIf Not ProcessExists($PID) Then
			;MsgBox(0,'pause','pause')
			$PID = 0
			$chars = FileRead($tempfile & '\inet_temp.html')
			If @error Then $chars = 0
			Return $chars
		Else
			Sleep(500)
			ContinueLoop
		EndIf
		ExitLoop
	WEnd
EndFunc   ;==>inetread_timeout

