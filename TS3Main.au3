
#Region ****#~ endfunc_comment=1
;Directives created by AutoIt3Wrapper_GUI * * * *
#AutoIt3Wrapper_Run_Before="C:\Users\kieuq\Desktop\software\new folder (2)\ShowOriginalLine.exe" %in%
#AutoIt3Wrapper_Run_After="C:\Users\kieuq\Desktop\software\new folder (2)\ShowOriginalLine.exe" %in%
#AutoIt3Wrapper_Icon=Work\Icon1.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_AU3Check_Parameters= -w 1
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/sci 1
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/striponly /sci 1 /mo
#~ endfunc_comment=1
#AutoIt3Wrapper_run_debug_mode=Y
#EndRegion ****#~ endfunc_comment=1
;#RequireAdmin
#AutoIt3Wrapper_Run_Debug=N
#include <String.au3>
#include <array.au3>
#include <File.au3>
#include <_Array2D.au3>
#include <Date.au3>
#include <_HTML_Ex.au3>
#include <UnixTime.au3>
#include <_Excel_Rewrite.au3>
#include <Timers.au3>
#include "C:\Users\kieuq\Desktop\software\new folder (2)\_AutoItErrorTrap.au3"
Opt("TrayIconDebug", 1)
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

Global $StampTime = _TimeMakeStamp(), $serverfile = SetServer('xlsxfile'), $serverIP = SetServer('ServerIP'), $QueryPort = SetServer('queryport'), $QueryTransfer = SetServer('querytransfer')
Global $QueryVituralServer = SetServer('queryvituralserver'), $QueryUser = SetServer('queryuser'), $QueryPass = SetServer('querypass'), $CopyTo = SetServer('copyto')
Global $clan = [['A', 1000002504],['R', 1000001787],['C', 1000006228]]

$temp = IniReadSection(@ScriptDir & '\temp\INFO.ini', 'PassUsers')
If @error Then
	Global $PassUser[1][3] = [[0, 0, 0]]
Else
	Local $temp1[$temp[0][0] + 1][3]
	$temp1[0][0] = $temp[0][0]
	For $i = 1 To $temp[0][0]
		$temp1[$i][2] = $temp[$i][0]
		$temp1[$i][0] = StringSplit($temp[$i][1], ':')[1]
		$temp1[$i][1] = StringSplit($temp[$i][1], ':')[2]
	Next
	$PassUser = $temp1
	$temp1 = 0
EndIf

$temp = 0
test()
start1()
Func test()
	;$temp[[4,'','','',''],['Dungwad',4,2,2,1393940501],['Hardcore125',5,9,3,1393944308],['Boombab',29,29,4,1393901643],['Wibs E5',5,114,5,1393897420]]
EndFunc   ;==>test

Func SetServer($c, $mode = 0)
	Do
		$iniA = @ScriptDir & '\temp\INFO.ini'
		$iniA = IniReadSection($iniA, 'Server')
		If @error Then
			FileInstall('.\temp\INFO - Copy.ini', @ScriptDir & '\temp\INFO.ini', 1)
		EndIf
	Until IsArray($iniA)
	Local $this, $a[1][1]
	$c = StringLower($c)
	$this = _ArraySearch($iniA, $c, 0, 0, 0, 1, 0)
	If $this <> -1 And $mode = 0 Then
		$this = $iniA[$this][1]
		If $this = '' Then
			$this = SetServer($c, 1)
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
			Case 'copyto'
				$this = ''
		EndSwitch
	EndIf
	Return $this
EndFunc   ;==>SetServer

Func __LogWrite($file, $text, $mode)
	Local $Ram = ProcessGetStats()
	_FileWriteLog($file, Int($Ram[0] / 1048576) & "MB | " & Int($Ram[1] / 1048576) & 'MB ' & $text, $mode)
	FileClose($file)
EndFunc   ;==>__LogWrite

Func start1()
	Local $k = 1
	If @Compiled = 1 Then
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
		Do
			Sleep(500)
			If @HOUR = 23 And $k = 0 And ProcessExists('Last24h.exe') = 0 Then
				$k = 1
				__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call Last24h.', 1)
				Run(@ScriptDir & '\Last24h.exe', "", @SW_HIDE, 0x8)
			ElseIf @HOUR <> 23 And $k = 1 Then
				$k = 0
			EndIf
			If TimerDiff($begin) >= 5 * 1000 * 60 Then
				$StampTime = _TimeMakeStamp()
				$temp = @ScriptDir & '\temp\INFO.ini'
				$temp = IniReadSection($temp, 'Server')
				$serverfile = SetServer('xlsxfile')
				__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call Main.', 1)
				$begin = TimerInit()
				main()
				If $CopyTo <> '' Then
					$temp = @ScriptDir & '\temp\INFO.ini'
					$temp = IniReadSection($temp, 'Server')
					$serverfile = SetServer('xlsxfile')
					__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Copy to ' & $CopyTo & ' error(' & FileCopy($serverfile, $CopyTo & '\TS3alpha.xlsx', 9) & ').', 1)
				EndIf
			EndIf
		Until 0
	ElseIf @Compiled = 0 Then
		Local $begin = TimerInit(), $non = 0
		If $CopyTo <> '' Then
			FileCopy($serverfile, $CopyTo & '\TS3alpha.xlsx', 9)
		EndIf
		main()
		Do
			Sleep(500)
			If TimerDiff($begin) >= 1 * 1000 * 60 Then
				$StampTime = _TimeMakeStamp()
				$temp = @ScriptDir & '\temp\INFO.ini'
				$temp = IniReadSection($temp, 'Server')
				$serverfile = SetServer('xlsxfile')
				main()
				$begin = TimerInit()
				$non = $non + 1
			EndIf
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
		$string = StringReplace($string, 'error id=0 msg=ok', "channel" & @CR & @LF & "::", 1)
		$string = StringReplace($string, 'error id=0 msg=ok', "::" & @CR & @LF & "client" & @CR & @LF & "::", 1)
		$string = StringReplace($string, 'error id=0 msg=ok', "")
		$string = StringReplace($string, '0quit', '0' & @CR & @LF)
;~ $file=FileOpen("D:/quang2.txt",2)
;~ FileWrite($file,$string)
;~ FileWrite($file,@cr&@lf)
;~ ConsoleWrite($string)
;~ FileFlush($file)
;~ FileClose($file)

		$array = StringSplit($string, '::', 1)
		$temp1 = $temp1 + 1
		If $temp1 >= 10 Then
			Local $temp1 = [[0, 0, 0, 0]]
			LogUsers($temp1)
			Return
		EndIf
	Until IsArray($array) = 1 And UBound($array) = 5
	$array[2] = StringReplace($array[2], 'channel_name=', "")
	$array[2] = StringReplace($array[2], 'cid=', '')
	$array[2] = StringReplace($array[2], 'pid=', '')
	$array[2] = StringReplace($array[2], 'total_clients=', '')
	$array[4] = StringReplace($array[4], @CR & @LF, '')
	$array[4] = StringReplace($array[4], 'clid=', '')
	$array[4] = StringReplace($array[4], 'cid=', '')
	$array[4] = StringReplace($array[4], 'client_database_id=', '')
	$array[4] = StringReplace($array[4], 'client_nickname=', '')
	$temp = StringSplit($array[2], '|')

	Global $array1[$temp[0] + 1][6]
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

	Global $array2[$temp[0] + 1][5]
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
	#cs	If _ArraySearch($array2, 'quangkieu', 0, 0, 0, 1, 1, 0) <> -1 Then
		$temp = _ArraySearch($array2, 'quangkieu', 0, 0, 0, 1, 1, 0)
		$array2[$temp][0] = 'Quang Kieu KTAQ'
	#ce EndIf
	$array2[0][1] = 'cid'
	$array2[0][2] = 'clid'
	$array2[0][3] = 'cdid'
	$array2[0][4] = 'last'


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
				"client_unique_identifier=0= client_nickname=" & $QueryUser & " client_database_id=0 client_created=0 client_lastconnected=" & _TimeGetStamp() & _
				" client_totalconnections=1 client_flag_avatar client_description client_month_bytes_uploaded=0 client_month_bytes_downloaded=0 client_total_bytes_uploaded=0" & _
				" client_total_bytes_downloaded=0 client_icon_id=0 client_base64HashClientUID=0 client_lastip=0error id=0 msg=ok")
		$string = StringReplace($string, 'error id=0 msg=ok', "::", $array2[0][0] - 1)
		$string = StringReplace($string, 'error id=0 msg=ok', '')
		$string = StringReplace($string, '0quit', '0' & @CR & @LF)
		$array = StringSplit($string, "::", 1)
		$temp1 = $temp1 + 1
		If $temp1 >= 10 Then
			Local $temp1 = [[0, 0, 0, 0]]
			LogUsers($temp1)
			Return
		EndIf
	Until IsArray($array) = 1 And UBound($array) = $array2[0][0]
	For $i = 1 To $array[0]
		$temp = StringSplit($array[$i], ' ')
		$array2[$i][4] = StringReplace($temp[5], "client_lastconnected=", '', 1)
	Next
;~ $file=FileOpen("D:/quang2.txt",1)
;~ FileWrite($file,$string)
;~ ConsoleWrite($string)
;~ FileFlush($file)
;~ FileClose($file)
	LogUsers($array2)
	__ImageServer($array1)




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
	IniWriteSection(@ScriptDir & '\temp\INFO.ini', 'PassUsers', '')
	IniWriteSection(@ScriptDir & '\temp\INFO.ini', 'PassUsers', $temp)
	$temp = 0
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

Func WriteNewLine(ByRef $a, ByRef $b, ByRef $c, ByRef $d)
	_Excel_RangeWrite($a, 3, $d, $b[$c][0] & '3')
EndFunc   ;==>WriteNewLine


Func LogUsers(ByRef $c)
	#AutoIt3Wrapper_Run_Debug=Y
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call LogUsers.', 1)
	Local $array2 = $c, $temp, $temp1, $starttime = _TimeMakeStamp(0, 0, 0)
;~ 	ReDim $array2[$array2[0][0]][5]
;~ 	$array2[0][0] = $array2[0][0] - 1
;~ 	;_ArrayDisplay($PassUser)
;~ 	If $PassUser[0][0] >= 2 Then
;~ 		_ArrayDelete($PassUser, 3)
;~ 		$PassUser[0][0] = $PassUser[0][0] - 1
;~ 	EndIf
	Do
		ObjGet("", "Excel.Application")
		$temp = @error
		Sleep(100)
	Until $temp <> 0
	$oEx = _Excel_Open(False)
	ConsoleWrite(@error & ' ' & @ScriptLineNumber)
	Sleep(500)
	$oBook = _Excel_BookOpen($oEx, $serverfile, False, True)
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

							$table2[0][2] = $StampTime - (15 * 60)
							WriteNewLine($oBook, $table1, $temp, $table2)
							InsertNewLine($oBook, $table1, $temp)
							Dim $table2 = [['', $StampTime - (15 * 60) + 1, $StampTime, $PassUser[$i][1]]]
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

	;_ClanWar($c,$oBook)

	_Excel_BookSave($oBook)
	_Excel_BookClose($oBook)
	_Excel_Close($oEx)
	#AutoIt3Wrapper_Run_Debug=N
EndFunc   ;==>LogUsers

Func __ImageServer(ByRef $array1)
	Local $temp, $temp1, $imagefile
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[TS3Main] Call ServerImage.', 1)
	Local $imagefile = FileOpen(@ScriptDir & '\temp\TS3_Image\' & @YEAR & '-' & @MON & '-' & @MDAY & '_' & @HOUR & @MIN & @SEC & @MSEC & '.txt', 10)
	;$imagefile = FileOpen('D:/quang4.txt', 10)
	FileWriteLine($imagefile, 'Server image at ' & @MON & '/' & @MDAY & '/' & @YEAR & ' ' & @HOUR & ':' & @MIN & ':' & @SEC & ':' & @CR & @LF & @CR & @LF & '(...):channel' & @CR & @LF & @CR & @LF)
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
			ConsoleWrite($temp2 & ' ' & $array1[$temp2][2] & @CR & @LF)
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
		FileWrite($imagefile, '(' & $array1[$i][0] & ')' & @CR & @LF)
		$temp1 = 1
		Do
			$temp2 = _ArraySearch($array2, $array1[$i][1], $temp1, $array2[0][0], 0, 0, 1, 1)
			If $temp2 <> -1 Then
				For $j = 1 To $temp + 2
					FileWrite($imagefile, @TAB)
				Next
				FileWrite($imagefile, '(' & _StringFormatTime("%c", $array2[$temp2][4]) & ')' & @TAB & $array2[$temp2][0] & @CR & @LF)
				$temp1 = $temp2 + 1
			EndIf
		Until $temp2 = -1
		$temp1 = 0
		$temp2 = 0
	Next
	FileFlush($imagefile)
	FileClose($imagefile)
EndFunc   ;==>__ImageServer

Func _ClanWar(Const ByRef $array2, ByRef $oBook )
	Local $temp,$temp1,$temp2,$table
	;$temp=_Date_Time_GetTimeZoneInformation()
	For $i=0 To 2
		$temp1=_HTML_GetSource('http://api.worldoftanks.com/wot/clan/battles/?application_id=8b379c513ce91c4e9ddfe538a12b0e67&clan_id=1000002504')
		;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
		Sleep(500)
		$temp2=_HTML_GetSource('http://api.worldoftanks.com/wot/clan/battles/?application_id=8b379c513ce91c4e9ddfe538a12b0e67&clan_id=1000002504')
		;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
		If $temp1<>$temp2 Then ExitLoop
	Next
	If @error Or StringInStr($temp2,'"error"')<>0 Then Return SetError(1,0,0);no connection or wrong ClanID
	$temp1=StringRegExp($temp2,'(?!"provinces":\[)(,*"\w+_\d+")+(?=],"started")',3)
	$temp=UBound($temp1)
	If $temp=0 Then Return SetError(0,0,0);no CW battle
	$table=$temp2
	For $i=0 To 2
		$temp1=_HTML_GetSource('http://cw.wotapi.ru/na/1000002504?type=json')
		;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
		Sleep(500)
		$temp2=_HTML_GetSource('http://cw.wotapi.ru/na/1000002504?type=json')
		;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
		If $temp1<>$temp2 Then ExitLoop
	Next
	If @error Or StringInStr($temp2,'"error"')<>0 Then Return SetError(1,0,0); no connection or wrong ClanID
	If $temp<>StringRegExp($temp2,'(?!"battles_count":)\d+(?=,"last_update":)',3)[0] Then Return SetError(2,0,0);problem in sync between 2 servers
	$temp1=$temp2
	$temp2=$table
	;started | time | region | province | map | code | type | server |
	;	/		/							/	/		/
	If $temp<>0 Then
		Local $table[$temp+1][8]
		$table[0][0]=$temp
		$temp=$temp1
		$temp1=StringRegExp($temp2,'"provinces":\[\K.*?(?=],"started")',3)
		For $i=1 To $table[0][0]
			$table[$i][5]=StringReplace($temp1[$i-1],'"','')
		Next
		$temp1=StringRegExp($temp2,'"started":\K.*?(?=,"private")',3)
		For $i=1 To $table[0][0]
			$table[$i][0]=$temp1[$i-1]
		Next
		$temp1=StringRegExp($temp2,'"time":\K.*?(?=,"arenas")',3)
		For $i=1 To $table[0][0]
			$table[$i][1]=$temp1[$i-1]
		Next
		$temp1=StringRegExp($temp2,'"name_i18n":"\K.*?(?=","name")',3)
		For $i=1 To $table[0][0]
			$table[$i][4]=$temp1[$i-1]
		Next
		$temp1=StringRegExp($temp2,'"type":"\K.*?(?="})',3)
		For $i=1 To $table[0][0]
			$table[$i][6]=$temp1[$i-1]
		Next
	EndIf

	$temp2=$temp
	;code | type | region | province | server |
	Local $temp[$table[0][0]+1][5]
	$temp[0][0]=$table[0][0]
	If $table[0][0]<>0 Then
		$temp1=StringRegExp($temp2,'e":{"name".*?"id":"\K.*?(?=",)',3)
		For $i=1 To $temp[0][0]
			$temp[$i][0]=$temp1[$i-1]
		Next
		$temp1=StringRegExp($temp2,'"type":"\K.*?(?=",)',3)
		For $i=1 To $temp[0][0]
			$temp[$i][1]=$temp1[$i-1]
		Next
		$temp1=StringRegExp($temp2,'n":{"name":"\K.*?(?=","id")',3)
		For $i=1 To $table[0][0]
			$temp[$i][2]=$temp1[$i-1]
		Next
		$temp1=StringRegExp($temp2,'e":{"name":"\K.*?(?=",)',3)
		For $i=1 To $temp[0][0]
			$temp[$i][3]=$temp1[$i-1]
		Next
		$temp1=StringRegExp($temp2,'r":"\K.*?(?="})',3)
		For $i=1 To $temp[0][0]
			$temp[$i][4]=$temp1[$i-1]
		Next
	EndIf
	;started | time | region | province | map | code | type | server |
	;	0		1		2			3		4	5		6		7
	;code | type | region | province | server |
	For $i=$table[0][0] To 1 Step -1
		If $table[$i][6]<>'meeting_engagement' Then
		$temp1=_ArrayFindAll($temp,$table[$i][5],1)
			For $j=0 To UBound($temp1)-1
				If $temp[$temp1[$j]][1]<>'meeting_engagement' Then
					$table[$i][2]=$temp[$temp1[$j]][2]
					$table[$i][3]=$temp[$temp1[$j]][3]
					$table[$i][7]=$temp[$temp1[$j]][4]
					_ArraySwap($table[$i][0],$table[$i][5])
					$table[$i][2]=$table[$i][4]&','&$table[$i][2]&','&$table[$i][3]
					$table[$i][3]=$table[$i][7]
					ExitLoop
				EndIf
			Next
		ElseIf $table[$i][6]='meeting_engagement' Then
			_ArrayDelete($table,$i)
			$table[0][0]=UBound($table)-1
		EndIf
	Next
	;code | time | map,region,province | server
	ReDim $table[$table[0][0]+1][
	Local $table1
	If _Excel_RangeRead($oBook,2,'A1')=0 Or _Excel_RangeRead($oBook,2,'A1')='' Then
		$table[0][1]=0100
		$table[0][2]=@MON&'\'&@MDAY&'\'&@YEAR
		_Excel_RangeWrite($oBook,2,$table







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
EndFunc
