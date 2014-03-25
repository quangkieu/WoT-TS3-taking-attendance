#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Work\Icon1.ico
#AutoIt3Wrapper_Outfile=ExcelIndex.exe
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
;#AutoIt3Wrapper_run_debug_mode=Y
Opt("TrayIconDebug", 1) ;0=no info, 1=debug line info
#Region ****#~ endfunc_comment=1
;Directives created by AutoIt3Wrapper_GUI * * * *
#~ endfunc_comment=1
#EndRegion ****#~ endfunc_comment=1
#Region ****#~ endfunc_comment=1
;Directives created by AutoIt3Wrapper_GUI * * * *
#~ endfunc_comment=1
#EndRegion ****#~ endfunc_comment=1
#include <ExcelConstants.au3>
#include <String.au3>
#include <array.au3>
#include <File.au3>
#include <_Array2D.au3>
#include <UnixTime.au3>
#include "_Excel_Rewrite.au3"
#include <_Array2D.au3>
;#include <_Excel_Rewrite.au3>

$temp = @ScriptDir & '\temp\INFO.ini'
$temp = IniReadSection($temp, 'Server')
Global $serverfile = SetServer($temp, 'xlsxfile'), $serverIP = SetServer($temp, 'ServerIP'), $QueryPort = SetServer($temp, 'queryport'), $QueryTransfer = SetServer($temp, 'querytransfer')
Global $QueryVituralServer = SetServer($temp, 'queryvituralserver'), $QueryUser = SetServer($temp, 'queryuser'), $QueryPass = SetServer($temp, 'querypass')

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
		EndSwitch
	EndIf
	Return $this
EndFunc   ;==>SetServer

Func __LogWrite($file, $text, $mode)
	Local $Ram = ProcessGetStats()
	_FileWriteLog($file, Int($Ram[0] / 1048576) & "MB | " & Int($Ram[1] / 1048576) & 'MB ' & $text, $mode)
	FileClose($file)
EndFunc   ;==>__LogWrite

main()
;IndexExcel()

Func main()
	Local $mainSocket, $string = '', $temp, $table[900][4], $i = 0, $j = 0, $error = 0, $temp1 = 0, $temp2 = 0
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[ExcelIndex] Get database from server.', 1)
	TCPStartup()
	Do
		$mainSocket = TCPConnect($serverIP, $QueryPort)
	Until Not (@error)
	TCPSend($mainSocket, "login client_login_name=" & $QueryUser & " client_login_password=" & $QueryPass & @CRLF)
	TCPSend($mainSocket, "use 1" & @CRLF)
	Sleep(4000)
	TCPRecv($mainSocket, 1000000)
	Do
		TCPSend($mainSocket, 'clientdblist start=' & $temp1 & @CRLF)
		ConsoleWrite($temp1 & @CRLF)
		Sleep(500)
		$temp = TCPRecv($mainSocket, 10000000000)
		While Not (StringRegExp($temp, '(?i)error id=.*?' & @LF))
			Sleep(500)
			$temp = $temp & TCPRecv($mainSocket, 1000000)
		WEnd
		If StringInStr($temp, 'error id=1281') <> 0 Then
			$temp = ''
			$temp1 = 0
			$table[$temp2[1] + 1][0] = 'done'
			$temp2 = 0
			$error = -1
			ExitLoop
		EndIf
		$temp = StringReplace($temp, @CR, '')
		$temp = StringReplace($temp, @LF, '')
		$temp = StringReplace($temp, 'client_unique_identifier=ServerQuery ', '')
		$temp = StringReplace($temp, 'TS3Welcome to the TeamSpeak 3 ServerQuery interface, type "help" for a list of commands and "help <command>" ' & _
				'for information on a specific command.error id=0 msg=ok', "")
		$temp = StringRegExpReplace($temp, '(?i)client_unique.*?= ', '')
		$temp = StringRegExpReplace($temp, '(?i)client_created.*? ', '')
		$temp = StringRegExpReplace($temp, '(?i)client_total.*?\|', '|')
		$temp = StringRegExpReplace($temp, '(?i)client_total.*?error', 'error')
		$temp = StringReplace($temp, 'error id=0 msg=ok', '')
		$temp = StringReplace($temp, '|', @CRLF)
		$temp = StringRegExpReplace($temp, ' client_.*?=', ' ')
		$temp = StringRegExpReplace($temp, 'cldbid=', '')
		$temp = StringSplit($temp, @CR & @LF, 1)
		For $k = 1 To $temp[0]
			$temp2 = StringSplit($temp[$k], ' ')
			$table[$temp2[1]][0] = $temp2[1]
			$table[$temp2[1]][1] = $temp2[2]
			$table[$temp2[1]][2] = $temp2[3]
			;ConsoleWrite($temp2[1]&' '&$table[$temp2[1]][1]&@CRLF)
		Next
		$temp1 = $temp1 + $temp[0]
		;$table[$k+1]='done'
		ConsoleWrite($temp)
		$temp = ''
	Until ($error = -1)
	$temp = _ArraySearch($table, 'done')
	ConsoleWrite($temp)
	ReDim $table[$temp][4]
	;;_ArrayDisplay($table)
	$table[0][0] = UBound($table) - 1
	TCPRecv($mainSocket, 10000)
	For $i = $table[0][0] To 1 Step -1
		;ConsoleWrite($i)
		$table[$i][1] = StringReplace($table[$i][1], '\s', ' ')
		$table[$i][1] = StringReplace($table[$i][1], '\/', '/')
		$table[$i][1] = StringReplace($table[$i][1], '\p', '|')
		$table[$i][1] = StringReplace($table[$i][1], '\\', '\')
		$table[$i][1] = StringRegExpReplace($table[$i][1], '(?i)[^a-z0-9 ]', '')
		$temp = StringToASCIIArray($table[$i][1])
		For $j = UBound($temp) - 1 To 0 Step -1
			If $temp[$j] >= 128 Then _ArrayDelete($temp, $j)
		Next
		$table[$i][1] = StringFromASCIIArray($temp)
		If $table[$i][0] = '' Then
			$temp = ''
			TCPSend($mainSocket, 'clientdbinfo cldbid=' & $i & @CRLF)
			While StringRegExp($temp, 'error id=.*?' & @LF) = 0
				Sleep(500)
				$temp = $temp & TCPRecv($mainSocket, 1000000)
			WEnd
			;MsgBox(0,'',$temp)
			If StringInStr($temp, 'error id=0') <> 0 Then
				$temp = StringRegExpReplace($temp, '(*ANYCRLF)', '')
				$temp = StringReplace($temp, 'TS3Welcome to the TeamSpeak 3 ServerQuery interface, type "help" for a list of commands and "help <command>" ' & _
						'for information on a specific command.error id=0 msg=ok', "")
				$temp = StringRegExpReplace($temp, '(?i)client_unique.*?= ', '')
				$temp = StringRegExpReplace($temp, '(?i)client_created.*? ', '')
				$temp = StringRegExpReplace($temp, '(?i)client_total.*?\|', '|')
				$temp = StringRegExpReplace($temp, '(?i)client_total.*?error', 'error')
				$temp = StringRegExpReplace($temp, '(?i)client_.*?=', '')
				$temp = StringReplace($temp, 'error id=0 msg=ok', '')
				$temp = StringReplace($temp, 'cldbid=', '')
				$temp = StringSplit($temp, ' ')
				$table[$i][0] = $temp[2]
				$table[$i][1] = $temp[1]
				$table[$i][2] = $temp[3]
				$table[$i][3] = ''
				$table[$i][1] = StringReplace($table[$i][1], '\s', ' ')
				$table[$i][1] = StringReplace($table[$i][1], '\/', '/')
				$table[$i][1] = StringReplace($table[$i][1], '\p', '|')
				$table[$i][1] = StringReplace($table[$i][1], '\\', '\')
				$table[$i][1] = StringRegExpReplace($table[$i][1], '(?i)[^a-z0-9 ]', '')
				$temp = StringToASCIIArray($table[$i][1])
				For $j = UBound($temp) - 1 To 0 Step -1
					If $temp[$j] >= 128 Then _ArrayDelete($temp, $j)
				Next
				$table[$i][1] = StringFromASCIIArray($temp)
			Else
				_ArrayDelete($table, $i)
			EndIf
		EndIf
	Next

	$table[0][0] = UBound($table) - 1
	For $j = 9 To 12
		TCPSend($mainSocket, 'servergroupclientlist sgid=' & $j & @CRLF)
		$temp = ''
		While StringRegExp($temp, 'error id=.*?' & @LF) = 0
			Sleep(500)
			$temp = $temp & TCPRecv($mainSocket, 1000000)
		WEnd
		$temp = StringRegExpReplace($temp, '(*ANYCRLF)', '')
		$temp = StringReplace($temp, 'error id=0 msg=ok', '')
		$temp = StringReplace($temp, 'cldbid=', '')
		$temp = StringSplit($temp, '|')
		For $i = 1 To $table[0][0]
			;ConsoleWrite($i)
			If _ArraySearch($temp, $table[$i][0], 1, 0, 0, 0, 1, 0) <> -1 Then
				$table[$i][3] = $table[$i][3] & '1'
			Else
				$table[$i][3] = $table[$i][3] & '0'
			EndIf
		Next
	Next
	;MsgBox
	;;_ArrayDisplay($table)
	;MsgBox(0,'',$table[$table[0][0]][1]-1)
	;MsgBox(0,'',$table[$table[0][0]][1])
	;_ArrayDisplay($table)
	TCPSend($mainSocket, 'quit')
	TCPCloseSocket($mainSocket)
	TCPShutdown()
	;_ArrayDisplay($table)
	IndexExcel($table)
	;_ArrayDisplay($table)
	;Sleep(5000)
EndFunc   ;==>main
Func IndexExcel(ByRef $c)
	Do
		ObjGet("", "Excel.Application")
		$temp = @error
		Sleep(100)
	Until $temp <> 0
	;----$table		=|CLID	   |Full name|Last connect|Clan
	;----$table1	=|Collumn|CLID	   |Short name	|Clan
	;----$header	=|CLID	   |Full name|Clan		|N
	Local $header[2][4] = [['CDID', 'TS3names|ingameNames', '0000', 0],['Game', 'Start', 'End', 'Channel']], $table = $c
	$temp = $header
	__LogWrite(FileOpen(@ScriptDir & "\temp\AutomationLog.log", 9), '[ExcelIndex] Write to Excel.', 1)
	FileCopy($serverfile, @ScriptDir & '\temp\backup\' & @YEAR & '_' & @MON & '_' & @MDAY & '_TS3.xlsx', 9)
	;If @error Then MsgBox(0, '', @error)
	;$temp = _FileListToArrayRec(@ScriptDir & '\temp\backup\', "*.xlsx|default|default", 1, 0, 1, 0)
	;_ArrayDisplay($temp)
	$oEx = _Excel_Open(False)
	ConsoleWrite(@error & ' line:' & @ScriptLineNumber & @CRLF)
	$oBook = _Excel_BookOpen($oEx, $serverfile, False, True)
	ConsoleWrite(@error & ' line:' & @ScriptLineNumber & @CRLF)
	$table1 = _Excel_RangeRead($oBook, 3, 'A3:D' & _Excel_RangeRead($oBook, 3, 'A1') + 3)
	ConsoleWrite('>Error code: ' & @error & @CRLF & @CRLF & '@ Trace(' & @ScriptLineNumber & ') : $table1=_Excel_RangeRead($oBook,3,''A3:D''&_Excel_RangeRead($oBook,3,''A1'')+3)' & @CRLF) ;### Trace Console
	$temp = 0
	;_ArrayDisplay($table1)
	For $i = UBound($table1) - 1 To 0 Step -1
		If $table1[$i][0] = '' Or $table1[$i][1] = '' Or $table1[$i][2] = '' Then
			_ArrayDelete($table1, $i)
		EndIf
	Next
	If IsArray($table1) And UBound($table1, 0) = 2 And UBound($table1, 2) = 4 And UBound($table1, 1) >= 1 Then
		For $i = UBound($table1) - 1 To 0 Step -1
			If $table1[$i][0] = '' Or $table1[$i][1] = '' Or $table1[$i][2] = '' Then
				_ArrayDelete($table1, $i)
			EndIf
		Next
		For $i = 1 To $table[0][0]
			ConsoleWrite('>@  ' & $i & @CRLF)
			$temp = StringRegExp(StringRegExpReplace($table[$i][1], '(?i)[^a-z0-9 ]', ''), '(?i)([0-9]{0,}[a-z]{1,4}[0-9]{0,}[a-z]{0,4})|(?i)([a-z]{1,10}[0-9]{0,})|(?i)([a-z]{1,10})', 3)[0]
			If _ArraySearch($table1, $table[$i][0], 0, 0, 0, 0, 1, 1) = -1 Then
				Dim $temp = ['', $table[$i][0], '', $table[$i][3]]
				$temp[0] = _Excel_ColumnToLetter(4 * (UBound($table1) + 1) + 1)
				$temp[2] = StringRegExp($table[$i][1], '(?i)([0-9]{0,}[a-z]{1,4}[0-9]{0,}[a-z]{0,4})|(?i)([a-z]{1,10}[0-9]{0,})|(?i)([a-z]{1,10})', 3)[0] & '|'
				$temp1 = _ArrayAdd2D($table1, $temp)
				_Excel_RangeWrite($oBook, 3, UBound($table1), 'A1')
				Dim $temp = $header
				$temp[0][0] = $table[$i][0]
				$temp[0][1] = $table[$i][1] & '|'
				$temp[0][2] = $table[$i][3]
				_Excel_RangeWrite($oBook, 3, $temp, $table1[$temp1][0] & '1')
				;MsgBox(0, '', $temp1 & ' ' & @error)
				;_ArrayDisplay($table1)
				;_ArrayDisplay($table)
			ElseIf _ArraySearch($table1, $table[$i][0], 0, 0, 0, 0, 1, 1) <> -1 Then
				$temp1 = _ArraySearch($table1, $table[$i][0], 0, 0, 0, 0, 1, 1)
				If $table1[$temp1][3] <> $table[$i][3] Then
					$table1[$temp1][3] = $table[$i][3]
					_Excel_RangeWrite($oBook, 3, $table1[$temp1][1], $table1[$temp1][0] & '1')
					_Excel_RangeWrite($oBook, 3, $table[$i][3], _Excel_ColumnToLetter(_Excel_ColumnToNumber($table1[$temp1][0]) + 2) & '1')
				EndIf
				If $table1[$temp1][2] <> ($temp & '|') Then
					$table1[$temp1][2] = $temp
					_Excel_RangeWrite($oBook, 3, $table1[$temp1][1], $table1[$temp1][0] & '1')
					_Excel_RangeWrite($oBook, 3, $table[$i][2] & '|', _Excel_ColumnToLetter(_Excel_ColumnToNumber($table1[$temp1][0]) + 1) & '1')
				EndIf
			EndIf
		Next
		$temp = 0
	Else
		_Excel_RangeWrite($oBook, 3, UBound($table1), 'A1')
		For $i = 1 To $table[0][0]
			ConsoleWrite($i)
			Dim $temp = ['', $table[$i][0], '', $table[$i][3]]
			$temp1 = _Excel_ColumnToLetter(4 * (_Excel_RangeRead($oBook, 3, 'A1') + 1) + 1)
			$temp[0] = $temp1
			$temp1 = StringRegExp($table[$i][1], '(?i)([0-9]{0,}[a-z]{1,4}[0-9]{0,}[a-z]{0,4})|(?i)([a-z]{1,10}[0-9]{0,})|(?i)([a-z]{1,10})', 3)
			$temp[2] = $temp1[0] & '|'
			_ArrayAdd2D($table1, $temp)
			_Excel_RangeWrite($oBook, 3, UBound($table1), 'A1')
			$temp1 = $header
			$temp1[0][0] = $table[$i][0]
			$temp1[0][1] = $table[$i][1] & '|'
			$temp1[0][2] = $table[$i][3]
			_Excel_RangeWrite($oBook, 3, $temp1, $table1[$i - 1][0] & '1')
			;MsgBox(0,'',@error)
		Next
	EndIf
	;_ArrayDisplay($table1)
	_Excel_RangeWrite($oBook, 3, $table1, 'A3:D' & (UBound($table1) + 2))
	;If @error Then MsgBox(0, '', @error)
	_Excel_BookSave($oBook)
	_Excel_BookClose($oBook)
	_Excel_Close($oEx)


EndFunc   ;==>IndexExcel

