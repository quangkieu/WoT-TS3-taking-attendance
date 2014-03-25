#Include "JSMN.au3"
#include <array.au3>
#include <Inet.au3>
#include "_Excel_rewrite.au3"
#include <_Array2D.au3>
#include <IE.au3>
#include <File.au3>
#include <_HTML_Ex.au3>
#AutoIt3Wrapper_run_debug_mode=Y

#include <MsgBoxConstants.au3>

; Example getting an Object using it's class name
; ; Excel must be activated for this example to be successful
$oMyError = ObjEvent("AutoIt.Error","MyErrFunc")
TCPStartup()
MsgBox(0,'',TCPNameToIP('127.0.0.1'))
TCPShutdown()
tt()
;Row|Col 0|Col 1|Col 2|Col 3|Col 4
#cs Local $a[0][5]
Dim $temp=[19,'cid','clid','cdid','last']
_ArrayAdd2D($a,$temp)
Dim $temp=['bestkiller2',5,2,224,1394982252]
_ArrayAdd2D($a,$temp)
Dim $temp=['Wibs',5,3,62,1394699810]
Dim $temp=['Deathronicus',6,6,67,1394982827]
Dim $temp=['bigdognj1 t57', e100,4,8,86,1394973815]
Dim $temp=['Hardcore125 E-100-IS7-4-typeE',1,11,57,1394993157]
Dim $temp=['Ravonos',4,12,138,1394993295]
Dim $temp=['killwild',5,13,35,1394983746]
Dim $temp=['bazooka007',7,18,200,1394988294]
Dim $temp=['quangkieu',1,21,183,1394996441]
Dim $temp=['MKCSGlen',4,23,257,1394994866]
Dim [11],fasterwizard,5,27,245,1394981758
[12],Roguehitman,7,28,249,1394989711
[13],MightyDestroyer,6,34,235,1394827323
[14],OlBones,6,36,29,1394991713
[15],serveradmin from 66.31.54.140:65105,1,43,1,1394997613
[16],Arclight,5,44,16,1394995605
[17],abrunner25,1,46,102,1394995647
[18],Rougeace4,30,58,212,1394995982
{
  "data",{
    "battles_count",2
    "last_update","1394845659"
    "status","ok"
    "battles",[
      {
        "type","landing"
        "type_img_style","background-image:url(http://wotapi.ru/images/battle_types.png);display:inline-block;width:16px;height:16px;background-position:-48px;"
        "time","21:02"
        "primetime","20:00"
        "region",{
          "name","North America"
          "id","reg_08"
          "link","http://worldoftanks.com/clanwars/maps/globalmap/?province=US_129"
        }
        "province",{
          "name","Augusta, Maine"
          "id","US_129"
          "link","http://worldoftanks.com/clanwars/maps/globalmap/?province=US_129"
        }
        "maps","Fisherman's Bay"
        "server","US301"
      }
      {
        "type","landing"
        "type_img_style","background-image:url(http://wotapi.ru/images/battle_types.png);display:inline-block;width:16px;height:16px;background-position:-48px;"
        "time","22:13"
        "primetime","21:00"
        "region",{
          "name","North America"
          "id","reg_08"
          "link","http://worldoftanks.com/clanwars/maps/globalmap/?province=US_56"
        }
        "province",{
          "name","Denver, Colorado"
          "id","US_56"
          "link","http://worldoftanks.com/clanwars/maps/globalmap/?province=US_56"
        }
        "maps","Lakeville"
        "server","US301"
      }
    ]
  }
}
#ce [19],Boombaby,5,72,79,1394851260
#include <IE.au3>
#include <MsgBoxConstants.au3>
Func MyErrFunc()
  Msgbox(0,"AutoItCOM Test","We intercepted a COM Error !"    & @CRLF  & @CRLF & _
             "err.description is: "    & @TAB & $oMyError.description    & @CRLF & _
             "err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
             "err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @CRLF & _
             "err.helpcontext is: "    & @TAB & $oMyError.helpcontext _
            )
    Local $err = $oMyError.number
    If $err = 0 Then $err = -1
    SetError($err)
Endfunc
#include <MsgBoxConstants.au3>

OnAutoItExitRegister("MyTestFunc")
OnAutoItExitRegister("MyTestFunc2")

Sleep(1000)

Func MyTestFunc()
    MsgBox($MB_SYSTEMMODAL, "Exit Results 1", 'Exit Message from MyTestFunc()')
EndFunc   ;==>MyTestFunc

Func MyTestFunc2()
    MsgBox($MB_SYSTEMMODAL, "Exit Results 2", 'Exit Message from MyTestFunc()')
EndFunc   ;==>MyTestFunc2
Func tt()
	Local $temp,$temp1,$temp2,$table

	;$temp=_Date_Time_GetTimeZoneInformation()
	#cs For $i=0 To 2
		$temp1=_HTML_GetSource('http://api.worldoftanks.com/wot/clan/battles/?application_id=8b379c513ce91c4e9ddfe538a12b0e67&clan_id=1000002504')
		;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
		Sleep(500)
		$temp2=_HTML_GetSource('http://api.worldoftanks.com/wot/clan/battles/?application_id=8b379c513ce91c4e9ddfe538a12b0e67&clan_id=1000002504')
		;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
		If $temp1<>$temp2 Then ExitLoop
	#ce Next
	$temp2='{"status":"ok","count":1,"data":{"1000007267":[{"provinces":["CA_84"],"started":false,"private":null,"time":1395451803,"arenas":[{"name_i18n":"Swamp","name":"22_slough"}],"type":"for_province"},{"provinces":["CA_81"],"started":false,"private":null,"time":1395451803,"arenas":[{"name_i18n":"Tundra","name":"63_tundra"}],"type":"for_province"},{"provinces":["CA_81","CA_84"],"started":false,"private":null,"time":1395447603,"arenas":[{"name_i18n":"Swamp","name":"22_slough"},{"name_i18n":"Tundra","name":"63_tundra"}],"type":"meeting_engagement"}]}}'
	If @error Or StringInStr($temp2,'"error"')<>0 Then Return SetError(1,0,0);no connection or wrong ClanID
	$temp1=StringRegExp($temp2,'(?!"provinces":\[)(,*"\w+_\d+")+(?=],"started")',3)
	$temp=UBound($temp1)
	If $temp=0 Then Return SetError(0,0,0);no CW battle
	$table=$temp2
	#cs For $i=0 To 2
		$temp1=_HTML_GetSource('http://cw.wotapi.ru/na/1000002504?type=json')
		;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
		Sleep(500)
		$temp2=_HTML_GetSource('http://cw.wotapi.ru/na/1000002504?type=json')
		;http://cw.wotapi.ru/na/1000002504?utc='&(0-(($temp[0]=0 Or $temp[0]=1)?($temp[1]):(($temp[0]=2)?($temp[1]+$temp[7]):('0')) ))&'&type=json
		If $temp1<>$temp2 Then ExitLoop
	#ce Next
	$temp2='{"data":{"battles_count":3,"last_update":"1395441011","status":"ok","battles":[{"type":"meeting_engagement","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: 0px;","time":"00:20","primetime":"00:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_81"},"province":{"name":"Western Kativik","id":"CA_81","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_81"},"maps":"Tundra","server":"US301"},{"type":"for_province","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: -32px;","time":"01:30","primetime":"00:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_81"},"province":{"name":"Western Kativik","id":"CA_81","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_81"},"maps":"Tundra","server":"US301"},{"type":"for_province","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: -32px;","time":"01:30","primetime":"00:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_84"},"province":{"name":"Eastern Kativik","id":"CA_84","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_84"},"maps":"Swamp","server":"US301"}]}}'
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
	;	/		/		\			\		/	/		/		\
	;code | type | region | province | server |
	For $i=1 To $table[0][0]
		If $table[$i][6]<>'meeting_engagement' Then
			$temp1=_ArrayFindAll($temp,$table[$i][5],1)
			For $j=0 To UBound($temp1)-1
				If $temp[$temp1[$j]][1]<>'meeting_engagement' Then
					$table[$i][2]=$temp[$temp1[$j]][2]
					$table[$i][3]=$temp[$temp1[$j]][3]
					$table[$i][7]=$temp[$temp1[$j]][4]
					ExitLoop
				EndIf
			Next
		EndIf
	Next
	_ArrayDisplay($table)
EndFunc


Func Test1()
	Local $Json1 = FileRead(@ScriptDir & "\test.json")
	Local $Data1 = Jsmn_Decode($Json1)
	Local $Json2 = Jsmn_Encode($Data1)

	Local $Data2 = Jsmn_Decode($Json2)
	Local $Json3 = Jsmn_Encode($Data2)

	ConsoleWrite("Test1 Result: " & $Json3 & @CRLF)
	Return ($Json2 = $Json3)
EndFunc

Func Test2()
	Local $Json1 = FileRead(@ScriptDir & "\test.json")
	Local $Data1 = Jsmn_Decode($Json1)
	Local $Json2 = Jsmn_Encode($Data1)

	Local $2DArray = Jsmn_ObjTo2DArray($Data1)
	_ArrayDisplay($2DArray)
	Local $Data2 = Jsmn_ObjFrom2DArray($2DArray)

	Local $Json3 = Jsmn_Encode($Data2)

	ConsoleWrite("Test2 Result: " & $Json3 & @CRLF)
	Return ($Json2 = $Json3)
EndFunc

Func Test3()
	Local $Json1 = '{"data":{"battles_count":3,"last_update":"1395406823","status":"ok","battles":[{"type":"for_province","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: -32px;","time":"21:30","primetime":"20:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_84"},"province":{"name":"Eastern Kativik","id":"CA_84","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_84"},"maps":"Swamp","server":"US301"},{"type":"for_province","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: -32px;","time":"21:30","primetime":"20:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_81"},"province":{"name":"Western Kativik","id":"CA_81","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_81"},"maps":"Tundra","server":"US301"},{"type":"meeting_engagement","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: 0px;","time":"20:20","primetime":"20:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_81"},"province":{"name":"Western Kativik","id":"CA_81","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=CA_81"},"maps":"Tundra","server":"US301"}]}}'
	;InetRead('http://cw.wotapi.ru/na/1000002504?utc=-4&type=json',1)
	;'http://api.worldoftanks.com/wot/clan/info/?application_id=demo&fields=members_count,members.account_name&clan_id=1000002504',1);{"data":{"battles_count":2,"last_update":"1391562612","status":"ok","battles":[{"type":"for_province","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: -32px;","time":"21:30","primetime":"19:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=US_22"},"province":{"name":"Sierra Nevada Mountains","id":"US_22","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=US_22"},"maps":"Karelia","server":"US301"},{"type":"for_province","type_img_style":"background-image:url(http:\/\/wotapi.ru\/images\/battle_types.png);display:inline-block;width:16px;height:16px;background-position: -32px;","time":"21:30","primetime":"19:00","region":{"name":"North America","id":"reg_08","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=US_33"},"province":{"name":"Winnemucca, Nevada","id":"US_33","link":"http:\/\/worldoftanks.com\/clanwars\/maps\/globalmap\/?province=US_33"},"maps":"Karelia","server":"US301"}]}}'
	$Json1=BinaryToString($Json1)
	$Json1=StringReplace($Json1,'\n','')
	MsgBox(0,'',$Json1)
	Local $Data1 = Jsmn_Decode($Json1)
	Local $2DArray = Jsmn_ObjTo2DArray($Data1)
	#cs _ArrayDisplay($2DArray)
	$temp=$2DArray[1][1]
	_ArrayDisplay($temp)
	if $temp[1][1]<>0 Then
		$temp=$temp[4][1]
		_ArrayDisplay($temp)
		$temp=$temp[0]
		_ArrayDisplay($temp)
		$temp=$temp[6][1]
		_ArrayDisplay($temp)
	EndIf
	;$2DArray = Jsmn_ObjTo2DArray($temp)
	$temp=$temp[1][1]
	$temp=$temp[2][1]
	$temp=$temp[1][1]
	;$temp=$temp[0]
	;$temp=$temp[0][1]
	MsgBox(0,'',$temp)
	MsgBox(0,"",UBound($temp))
	;$temp=$temp[0]
	;$temp=$temp[6][1]
	#ce _ArrayDisplay($temp)

	Local $Json2 = Jsmn_Encode($Data1, $JSMN_UNQUOTED_STRING)
	Local $Data2 = Jsmn_Decode($Json2)

	Local $Json3 = Jsmn_Encode($Data2, $JSMN_PRETTY_PRINT, "  ", "\n", "\n", ",")
	Local $Data3 = Jsmn_Decode($Json3)

	Local $Json4 = Jsmn_Encode($Data3, $JSMN_STRICT_PRINT)

	ConsoleWrite("Test3 Unquoted Result: " & $Json2 & @CRLF)
	ConsoleWrite("Test3 Pretty Result: " & $Json3 & @CRLF)
	Return ($Json1 = $Json4)
EndFunc


; -----------------------------------------------------------------------------------------
; Proof of concept: Implement nested DLLStructs using Arrays and normal DLLStructs
; Author: ProgAndy
; Nested Structs habe to be defined as Global Constants as TYPENAME or tagTYPENAME
; Access members via qualifier string: first_member.submember
; Array access with:                   member[2].submember[5]
;
; Access members via index:            #2.#1.#6
; Fetching Pointer for array element is possible: _DllStructGetPtr($struct, "array[3]")
; Member index and array index are 1-based.
; Supports STRUCT;...;ENDSCTRUCT of AutoIt-Beta + added naming structs and multiple nesting levels.
; The indexing is different:
;    char a;STRUCT name;int b;ENDSTRUCT;int x
;    Access-Qualifiers: a        = #1
;                       name.b   = #2.#1
;                       x        = #3
;
; _DllStructCreate accepts an optional 3rd parameter: a Scripting.Dictionary.
; This dictionary can contain your substructure definitions. (Key = TypeName, value = definition)
; -----------------------------------------------------------------------------------------

; ### Nested DllStructs #######################################################################################
#CS
	#region Nested DLLStructs


	Func _DLLStructCreate($sStruct, $pPtr = 0, $oSubStructs=0)
	; Author: ProgAndy
	Local $sStructTag, $fClean=False
	If IsObj($oSubStructs) Then $fClean = True
	__DllStructExtractSubStructs($sStruct, $oSubStructs)
	Local $aDllStruct = __DLLStructParse($sStruct, $sStructTag, $oSubStructs)
	Local $err = @error
	If $fClean Then
	Local $a
	For $a In $oSubStructs.Keys
	If StringLeft($a, 27) = "A__XX__XX__UNNAMED_STRUCT__" Then $oSubStructs.Remove($a)
	Next
	EndIf
	If $err Then Return SetError(5, $err, 0)
	If @NumParams = 1 Or (@NumParams = 3 And $pPtr = 0) Then
	$aDllStruct[0][0] = DllStructCreate($sStructTag)
	Else
	$aDllStruct[0][0] = DllStructCreate($sStructTag, $pPtr)
	EndIf
	If @error Then Return SetError(@error, 0, 0)
	Return $aDllStruct
	EndFunc   ;==>_DLLStructCreate

	Func _DllStructGetData(ByRef $aDllStruct, $sQualifier)
	; Author: ProgAndy
	Local $sName = __DllStructFindIndex($aDllStruct, $sQualifier)
	If @error Then Return SetError(@error, 0, 0)
	Local $iIndex = @extended
	If IsArray($aDllStruct[$sName][0]) Then
	Local $aRet = $aDllStruct[$sName][0]
	$aRet[0][0] = DllStructCreate($aDllStruct[$sName][1], DllStructGetPtr($aDllStruct[0][0], $sName))
	If $iIndex > 1 Then $aRet[0][0] = DllStructCreate($aDllStruct[$sName][1], DllStructGetPtr($aDllStruct[0][0], $sName) + DllStructGetSize($aRet[0][0]) * ($iIndex - 1))
	If $sQualifier = '' Then
	Return $aRet
	Else
	Return _DllStructGetData($aRet, $sQualifier)
	EndIf
	Else
	If $iIndex = 0 Then Return DllStructGetData($aDllStruct[0][0], $sName)
	Return DllStructGetData($aDllStruct[0][0], $sName, $iIndex)
	EndIf
	EndFunc   ;==>_DllStructGetData

	Func _DllStructGetPtr(ByRef $aDllStruct, $sQualifier="")
	; Author: ProgAndy
	If @NumParams = 1 Then Return DllStructGetPtr($aDllStruct[0][0])
	Local $sName = __DllStructFindIndex($aDllStruct, $sQualifier)
	If @error Then Return SetError(@error, 0, 0)
	Local $iIndex = @extended
	If IsArray($aDllStruct[$sName][0]) Then
	If $iIndex <= 1 And $sQualifier = "" Then
	Return DllStructGetPtr($aDllStruct[0][0], $sName)
	Else
	Local $tTemp = DllStructCreate($aDllStruct[$sName][1], 1)
	If $iIndex < 1 Then $iIndex = 1
	Local $pPtr = DllStructGetPtr($aDllStruct[0][0], $sName) + ($iIndex - 1) * DllStructGetSize($tTemp)
	If $sQualifier = "" Then Return $pPtr
	$tTemp = DllStructCreate($aDllStruct[$sName][1], $pPtr)
	Local $aRet = $aDllStruct[$sName][0]
	$aRet[0][0] = $tTemp
	Return _DllStructGetPtr($aRet, $sQualifier)
	EndIf
	Else
	If $iIndex > 1 Then Return DllStructGetPtr($aDllStruct[0][0], $sName) + DllStructGetSize(DllStructCreate($aDllStruct[$sName][1], 1)) * ($iIndex-1)
	Return DllStructGetPtr($aDllStruct[0][0], $sName)
	EndIf
	EndFunc   ;==>_DllStructGetPtr

	Func _DllStructGetSize(ByRef $aDllStruct)
	; Author: ProgAndy
	Return DllStructGetSize($aDllStruct[0][0])
	EndFunc

	Func _DllStructSetData(ByRef $aDllStruct, $sQualifier, $vValue)
	; Author: ProgAndy
	Local $sName = __DllStructFindIndex($aDllStruct, $sQualifier)
	If @error Then Return SetError(@error, 0, 0)
	Local $iIndex = @extended
	If IsArray($aDllStruct[$sName][0]) Then
	Local $t = DllStructCreate($aDllStruct[$sName][1], DllStructGetPtr($aDllStruct[0][0], $sName))
	If $iIndex > 1 Then $t = DllStructCreate($aDllStruct[$sName][1], DllStructGetPtr($aDllStruct[0][0], $sName) + DllStructGetSize($t) * ($iIndex - 1))
	If $sQualifier = '' Then
	DllStructSetData(DllStructCreate("byte[" & DllStructGetSize($t) & "]", DllStructGetPtr($t)), 1, $vValue)
	Else
	Local $aRet = $aDllStruct[$sName][0]
	$aRet[0][0] = $t
	Return _DllStructSetData($aRet, $sQualifier, $vValue)
	EndIf
	Else
	If $iIndex = 0 Then Return DllStructSetData($aDllStruct[0][0], $sName, $vValue)
	Return DllStructSetData($aDllStruct[0][0], $sName, $vValue, $iIndex)
	EndIf
	EndFunc   ;==>_DllStructSetData

	Func __DllStructFindIndex(ByRef $aDllStruct, ByRef $sQualifier)
	; Author: ProgAndy
	Local $sName, $iIndex = 0
	If IsInt($sQualifier) Then
	$sName = $sQualifier
	$sQualifier = ""
	Else
	$sName = __DllStructSplitNameIndexNext($sQualifier)
	$iIndex = @extended
	EndIf
	If IsInt($sName) Then
	Local $iBound = UBound($aDllStruct)-1
	If $sName < 1 Or $sName > $iBound Then Return SetError(1, 0, 0)
	; Find index, skipping alignment bytes
	For $i = 1 To $sName
	If IsInt($aDllStruct[$i][0]) Then $sName += 1
	If $sName > $iBound Then Return SetError(1,0,0)
	Next
	Else
	$sName = __DllStructGetNameIndex($aDllStruct[0][1], $sName)
	If @error Then Return SetError(1, 0, 0)
	EndIf
	SetExtended($iIndex)
	Return $sName
	EndFunc

	Func __DllStructSplitNameIndexNext(ByRef $sName)
	; Author: ProgAndy
	$sName = StringStripWS($sName, 8)
	Local $iPos = StringInStr($sName, ".", 1, 1), $aName
	If $iPos Then
	$aName = StringSplit(StringLeft($sName, $iPos - 1), '[]')
	$sName = StringTrimLeft($sName, $iPos)
	Else
	$aName = StringSplit(StringStripWS($sName, 8), '[]')
	$sName = ""
	EndIf
	If StringLeft($aName[1], 1) = "#" Then $aName[1] = Number(StringTrimLeft($aName[1], 1))
	If $aName[0] = 1 Then Return SetExtended(0, $aName[1])
	Return SetExtended($aName[2], $aName[1])
	EndFunc   ;==>__DllStructSplitNameIndexNext

	Func __DllStructGetNameIndex(ByRef $sData, $sName)
	; Author: ProgAndy
	If $sName = "" Then Return SetError(2, 0, 0)
	Local $iPos = StringInStr($sData, ";" & StringLower($sName) & ";", 2, 1)
	If Not $iPos Then Return SetError(1, 0, 0)
	StringReplace(StringLeft($sData, $iPos), ";", ";", 0, 1)
	Return @extended
	EndFunc   ;==>__DllStructGetNameIndex

	Func __DLLStructParse($sStruct, ByRef $sStructTag, $oSubStructs=0, $iAlignment=0)
	; Author: ProgAndy
	Local Static $sTypes = "|align|BYTE|BOOLEAN|CHAR|WCHAR|short|USHORT|WORD|int|long|BOOL|UINT|ULONG|DWORD|INT64|UINT64|ptr|HWND|HANDLE|float|double|INT_PTR|LONG_PTR|LRESULT|LPARAM|UINT_PTR|ULONG_PTR|DWORD_PTR|WPARAM|"
	Local $a = StringSplit($sStruct, ";")
	Local $sName, $r = "", $sItems = ";", $aDllStruct[UBound($a)*2][2], $iMemberIndx = 0
	$sStructTag = ""
	If StringStripWS($a[$a[0]], 8) = "" Then $a[0] -= 1
	If $a[0] = 0 Then Return SetError(2, 0, 0)
	Local $iOffset = 0, $iFirstAlign = $iAlignment;, $iAlignment = 0
	For $i = 1 To $a[0]
	$sName = StringRegExp($a[$i], "^\h*(\w+)\h*(\w*)\h*(?:\[(\d*)\])?\h*$", 1)
	If @error Then Return SetError(2, 0, 0)
	ReDim $sName[3]
	If $sName[0] = "" Then
	; nothing
	ElseIf StringInStr($sTypes, '|' & $sName[0] & '|', 2, 1) Then
	$r &= $a[$i] & ";"
	If $sName[0] <> "align" Then
	$iMemberIndx += 1
	$aDllStruct[$iMemberIndx][1] = $sName[0]
	$sItems &= $sName[1] & ";"
	If $sName[2] < 1 Then $sName[2] = 1
	$iOffset += __DllStructAlignBytesType($iOffset, $sName[0], $iAlignment) & __DllStructGetTypeSize($sName[0])*$sName[2]
	Else
	$iAlignment = Number($sName[1])
	If $i = 1 Then $iFirstAlign = $iAlignment
	EndIf
	Else
	$iMemberIndx += 1
	Local $sSubTag, $sSubByteMember, $t
	$t = __DllStructSubCreate($sName[0], $sName[1], $sName[2], $sSubByteMember, $sSubTag, $oSubStructs, $iAlignment)
	If @error Then Return SetError(2, 0 , 0)

	Local $iAlign = __DllStructAlignBytesSubStruct($iOffset, $t, $iAlignment)
	If $iAlign Then
	$r &= "byte[" & $iAlign & "];"
	$aDllStruct[$iMemberIndx][0] = 1 ; mark alignment bytes
	$iMemberIndx += 1
	$iOffset += $iAlign
	$sItems &= ";"
	EndIf

	$r &= $sSubByteMember
	$aDllStruct[$iMemberIndx][1] = $sSubTag
	$aDllStruct[$iMemberIndx][0] = $t
	$sItems &= $sName[1] & ";"

	$iOffset += DllStructGetSize(DllStructCreate($sSubTag, 1)) * $sName[2]
	EndIf
	Next

	; $iOffset is length of struct, make sure size is correct for array of structs
	;~  Local $iAlign = __DllStructAlignBytesSubStruct($iOffset, $aDllStruct, $iFirstAlign)
	Local $iAlign = __DllStructAlignBytesSubStruct($iOffset, $aDllStruct, $iAlignment)
	If $iAlign Then
	$r &= "byte[" & $iAlign & "];"
	$iOffset += $iAlign
	EndIf

	$aDllStruct[0][1] = $sItems
	ReDim $aDllStruct[$iMemberIndx + 1][2]
	$sStructTag = $r
	Return $aDllStruct
	EndFunc   ;==>__DLLStructParse

	Func __DllStructSubCreate(ByRef $sType, ByRef $sName, ByRef $iCount, ByRef $sByteMember, ByRef $sStructTag, $oSubStructs = 0, $iAlignment = 0)
	; Author: ProgAndy
	$sByteMember = ""
	Local $aDllStruct
	Select
	Case IsObj($oSubStructs) And $oSubStructs.Exists($sType)
	$aDllStruct = __DLLStructParse($oSubStructs($sType), $sStructTag, $oSubStructs, $iAlignment)
	If @error Then ContinueCase
	Case True
	$aDllStruct = __DLLStructParse(Eval("tag" & $sType), $sStructTag, $oSubStructs, $iAlignment)
	If @error Then ContinueCase
	Case Else
	$aDllStruct = __DLLStructParse(Eval($sType), $sStructTag, $oSubStructs, $iAlignment)
	If @error Then Return SetError(1, 0, 0)
	$sType = "tag" & $sType
	EndSelect
	Local $iSize = DllStructGetSize(DllStructCreate($sStructTag, 1))
	If @error Or $iSize = 0 Then Return SetError(1, 0, 0)
	If Number($iCount) < 1 Then $iCount = 1
	$sByteMember = "byte " & $sName & "[" & $iCount * $iSize & "];"
	Return $aDllStruct
	EndFunc   ;==>__DllStructSubCreate

	Func __DllStructAlignBytesSubStruct($iOffset, ByRef $aDllStruct, $iAlignment = Default)
	Return __DllStructAlignBytesType($iOffset, __DllStructFindLargestMember($aDllStruct), $iAlignment)
	EndFunc

	Func __DllStructFindLargestMember(ByRef $aDllStruct)
	Local $iLargest = 0, $iNew
	; find largest member alignment
	For $i = 1 To UBound($aDllStruct)-1
	If IsArray($aDllStruct[$i][0]) Then
	$iNew = __DllStructFindLargestMember($aDllStruct[$i][0])
	Else
	$iNew = __DllStructGetTypeSize($aDllStruct[$i][1])
	EndIf
	If $iNew > $iLargest Then $iLargest = $iNew
	Next
	Return $iLargest
	EndFunc

	Func __DllStructAlignBytesType($iOffset, $sType, $iAlignment=Default)
	If IsInt($sType) Then
	Local $iLen = $sType
	Else
	Local $iLen = __DllStructGetTypeSize($sType)
	EndIf
	If IsKeyword($iAlignment) Or $iAlignment < 1 Then $iAlignment = 8
	If $iLen = 1 Or $iAlignment = 1 Then
	Return 0
	ElseIf $iAlignment > $iLen Then
	$iOffset = $iLen - Mod($iOffset, $iLen)
	If $iOffset = $iLen Then Return 0
	Else
	$iOffset = $iAlignment - Mod($iOffset, $iAlignment)
	If $iOffset = $iAlignment Then Return 0
	EndIf
	Return $iOffset
	EndFunc

	Func __DllStructGetTypeSize($sType)
	Local Static $iPtrSize = DllStructGetSize(DllStructCreate("ptr", 1))
	Switch $sType
	Case "BYTE", "BOOLEAN", "CHAR"
	Return 1
	Case "WCHAR", "short", "USHORT", "WORD"
	Return 2
	Case "int", "long", "BOOL", "UINT", "ULONG", "DWORD", "float"
	Return 4
	Case "INT64", "UINT64", "double"
	Return 8
	Case "ptr", "HWND", "HANDLE", "INT_PTR", "LONG_PTR", "LRESULT", "LPARAM", "UINT_PTR", "ULONG_PTR", "DWORD_PTR", "WPARAM"
	Return $iPtrSize
	EndSwitch
	Return SetError(1,0,0)
	EndFunc

	Func __DllStructExtractSubStructs(ByRef $sStruct, ByRef $oSubStructs, $iStart = 0)
	; ProgAndy
	Local Static $counter = 0
	$sStruct = StringReplace(StringReplace(StringRegExpReplace($sStruct, "\s+", " "), "; ", ";", 0, 1), " ;", ";", 0, 1)
	$sStruct = StringReplace(";"&$sStruct&";", ";;", ";", 0, 1)
	If $iStart < 1 Then $iStart = StringInStr($sStruct, ";STRUCT", 2, 1)
	Local $iEnd, $iStart2, $iNameEnd, $sName
	While $iStart
	If Not IsObj($oSubStructs) Then $oSubStructs = ObjCreate("Scripting.Dictionary")
	$iNameEnd = StringInStr($sStruct, ";", 2, 1, $iStart+7)
	If Not $iNameEnd Then Return SetError(1, 0, 0)
	$sName = StringMid($sStruct, $iStart+7, $iNameEnd-$iStart-7)
	If $sName <> "" And Not StringRegExp($sName, "^\h\w+$") Then Return SetError(2,0,0)
	$iEnd = StringInStr($sStruct, ";ENDSTRUCT;", 2, 1, $iNameEnd)
	$iStart2 = StringInStr($sStruct, ";STRUCT", 2, 1, $iNameEnd)
	If $iStart2 > 0 And $iEnd > $iStart2 Then
	__DllStructExtractSubStructs($sStruct, $oSubStructs, $iStart2)
	Else
	Local $sTemp = StringLeft($sStruct,  $iStart)
	Local $sSub = StringMid($sStruct, $iNameEnd+1, $iEnd - $iNameEnd - 1)
	Local $sTemp2 = StringTrimLeft($sStruct, $iEnd+9)
	__DllStructExtractSubStructs($sSub, $oSubStructs)
	Local $sAlign = ""
	Local $iAlign = StringInStr(";" & $sSub, ";align ", 2, -1)
	If $iAlign Then
	Local $iCut = StringInStr($sSub & ";", ";", 1, 1, $iAlign)
	$sAlign = ";" & StringMid($sSub, $iAlign, $iCut-$iAlign)
	EndIf
	$counter += 1
	$oSubStructs("A__XX__XX__UNNAMED_STRUCT__" & $counter) = $sSub
	$sStruct = $sTemp & "A__XX__XX__UNNAMED_STRUCT__" & $counter & $sName & $sAlign & $sTemp2
	EndIf
	$iStart = StringInStr($sStruct, ";STRUCT", $iStart)
	WEnd
	If @NumParams = 2 Then $sStruct = StringMid($sStruct, 2, StringLen($sStruct)-2)
	EndFunc

	#endregion Nested DLLStructs
#CE

; ### End Nested DllStructs ##################################################################################
;Local $this=[[4,'','','',''],['Dungwad',2,2,2,1393940501],['Hardcore125',5,9,3,1393944308],['Boombab',11,29,4,1393901643],['Wibs E5',5,114,5,1393897420]]
;dim $PassUser=[[4,'','',''],['Dungwad',10,2,2],['Hardcore125',5,3,9],['Boombab',29,4,29],['Wibs E5',5,6,114]]
;LogUsers($this)

#cs Global Const $SCLIENT     = "wchar name[30]; int clid; int cdid;"
	Global Const $SCHANNEL    = "wchar name[30]; int cid; int on; SCLIENT clients[50]; SSUPCHANNEL channels[30];"
	GLobal Const $SSUPCHANNEL = "wchar name[30]; int cid; int parent; int on; SCLIENT clients[50];"
	$ServerStr=""
	$Server=0

	Global Const $SUBTESTTYPE = "wchar s[12]"
	Global Const $tagTEST = "int test;int blub; SUBTESTTYPE hi[3]"

	$x = _DLLStructCreate("byte[2]; TEST x")

	_DllStructSetData($x, "x.blub", 23)
	MsgBox(0, 'x.blub = #2.#2', _DllStructGetData($x, "#2.#2"))
	MsgBox(0, 'Ptr to #1 and #1[2]', _DllStructGetPtr($x, "#1") & @CRLF &  _DllStructGetPtr($x, "#1[2]"))
	_DllStructSetData($x, "x.hi[1].s", "hello")
	_DllStructSetData($x, "x.hi[3].s", "bb")
	_DllStructSetData($x, "x.hi[2].s", ":)")
	MsgBox(0, 'x.hi.s = x.hi[1].s', _DllStructGetData($x, "x.hi.s"))
	MsgBox(0, 'x.hi[2].s', _DllStructGetData($x, "x.hi[2].s"))
#ce MsgBox(0, 'x.hi[3].s', _DllStructGetData($x, "x.hi[3].s"))