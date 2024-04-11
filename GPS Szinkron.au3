#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=u-Blox.ico
#AutoIt3Wrapper_Outfile=..\GPS Szinkron.exe
#AutoIt3Wrapper_Outfile_x64=..\GPS Szinkron x64.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Res_Fileversion=1.01.04
#AutoIt3Wrapper_Res_Language=1038
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <WinAPI.au3>
#include <Date.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

FileInstall("Logo.jpg", "Logo.jpg")

Global $CommPort, $Port, $Buffer, $String, $Data, $MM, $DD, $YY, $H, $M, $S
Global $PCTime, $GPSTime

FindCOM("VID_1546&PID_01A7") ;~ u-Blox7 USB

_OpenPort()

Opt("GUIOnEventMode", 1)
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("GPS Idő Szinkronizáló - " & $CommPort, 420, 380, -1, -1)
GUISetOnEvent($GUI_EVENT_CLOSE, "Form1Close")
GUICtrlCreatePic("logo.jpg", 5, 5, 410, 160)

GUICtrlCreateGroup("GPS Idő", 5, 170, 410, 70)
$Label1 = GUICtrlCreateLabel("GPS IDŐ", 10, 190, 400, 45)
guictrlsetfont(-1, 22, 800)

GUICtrlCreateGroup("Rendszer Idő", 5, 250, 410, 70)
$Label2 = GUICtrlCreateLabel("PC IDŐ", 10, 270, 400, 45)
guictrlsetfont(-1, 22, 800)

$Button1 = GUICtrlCreateButton("Szinkron", 10, 330, 130, 40)
GUICtrlSetOnEvent(-1, "Button1Click")
GUICtrlSetFont(-1, 12, 800)
GUICtrlSetState(-1, $GUI_DISABLE)

$Button3 = GUICtrlCreateButton("Kilépés", 280, 330, 130, 40)
GUICtrlSetOnEvent(-1, "Button3Click")
GUICtrlSetFont(-1, 12, 800)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	_GetData()
	_ProcessString()
	GUICtrlSetData($Label1, $GPSTime)
	$PCTime = _Date_Time_GetSystemTime()
	GUICtrlSetData($Label2, _Date_Time_SystemTimeToDateTimeStr($PCTime))
	sleep(1000)
WEnd

Func Button1Click()
	Local $tNew = _Date_Time_EncodeSystemTime($MM, $DD, $YY, $H, $M, $S+1)
	$PCTime = _Date_Time_GetSystemTime()
	$Time1 = "2000/01/01 " & StringRight(_Date_Time_SystemTimeToDateTimeStr($PCTime), "8")
	$Time2 = "2000/01/01 " & $H & ":" & $M & ":" & $S
	$diff = _DateDiff("s", $Time2, $Time1)
	FileWrite("GPS Szinkron.log", @YEAR&"."&@MON&"."&@MDAY&" "&@HOUR&":"&@MIN&":"&@SEC&" - GPS Idő (UTC): "&$MM&"/"&$DD&"/"&$YY&" "&$H&":"&$M&":"&$S&" - PC Idő (UTC): "&_Date_Time_SystemTimeToDateTimeStr($PCTime)&" - Különbség: "&$diff&" sec"&@CRLF)
	_Date_Time_SetSystemTime($tNew)
EndFunc

Func Button3Click()
	Form1Close()
EndFunc

Func Form1Close()
	If @Compiled Then FileDelete("logo.jpg")
	Exit
EndFunc

Func _ProcessString()
	$Ft = StringLeft($DATA[2], 6)
	$Rt = StringSplit($Ft, "")
	$Fd = StringSplit($DATA[10], "")

	If $Fd[0] = "0" Then
		$GPSTime = "00/00/0000 00:00:00"
		GUICtrlSetState($button1, $GUI_DISABLE)
	Else
		GUICtrlSetState($button1, $GUI_ENABLE)
		$MM = $Fd[3] & $Fd[4]
		$DD = $Fd[1] & $Fd[2]
		$YY = "20" & $Fd[5] & $Fd[6]
		$H  = $Rt[1] & $Rt[2]
		$M  = $Rt[3] & $Rt[4]
		$S  = $Rt[5] & $Rt[6]
		$GPSTime = $MM & "/" & $DD & "/" & $YY & " " & $H & ":" & $M & ":" & $S
	EndIf
EndFunc

Func _OpenPort()
	$Port = _WinAPI_CreateFile($CommPort, 2, 2)
	If $Port = 0 then
		msgbox(16, "Hiba", "Nem találom a GPS modult")
		exit
	else
		$Buffer = DllStructCreate("byte[1]")
	EndIf
EndFunc

Func _GetData()
	While 1
		_GetString()
		if StringLeft($String,6) = "$GPRMC" then ; $GPGGA $GPGSV $GPGLL $GPGSA $GPGST
			$Data = stringsplit($String, ",")
			Return
		EndIf
	WEnd
EndFunc

Func _GetString()
	Local  $C = 0, $TempString = "", $sText
	While 1
		Local $nBytes=1
	    _WinAPI_ReadFile($Port, DllStructGetPtr($Buffer), 1, $nBytes)
	    $sText = BinaryToString(DllStructGetData($Buffer, 1))
		if $stext=chr(13) or $sText=chr(10) then
			$C=$C+1
			if $C=2 then
				$String = $TempString
				$TempString = ""
				$C=0
				Return
			endif
	    else
			$TempString= $TempString & $sText
	   EndIf
	WEnd
EndFunc

Func FindCOM($DEVID)
    Local $NumericPort, $Leftprens, $Rightprens, $TestDevice
    Local $oWMIService = ObjGet("winmgmts:\\localhost\root\CIMV2")
    If @error Then Return SetError(@error, 0, "")
    Local $oItems = $oWMIService.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(COM%)'", "WQL", 48)
    $CommPort = ""
    For $oItem In $oItems
        $TestDevice = StringInStr($oItem.DeviceID, $DEVID)
        If $TestDevice = 0 Then
			ContinueLoop
        EndIf
        $Leftprens = StringInStr($oItem.Name, "(")
        $Rightprens = StringInStr($oItem.Name, ")")
        $CommPort = StringMid($oItem.Name, $Leftprens + 1, $Rightprens - $Leftprens - 1)
        $NumericPort = Int(StringRight($CommPort, StringLen($CommPort) - 3))
		$CommPort = "COM" & StringLeft($NumericPort, 3)
	Next
    Return $CommPort
EndFunc