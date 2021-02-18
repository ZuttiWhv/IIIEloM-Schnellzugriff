#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListBox.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <Excel.au3>
#include <ExcelConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiComboBoxEx.au3>
#include <GUIConstantsEx.au3>
#include <GuiImageList.au3>
#include <Process.au3>
#include <GuiMenu.au3>
#include <WinAPI.au3>
#include <WinAPIFiles.au3>
#RequireAdmin

$configfilepath = "./Data/config.ini"
$logonip = IniRead($configfilepath, "netzwerk", "iphicom", "default")
$telnethicom = IniRead($configfilepath, "netzwerk", "telnethicom", "default")
$puttycommand = "putty.exe -telnet " & $telnethicom &" 23"

   $auswahladapter1 = IniRead($configfilepath, "netzwerk", "adapter1", "default")
   $auswahladapter2 = IniRead($configfilepath, "netzwerk", "adapter2", "default")
    $GUIadapterauswahl = GUICreate("Netzwerkauswahl", 250, 150)
    $adapterComboBox = GUICtrlCreateCombo($auswahladapter1, 30, 10, 185, 20)
	$testmodus = GUICtrlCreateCheckbox("Testmodus", 10, 90)
    $adapter_Send = GUICtrlCreateButton("Senden", 10, 120, 70, 25)
    $idButton_Close = GUICtrlCreateButton("Schließen", 170, 120, 70, 25)
    GUICtrlSetData($adapterComboBox, $auswahladapter2 & "|Manuell", $auswahladapter1)

    GUISetState(@SW_SHOW, $GUIadapterauswahl)

    Local $adapterread = ""
    While 1
	  Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idButton_Close
			   Exit

			Case $adapterComboBox
				  $adapterread = GUICtrlRead($adapterComboBox)
			   If $adapterread = "Manuell" Then
				  GUICtrlCreateLabel("Netzwerkkarte:",30,50)
				  $adaptermanuell = GUICtrlCreateInput("",110,46,105)
			   EndIf
			Case $adapter_Send
			   $modus = 0
			   Global $modus = GUICtrlRead($testmodus)
			   If $modus = $GUI_CHECKED Then
			   Global $aktivesfenster = "Unbenannt"
			   Else
			   Global $aktivesfenster = $telnethicom &" - PuTTY"
			   EndIf
				  $adapterread = GUICtrlRead($adapterComboBox)
			   If $adapterread = $auswahladapter1 Then
				  $logonadapter = $auswahladapter1
				  ExitLoop
			   ElseIf $adapterread = $auswahladapter2 Then
				  $logonadapter = $auswahladapter2
				   ExitLoop

			   ElseIf $adapterread = "Manuell" Then
						$adapterinput = GUICtrlRead($adaptermanuell)

					 If $adapterinput = "" Then
						MsgBox($MB_OK, "Fehler", "Manuellen Netzwerknamen eingeben!")
					 Else
						$logonadapter = $adapterinput
						IniWrite($configfilepath, "netzwerk", "adapter1", $logonadapter)
						 ExitLoop
					 EndIf
				  EndIf
			EndSwitch
		 WEnd
GUIDelete($GUIadapterauswahl)
If $modus = $GUI_CHECKED Then
$hipathfenster = "IIIEloM Tool ©2016-2020 TESTMODUS"
Else
$hipathfenster = "IIIEloM Tool ©2016-2020"
EndIf
$logonip = IniRead($configfilepath, "netzwerk", "iphicom", "default")
$logonsubnet = IniRead($configfilepath, "netzwerk", "subnethicom", "default")
$sCMDlogon = "netsh interface ipv4 set address "&$logonadapter&" static "&$logonip&" "&$logonsubnet
Run($sCMDlogon)
Sleep(4000)


If $aktivesfenster = $telnethicom & " - PuTTY" Then
Local $PIDCOMWIN = Run($puttycommand)
WinActivate($aktivesfenster)
WinWaitActive($aktivesfenster)
WinMove($aktivesfenster, "", 0,40,600,600)
Sleep(200)
Send("root")
Send("{ENTER}")
Send("hicom")
Send("{Enter}")

   Else
   $PIDNOTEPAD = Run("notepad.exe","")
   WinWaitActive($aktivesfenster)
   WinActivate($aktivesfenster)
   WinMove($aktivesfenster, "", 0,40,600,600)
   EndIf


Local $oAppl = _Excel_Open()
Local $oWorkbook = _Excel_BookOpen($oAppl, @ScriptDir & "\Data\Berechtigung.xlsx")
Local $Berechtigungexcel = "Berechtigung - Excel"
WinMove($Berechtigungexcel, "", 600, 40, 600, 600)


GUICreate($hipathfenster, 400, 350)
   Opt("GUICoordMode", 1)
   GUICtrlCreateLabel("Berechtigungen",45,10,150)
   GUICtrlSetFont(-1, 9, 700, 4)
   $HafenohneLWL = GUICtrlCreateButton("Hafen ohne LWL", 10, 30, 150)
   $HafenmitLWL = GUICtrlCreateButton("Hafen mit LWL/SHF", 10, 60, 150)
   $See = GUICtrlCreateButton("See", 10, 90, 150)
   $RiverStateRed = GUICtrlCreateButton("River City State", 10, 120, 150)
   $Einzelberechtigung = GUICtrlCreateButton("Einzelberechtigung", 10, 150, 150)
   $Crewcall = GUICtrlCreateButton("CrewCall Zeitsteuerung", 10, 180, 150)
   GUICtrlCreateLabel("Telefoneinstellungen",220,10,150)
   GUICtrlSetFont(-1, 9, 700, 4)
   $PersiUpdate = GUICtrlCreateButton("Persi Update", 200, 30, 150)
   $ZIVO = GUICtrlCreateButton("ZIVO", 200, 60, 150)
   $Amtskarten = GUICtrlCreateButton("Amtskarten", 200, 90, 150)
   $DatumZeit = GUICtrlCreateButton("Datum/Zeit", 200, 120, 150)
   GUICtrlCreateLabel("Administration",50,220,150)
   GUICtrlSetFont(-1, 9, 700, 4)
   $Konfiguration = GUICtrlCreateButton("Konfiguration", 10, 240, 150)
   $Backup = GUICtrlCreateButton("Backup", 10, 270, 150)
   $EXIT = GUICtrlCreateButton("EXIT", 10, 300, 150)
   GUICtrlCreateLabel("Copyright 2016-2020",260,290,150)
   GUICtrlSetFont(-1, 9, 700)
   GUICtrlCreateLabel("Remkes, HB und IIIEloM",260,305,150)
   GUICtrlSetFont(-1, 9, 700)
   GUICtrlCreateLabel("Version 3.0",260,320,150)
   GUICtrlSetFont(-1, 7, 700)
   GUICtrlCreateIcon("shell32.dll", 14, 220, 295)

$excelzeilen = IniRead($configfilepath, "berechtigung", "excelzeilen", "default")

WinActivate($hipathfenster)
WinWaitActive($hipathfenster)


GUISetState()
While 1
	$msg = GUIGetMsg()
	Select
	  Case $msg = $GUI_EVENT_CLOSE
			ExitLoop
		 Case $msg = $HafenohneLWL
			HafenohneLWL()
		 Case $msg = $HafenmitLWL
			HafenmitLWL()
		 Case $msg = $See
			See()
		 Case $msg = $RiverStateRed
			RiverStateRed()
		 Case $msg = $Einzelberechtigung
			Einzelberechtigung()
		 Case $msg = $Crewcall
			Crewcall()
		 Case $msg = $PersiUpdate
			PersiUpdate()
		 Case $msg = $ZIVO
			ZIVO()
		 Case $msg = $Amtskarten
			Amtskarten()
		 Case $msg = $DatumZeit
			DatumZeit()
		 Case $msg = $Konfiguration
			Konfiguration()
		 Case $msg = $Backup
			Backup()
		 Case $msg = $EXIT
			Beenden()

	EndSelect
WEnd

Func HafenohneLWL()
   For $i = 1 To $excelzeilen
$sheetnummer = $i +1
$sheetname = $i +1
$sheetberechtigung = $i +1
$sheetart = $i +1

Local $Nummer = _Excel_RangeRead($oWorkbook, Default, "D"&$sheetnummer)
Local $Name = _Excel_RangeRead($oWorkbook, Default, "E"&$sheetname)
Local $Berechtigung = _Excel_RangeRead($oWorkbook, Default, "F"&$sheetberechtigung)
Local $Art = _Excel_RangeRead($oWorkbook, Default, "I"&$sheetart)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
WinActivate($aktivesfenster)
WinWaitActive($aktivesfenster)
$sendcommand = "AENDERN-" &$Art& ":TLNNU=" &$Nummer& ",LCOSS1="&$Berechtigung&",LCOSS2="&$Berechtigung&",LCOSD1="&$Berechtigung&",LCOSD2="&$Berechtigung&";"
ClipPut($sendcommand)
$clip = ClipGet()
MouseMove(100,100,0)
if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
	Send("^v")
Else
MouseClick($MOUSE_CLICK_RIGHT)
EndIf
Send("{Enter}")
Next
Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
Send("{Enter}")

WinActivate($hipathfenster)
WinWaitActive($hipathfenster)
EndFunc
Func HafenmitLWL ()
 For $i = 1 To $excelzeilen
$sheetnummer = $i +1
$sheetname = $i +1
$sheetberechtigung = $i +1
$sheetart = $i +1

Local $Nummer = _Excel_RangeRead($oWorkbook, Default, "D"&$sheetnummer)
Local $Name = _Excel_RangeRead($oWorkbook, Default, "E"&$sheetname)
Local $Berechtigung = _Excel_RangeRead($oWorkbook, Default, "G"&$sheetberechtigung)
Local $Art = _Excel_RangeRead($oWorkbook, Default, "I"&$sheetart)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
WinActivate($aktivesfenster)
WinWaitActive($aktivesfenster)
$sendcommand = "AENDERN-" &$Art& ":TLNNU=" &$Nummer& ",LCOSS1="&$Berechtigung&",LCOSS2="&$Berechtigung&",LCOSD1="&$Berechtigung&",LCOSD2="&$Berechtigung&";"
ClipPut($sendcommand)
$clip = ClipGet()
MouseMove(100,100,0)
if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
	Send("^v")
Else
MouseClick($MOUSE_CLICK_RIGHT)
EndIf
Send("{Enter}")
Next
Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
Send("{Enter}")

WinActivate($hipathfenster)
WinWaitActive($hipathfenster)
EndFunc
Func See()
   For $i = 1 To $excelzeilen
$sheetnummer = $i +1
$sheetname = $i +1
$sheetberechtigung = $i +1
$sheetart = $i +1

Local $Nummer = _Excel_RangeRead($oWorkbook, Default, "D"&$sheetnummer)
Local $Name = _Excel_RangeRead($oWorkbook, Default, "E"&$sheetname)
Local $Berechtigung = _Excel_RangeRead($oWorkbook, Default, "H"&$sheetberechtigung)
Local $Art = _Excel_RangeRead($oWorkbook, Default, "I"&$sheetart)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
WinActivate($aktivesfenster)
WinWaitActive($aktivesfenster)
$sendcommand = "AENDERN-" &$Art& ":TLNNU=" &$Nummer& ",LCOSS1="&$Berechtigung&",LCOSS2="&$Berechtigung&",LCOSD1="&$Berechtigung&",LCOSD2="&$Berechtigung&";"
ClipPut($sendcommand)
$clip = ClipGet()
MouseMove(100,100,0)
if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
	Send("^v")
Else
MouseClick($MOUSE_CLICK_RIGHT)
EndIf
Send("{Enter}")
Next

Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
Send("{Enter}")

WinActivate($hipathfenster)
WinWaitActive($hipathfenster)
EndFunc
Func RiverStateRed()
    Local $hGUI = GUICreate("River City State", 230, 200)
	Local $rivercitystatered = GUICtrlCreateButton("RED", 10, 10, 100, 100)
	GUICtrlSetFont(-1, 15)
	GUICtrlSetBkColor(-2, $COLOR_RED)
	Local $rivercitystateyellow = GUICtrlCreateButton("YELLOW", 120, 10, 100, 100)
	GUICtrlSetFont(-1, 15)
	GUICtrlSetBkColor(-2, $COLOR_YELLOW)
	Local $rivercitystateyellow
    Local $idBeenden = GUICtrlCreateButton("Beenden", 135, 140, 85, 25)
    GUISetState(@SW_SHOW, $hGUI)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idBeenden
                ExitLoop
			Case $rivercitystatered
			For $i = 1 To $excelzeilen
$sheetnummer = $i +1
$sheetname = $i +1
$sheetberechtigung = $i +1
$sheetart = $i +1

Local $Nummer = _Excel_RangeRead($oWorkbook, Default, "D"&$sheetnummer)
Local $Name = _Excel_RangeRead($oWorkbook, Default, "E"&$sheetname)
Local $Berechtigung = _Excel_RangeRead($oWorkbook, Default, "J"&$sheetberechtigung)
Local $Art = _Excel_RangeRead($oWorkbook, Default, "I"&$sheetart)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
WinActivate($aktivesfenster)
WinWaitActive($aktivesfenster)
$sendcommand = "AENDERN-" &$Art& ":TLNNU=" &$Nummer& ",LCOSS1="&$Berechtigung&",LCOSS2="&$Berechtigung&",LCOSD1="&$Berechtigung&",LCOSD2="&$Berechtigung&";"
ClipPut($sendcommand)
$clip = ClipGet()
MouseMove(100,100,0)
if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
	Send("^v")
Else
MouseClick($MOUSE_CLICK_RIGHT)
EndIf
Send("{Enter}")
Next
Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
Send("{Enter}")
WinActivate("River City State")
WinWaitActive("River City State")

Case $rivercitystateyellow
For $i = 1 To $excelzeilen
$sheetnummer = $i +1
$sheetname = $i +1
$sheetberechtigung = $i +1
$sheetart = $i +1

Local $Nummer = _Excel_RangeRead($oWorkbook, Default, "D"&$sheetnummer)
Local $Name = _Excel_RangeRead($oWorkbook, Default, "E"&$sheetname)
Local $Berechtigung = _Excel_RangeRead($oWorkbook, Default, "K"&$sheetberechtigung)
Local $Art = _Excel_RangeRead($oWorkbook, Default, "I"&$sheetart)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
WinActivate($aktivesfenster)
WinWaitActive($aktivesfenster)
$sendcommand = "AENDERN-" &$Art& ":TLNNU=" &$Nummer& ",LCOSS1="&$Berechtigung&",LCOSS2="&$Berechtigung&",LCOSD1="&$Berechtigung&",LCOSD2="&$Berechtigung&";"
ClipPut($sendcommand)
$clip = ClipGet()
MouseMove(100,100,0)
if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
	Send("^v")
Else
MouseClick($MOUSE_CLICK_RIGHT)
EndIf
Send("{Enter}")
Next
Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
Send("{Enter}")
WinActivate("River City State")
WinWaitActive("River City State")
        EndSwitch
	 WEnd

    GUIDelete($hGUI)
   WinActivate($hipathfenster)
   WinWaitActive($hipathfenster)
EndFunc
Func Einzelberechtigung()
   Local $GUIEINZELBERECHTIGUNG = GUICreate("Einzelberechtigung", 250, 250)
	GUICtrlCreateLabel("Rufnummer:", 10, 10, 100, 20)
	Local $idrufnummer = GUICtrlCreateInput("", 90, 10, 150, 20)
	GUICtrlCreateLabel("Art:", 10, 40, 100, 20)
	Local $idart = GUICtrlCreateCombo("", 90, 40, 150, 20)
		GUICtrlSetData(-1, "Analog|Digital")
	GUICtrlCreateLabel("Berechtigung", 10, 70, 100, 20)
	Local $idberechtigung = GUICtrlCreateCombo("", 90, 70, 150, 20)
		GUICtrlSetData(-1, "6 - Intern|7 - Extern LAK|11 - Extern 51|12 - Extern 52|13 - Extern 53|14 - Extern 54")
	GUICtrlCreateLabel("Hinweis:", 10, 100, 230)
	GUICtrlCreateLabel("Ob ein Anschluss Analog oder Digital ist, lässt sich der Exceltabelle entnehmen.", 10, 115, 230, 30)
	GUICtrlCreateLabel("Analog = SCSU", 10, 150, 230)
	GUICtrlCreateLabel("Digital = SBCSU", 10, 165, 230)
	GUICtrlCreateLabel("Änderungen werden nicht in der Excelliste gespeichert", 10, 180, 230, 30)
    Local $idSend = GUICtrlCreateButton("Senden", 10, 220, 85, 25)
	Local $idClose = GUICtrlCreateButton("Beenden", 150, 220, 85, 25)
    GUISetState(@SW_SHOW, $GUIEINZELBERECHTIGUNG)

    While 1
        Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $idClose
				WinActivate($hipathfenster)
				WinActivate($hipathfenster)
                ExitLoop
			Case $idSend
				$idrufnummerread = GUICtrlRead($idrufnummer)
				$idartread = GUICtrlRead($idart)
				$idberechtigungread = GUICtrlRead($idberechtigung)

				if not $idrufnummerread = 0 Then
					$idrufnummersend = $idrufnummerread
				if $idartread = "Analog" or $idartread = "Digital" Then
					if $idartread = "Analog" Then $idartsend = "SCSU"
					if $idartread = "Digital" Then $idartsend = "SBCSU"
				if not $idberechtigungread = "" Then
					if $idberechtigungread = "6 - Intern" Then $idberechtigungsend = "6"
					if $idberechtigungread = "7 - Extern LAK" Then $idberechtigungsend = "7"
					if $idberechtigungread = "11 - Extern 51" Then $idberechtigungsend = "11"
					if $idberechtigungread = "12 - Extern 52" Then $idberechtigungsend = "12"
					if $idberechtigungread = "13 - Extern 53" Then $idberechtigungsend = "13"
					if $idberechtigungread = "14 - Extern 54" Then $idberechtigungsend = "14"
				#MsgBox($MB_OK, "Test", "AENDERN-" &$idartsend& ":TLNNU=" &$idrufnummersend& ",LCOSS1="&$idberechtigungsend&",LCOSS2="&$idberechtigungsend&",LCOSD1="&$idberechtigungsend&",LCOSD2="&$idberechtigungsend&";")


				if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
				WinActivate($aktivesfenster)
				WinWaitActive($aktivesfenster)
				$sendcommand = "AENDERN-" &$idartsend& ":TLNNU=" &$idrufnummersend& ",LCOSS1="&$idberechtigungsend&",LCOSS2="&$idberechtigungsend&",LCOSD1="&$idberechtigungsend&",LCOSD2="&$idberechtigungsend&";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				MouseMove(100,100,0)
				if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
				Send("^v")
				Else
				MouseClick($MOUSE_CLICK_RIGHT)
				EndIf
				Send("{Enter}")
				Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
				Send("{Enter}")

				WinActivate($GUIEINZELBERECHTIGUNG)
				WinWaitActive($GUIEINZELBERECHTIGUNG)

				Else
					MsgBox($MB_OK, "Fehler", "Bitte Berechtigung wählen")
				EndIf
				Else
					MsgBox($MB_OK, "Fehler", "Bitte Art wählen")
				EndIf
				Else
					MsgBox($MB_OK, "Fehler", "Bitte Rufnummer eingeben")
				EndIf

        EndSwitch
    WEnd

    GUIDelete($GUIEINZELBERECHTIGUNG)
EndFunc
Func CrewCall()
Local $GUICREWCALL = GUICreate("CrewCall", 400, 350, 287, 163)
GUICtrlCreateLabel("Crewcall AUS", 120, 10, 81, 17)
GUICtrlCreateLabel("Crewcall AN", 264, 10, 73, 17)
GUICtrlCreateLabel("Tag", 16, 30, 26, 17)
GUICtrlCreateLabel("Stunde", 112, 30, 44, 17)
GUICtrlCreateLabel("Minute", 160, 30, 42, 17)
GUICtrlCreateLabel("Stunde", 256, 30, 44, 17)
GUICtrlCreateLabel("Minute", 304, 30, 42, 17)

GUICtrlCreateLabel("Montag", 16, 60, 46, 17)
GUICtrlCreateLabel("Dienstag", 16, 90, 46, 17)
GUICtrlCreateLabel("Mittwoch", 16, 120, 47, 17)
GUICtrlCreateLabel("Donnerstag", 16, 150, 59, 17)
GUICtrlCreateLabel("Freitag", 16, 180, 36, 17)
GUICtrlCreateLabel("Samstag", 16, 210, 45, 17)

$DropdownStunde = "00|01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23"
$DropdownMinute = "00|10|20|30|40|50"
#############Montag################
$StundeMontagAUS = GUICtrlCreateCombo("", 104, 56, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteMontagAUS = GUICtrlCreateCombo("", 154, 56, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
$StundeMontagEIN = GUICtrlCreateCombo("", 250, 56, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteMontagEIN = GUICtrlCreateCombo("", 300, 56, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
#############Dienstag################
$StundeDienstagAUS = GUICtrlCreateCombo("", 104, 86, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteDienstagAUS = GUICtrlCreateCombo("", 154, 86, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
$StundeDienstagEIN = GUICtrlCreateCombo("", 250, 86, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteDienstagEIN = GUICtrlCreateCombo("", 300, 86, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
#############Mittwoch################
$StundeMittwochAUS = GUICtrlCreateCombo("", 104, 116, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteMittwochAUS = GUICtrlCreateCombo("", 154, 116, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
$StundeMittwochEIN = GUICtrlCreateCombo("", 250, 116, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteMittwochEIN = GUICtrlCreateCombo("", 300, 116, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
#############Donnerstag################
$StundeDonnerstagAUS = GUICtrlCreateCombo("", 104, 146, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteDonnerstagAUS = GUICtrlCreateCombo("", 154, 146, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
$StundeDonnerstagEIN = GUICtrlCreateCombo("", 250, 146, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteDonnerstagEIN = GUICtrlCreateCombo("", 300, 146, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
#############Freitag################
$StundeFreitagAUS = GUICtrlCreateCombo("", 104, 176, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteFreitagAUS = GUICtrlCreateCombo("", 154, 176, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
$StundeFreitagEIN = GUICtrlCreateCombo("", 250, 176, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteFreitagEIN = GUICtrlCreateCombo("", 300, 176, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
#############Samstag################
$StundeSamstagAUS = GUICtrlCreateCombo("", 104, 206, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteSamstagAUS = GUICtrlCreateCombo("", 154, 206, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
$StundeSamstagEIN = GUICtrlCreateCombo("", 250, 206, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownStunde)
$MinuteSamstagEIN = GUICtrlCreateCombo("", 300, 206, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $DropdownMinute)
GUICtrlCreateLabel("Bei Änderungen müssen alle Felder befüllt werden. Es dürfen nur die Werte im  "& @CRLF & "Dropdown-Menü genutzt werden.", 10, 240)

    Local $idSend = GUICtrlCreateButton("Senden", 50, 300, 85, 25)
	Local $idwerk = GUICtrlCreateButton("Grundeinstellungen", 140, 300, 105, 25)
	Local $idClose = GUICtrlCreateButton("Beenden", 250, 300, 85, 25)
GUISetState(@SW_SHOW)
				if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $idClose
				WinActivate($hipathfenster)
				WinActivate($hipathfenster)
                ExitLoop
		Case $idSend
			WinActivate($aktivesfenster)
			WinWaitActive($aktivesfenster)
			$StundeMontagAUSread = GUICtrlRead($StundeMontagAUS)
			$MinuteMontagAUSread = Guictrlread($MinuteMontagAUS)
			$StundeMontagEINread = GUICtrlRead($StundeMontagEIN)
			$MinuteMontagEINread = Guictrlread($MinuteMontagEIN)
			$StundeDienstagAUSread = GUICtrlRead($StundeDienstagAUS)
			$MinuteDienstagAUSread = Guictrlread($MinuteDienstagAUS)
			$StundeDienstagEINread = GUICtrlRead($StundeDienstagEIN)
			$MinuteDienstagEINread = Guictrlread($MinuteDienstagEIN)
			$StundeMittwochAUSread = GUICtrlRead($StundeMittwochAUS)
			$MinuteMittwochAUSread = Guictrlread($MinuteMittwochAUS)
			$StundeMittwochEINread = GUICtrlRead($StundeMittwochEIN)
			$MinuteMittwochEINread = Guictrlread($MinuteMittwochEIN)
			$StundeDonnerstagAUSread = GUICtrlRead($StundeDonnerstagAUS)
			$MinuteDonnerstagAUSread = Guictrlread($MinuteDonnerstagAUS)
			$StundeDonnerstagEINread = GUICtrlRead($StundeDonnerstagEIN)
			$MinuteDonnerstagEINread = Guictrlread($MinuteDonnerstagEIN)
			$StundeFreitagAUSread = GUICtrlRead($StundeFreitagAUS)
			$MinuteFreitagAUSread = Guictrlread($MinuteFreitagAUS)
			$StundeFreitagEINread = GUICtrlRead($StundeFreitagEIN)
			$MinuteFreitagEINread = Guictrlread($MinuteFreitagEIN)
			$StundeSamstagAUSread = GUICtrlRead($StundeSamstagAUS)
			$MinuteSamstagAUSread = Guictrlread($MinuteSamstagAUS)
			$StundeSamstagEINread = GUICtrlRead($StundeSamstagEIN)
			$MinuteSamstagEINread = Guictrlread($MinuteSamstagEIN)


			 For $i = 0 To 11
			 $POS = $i +1
				if $POS = "1" or $POS = "3" or $POS = "5" or $POS = "7" or $POS = "9" or $POS = "11"  Then $LCOSS1 = "6"
				if $POS = "2" or $POS = "4" or $POS = "6" or $POS = "8" or $POS = "10" or $POS = "12" Then $LCOSS1 = "14"
				if $POS = "1" or $POS = "2" Then $WOTAG = "1"
				if $POS = "3" or $POS = "4" Then $WOTAG = "2"
				if $POS = "5" or $POS = "6" Then $WOTAG = "3"
				if $POS = "7" or $POS = "8" Then $WOTAG = "4"
				if $POS = "9" or $POS = "10" Then $WOTAG = "5"
				if $POS = "11" or $POS = "12" Then $WOTAG = "6"
				if $POS = "1" Then
					$crewcallminute = $MinuteMontagAUSread
					$crewcallstunde = $StundeMontagAUSread
				EndIf
				if $POS = "2" Then
					$crewcallminute = $MinuteMontagEINread
					$crewcallstunde = $StundeMontagEINread
				EndIf
				if $POS = "3" Then
					$crewcallminute = $MinuteDienstagAUSread
					$crewcallstunde = $StundeDienstagAUSread
				EndIf
				if $POS = "4" Then
					$crewcallminute = $MinuteDienstagEINread
					$crewcallstunde = $StundeDienstagEINread
				EndIf
				if $POS = "5" Then
					$crewcallminute = $MinuteMittwochAUSread
					$crewcallstunde = $StundeMittwochAUSread
				EndIf
				if $POS = "6" Then
					$crewcallminute = $MinuteMittwochEINread
					$crewcallstunde = $StundeMittwochEINread
				EndIf
				if $POS = "7" Then
					$crewcallminute = $MinuteDonnerstagAUSread
					$crewcallstunde = $StundeDonnerstagAUSread
				EndIf
				if $POS = "8" Then
					$crewcallminute = $MinuteDonnerstagEINread
					$crewcallstunde = $StundeDonnerstagEINread
				EndIf
				if $POS = "9" Then
					$crewcallminute = $MinuteFreitagAUSread
					$crewcallstunde = $StundeFreitagAUSread
				EndIf
				if $POS = "10" Then
					$crewcallminute = $MinuteFreitagEINread
					$crewcallstunde = $StundeFreitagEINread
				EndIf
				if $POS = "11" Then
					$crewcallminute = $MinuteSamstagAUSread
					$crewcallstunde = $StundeSamstagAUSread
				EndIf
				if $POS = "12" Then
					$crewcallminute = $MinuteSamstagEINread
					$crewcallstunde = $StundeSamstagEINread
				EndIf
				$sendcommand = "AENDERN-CRON:POS="&$POS&",AUSF=R,MINUTE="&$crewcallminute&",STUNDE="&$crewcallstunde&",WOTAG="&$WOTAG&",KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1="&$LCOSS1&";"";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				MouseMove(100,100,0)
				if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
				Send("^v")
				Else
				MouseClick($MOUSE_CLICK_RIGHT)
				EndIf
				Send("{Enter}")
			 Next
			Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
			Send("{Enter}")
ExitLoop
WinActivate($hipathfenster)
WinWaitActive($hipathfenster)
		Case $idwerk
			WinActivate($aktivesfenster)
			WinWaitActive($aktivesfenster)
			$sendcommand = "AENDERN-CRON:POS=1,AUSF=R,MINUTE=00,STUNDE=08,WOTAG=1,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=6;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=2,AUSF=R,MINUTE=50,STUNDE=16,WOTAG=1,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=14;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=3,AUSF=R,MINUTE=00,STUNDE=08,WOTAG=2,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=6;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=4,AUSF=R,MINUTE=50,STUNDE=16,WOTAG=2,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=14;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=5,AUSF=R,MINUTE=00,STUNDE=08,WOTAG=3,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=6;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=6,AUSF=R,MINUTE=50,STUNDE=16,WOTAG=3,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=14;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=7,AUSF=R,MINUTE=00,STUNDE=08,WOTAG=4,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=6;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=8,AUSF=R,MINUTE=50,STUNDE=16,WOTAG=4,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=14;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=9,AUSF=R,MINUTE=00,STUNDE=08,WOTAG=5,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=6;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=10,AUSF=R,MINUTE=50,STUNDE=16,WOTAG=5,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=14;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=11,AUSF=R,MINUTE=00,STUNDE=08,WOTAG=6,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=6;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			$sendcommand = "AENDERN-CRON:POS=12,AUSF=R,MINUTE=50,STUNDE=16,WOTAG=6,KOMMANDO="&"""AENDERN-SCSU:TLNNU=362,LCOSS1=14;"";";"
				ClipPut($sendcommand)
				$clip = ClipGet()
				Send("^v")
				Send("{ENTER}")
			ExitLoop
WinActivate($hipathfenster)
WinWaitActive($hipathfenster)
	EndSwitch
WEnd
    GUIDelete($GUICREWCALL)
WinActivate($hipathfenster)
WinWaitActive($hipathfenster)
EndFunc


Func PersiUpdate()
   For $i = 1 To $excelzeilen
$sheetnummer = $i +1
$sheetname = $i +1
$sheetberechtigung = $i +1
$sheetart = $i +1

$Nummer = _Excel_RangeRead($oWorkbook, Default, "D"&$sheetnummer)
$Name = _Excel_RangeRead($oWorkbook, Default, "E"&$sheetname)
$Berechtigung = _Excel_RangeRead($oWorkbook, Default, "H"&$sheetberechtigung)
$Art = _Excel_RangeRead($oWorkbook, Default, "I"&$sheetart)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
WinActivate($aktivesfenster)
WinWaitActive($aktivesfenster)
$sendcommand = "AENDERN-PERSI:TYP=NAME,RUFNU="&$Nummer&",NEUNAME="&$Name&";"
ClipPut($sendcommand)
$clip = ClipGet()
MouseMove(100,100,0)
if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
	Send("^v")
Else
MouseClick($MOUSE_CLICK_RIGHT)
EndIf
Send("{Enter}")
Next

Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
Send("{Enter}")

WinActivate($hipathfenster)
WinWaitActive($hipathfenster)
EndFunc
Func ZIVO()
       Local $hGUI = GUICreate("SHF-Auswahl", 300, 200)

    Local $idZielrufnummer = GUICtrlCreateLabel("SHF-Leitung",10,10)
    Local $idComboBox = GUICtrlCreateCombo("", 90, 10, 100, 20)
	Local $idZielrufnummer = GUICtrlCreateLabel("Zielrufnummer",10,40)
	Local $idFile = GUICtrlCreateInput("", 90, 40, 200, 20)
    Local $idSend = GUICtrlCreateButton("Senden", 10, 170, 85, 25)
	Local $idwerk = GUICtrlCreateButton("Grundeinstellungen", 100, 170, 105, 25)
	Local $idClose = GUICtrlCreateButton("Beenden", 210, 170, 85, 25)


	GUICtrlCreateLabel("Grundeinstellungen:",10,70)
	$shf1ziel = IniRead($configfilepath, "shfauswahl", "shf1ziel", "default")
	$shf1nummer = IniRead($configfilepath, "shfauswahl", "shf1nummer", "default")
	GUICtrlCreateLabel("SHF 1 auf "&$shf1ziel&" - Nummer "&$shf1nummer ,10,85)
    $shf2ziel = IniRead($configfilepath, "shfauswahl", "shf2ziel", "default")
	$shf2nummer = IniRead($configfilepath, "shfauswahl", "shf2nummer", "default")
	GUICtrlCreateLabel("SHF 2 auf "&$shf2ziel&" - Nummer "&$shf2nummer ,10,100)
	$shf3ziel = IniRead($configfilepath, "shfauswahl", "shf3ziel", "default")
	$shf3nummer = IniRead($configfilepath, "shfauswahl", "shf3nummer", "default")
	GUICtrlCreateLabel("SHF 3 auf "&$shf3ziel&" - Nummer "&$shf3nummer ,10,115)
	$shf4ziel = IniRead($configfilepath, "shfauswahl", "shf4ziel", "default")
	$shf4nummer = IniRead($configfilepath, "shfauswahl", "shf4nummer", "default")
	GUICtrlCreateLabel("SHF 4 auf "&$shf4ziel&" - Nummer "&$shf4nummer ,10,130)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)
    GUICtrlSetData($idComboBox, "SHF 1|SHF 2|SHF 3|SHF 4")

    GUISetState(@SW_SHOW, $hGUI)
    Local $sComboRead = ""

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idClose
                ExitLoop

            Case $idComboBox
                $sComboRead = GUICtrlRead($idComboBox)
				$shf1variabel = IniRead($configfilepath, "shfauswahl", "shf1variabel", "default")
				$shf2variabel = IniRead($configfilepath, "shfauswahl", "shf2variabel", "default")
				$shf3variabel = IniRead($configfilepath, "shfauswahl", "shf3variabel", "default")
				$shf4variabel = IniRead($configfilepath, "shfauswahl", "shf4variabel", "default")
					If $sComboRead = "SHF 1" Then GUICtrlCreateLabel("Aktuell: "&$shf1variabel, 200, 10)
					If $sComboRead = "SHF 2" Then GUICtrlCreateLabel("Aktuell: "&$shf2variabel, 200, 10)
					If $sComboRead = "SHF 3" Then GUICtrlCreateLabel("Aktuell: "&$shf3variabel, 200, 10)
					If $sComboRead = "SHF 4" Then GUICtrlCreateLabel("Aktuell: "&$shf4variabel, 200, 10)
			Case $idwerk
			   if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
				WinActivate($aktivesfenster)
				WinWaitActive($aktivesfenster)
			   $sendcommand = "AENDERN-TACSU:LAGE=1-1-97-0,GERTYP=AS,ZIVO="&$shf1ziel&";"
			   ClipPut($sendcommand)
			   $clip = ClipGet()
			   MouseMove(100,100,0)
				if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
				Send("^v")
				Else
				MouseClick($MOUSE_CLICK_RIGHT)
				EndIf
				Send("{Enter}")
			   IniWrite($configfilepath, "shfauswahl", "shf1variabel", $shf1ziel)
			   $sendcommand = "AENDERN-TACSU:LAGE=1-1-97-1,GERTYP=AS,ZIVO="&$shf2ziel&";"
			   ClipPut($sendcommand)
			   $clip = ClipGet()
			   MouseMove(100,100,0)
				if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
				Send("^v")
				Else
				MouseClick($MOUSE_CLICK_RIGHT)
				EndIf
				Send("{Enter}")
			   IniWrite($configfilepath, "shfauswahl", "shf2variabel", $shf2ziel)
			   $sendcommand = "AENDERN-TACSU:LAGE=1-1-97-2,GERTYP=AS,ZIVO="&$shf3ziel&";"
			   ClipPut($sendcommand)
			   $clip = ClipGet()
			   MouseMove(100,100,0)
				if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
				Send("^v")
				Else
				MouseClick($MOUSE_CLICK_RIGHT)
				EndIf
				Send("{Enter}")
			   IniWrite($configfilepath, "shfauswahl", "shf3variabel", $shf3ziel)
			   $sendcommand = "AENDERN-TACSU:LAGE=1-1-97-3,GERTYP=AS,ZIVO="&$shf4ziel&";"
			   ClipPut($sendcommand)
			   $clip = ClipGet()
			   MouseMove(100,100,0)
				if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
				Send("^v")
				Else
				MouseClick($MOUSE_CLICK_RIGHT)
				EndIf
				Send("{Enter}")
			   IniWrite($configfilepath, "shfauswahl", "shf4variabel", $shf4ziel)
			   Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
			   Send("{Enter}")
			   WinActivate("SHF-Auswahl")

			Case $idSend
			   if NOT $sComboRead = 0 Then
			   $ZIVOREAD = GUICtrlRead($idFile)
			   If GUICtrlRead($idComboBox) = "SHF 1" Then Local $idLage = "1-1-97-0"
			   If GUICtrlRead($idComboBox) = "SHF 1" Then IniWrite($configfilepath, "shfauswahl", "shf1variabel", $ZIVOREAD)
			   If GUICtrlRead($idComboBox) = "SHF 2" Then Local $idLage = "1-1-97-1"
			   If GUICtrlRead($idComboBox) = "SHF 2" Then IniWrite($configfilepath, "shfauswahl", "shf2variabel", $ZIVOREAD)
			   If GUICtrlRead($idComboBox) = "SHF 3" Then Local $idLage = "1-1-97-2"
			   If GUICtrlRead($idComboBox) = "SHF 3" Then IniWrite($configfilepath, "shfauswahl", "shf3variabel", $ZIVOREAD)
			   If GUICtrlRead($idComboBox) = "SHF 4" Then Local $idLage = "1-1-97-3"
			   If GUICtrlRead($idComboBox) = "SHF 4" Then IniWrite($configfilepath, "shfauswahl", "shf4variabel", $ZIVOREAD)
			   if NOT $ZIVOREAD = 0 Then
			   if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
			   WinActivate($aktivesfenster)
			   WinWaitActive($aktivesfenster)
			   $sendcommand = "AENDERN-TACSU:LAGE="&$idLage&",GERTYP=AS,ZIVO="&GUICtrlRead($idFile)&";"
			   ClipPut($sendcommand)
			   $clip = ClipGet()
			   MouseMove(100,100,0)
				if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
				Send("^v")
				Else
				MouseClick($MOUSE_CLICK_RIGHT)
				EndIf
				Send("{Enter}")
			   Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
			   Send("{Enter}")
			   WinActivate("SHF-Auswahl")

			Else
				  MsgBox($MB_ICONERROR,"Fehler","Bitte eine Zielrufnummer eingeben")
				  Endif
			Else
			    MsgBox($MB_ICONERROR,"Fehler","Bitte eine SHF-Leitung wählen")
			   EndIf
        EndSwitch
    WEnd
    GUIDelete($hGUI)

WinActivate($hipathfenster)
WinWaitActive($hipathfenster)
EndFunc
Func Amtskarten()
       Local $hGUI = GUICreate("Amtskarten",300,150)
    Local $LAmtskarte = GUICtrlCreateLabel("Amtskarte",10,10)
    Local $amtskarte = GUICtrlCreateCombo("",80,10,200)
    GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	GUICtrlSetData($amtskarte, "Karte 1 (links)|Karte 2 (rechts)")
	Local $LAmtsleitung = GUICtrlCreateLabel("Amtsleitung",10,40)
    Local $amtsleitung = GUICtrlCreateCombo("",80,40,200)
    GUICtrlSetState(-1, $GUI_DROPACCEPTED)
    GUICtrlSetData($amtsleitung, "Amt 1|Amt 2|Amt 3|Amt 4|Alle")
	Local $LStatus = GUICtrlCreateLabel("Status",10,70)
    Local $amtsstatus = GUICtrlCreateCombo("",80,70,200)
    GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	GUICtrlSetData($amtsstatus, "Einschalten|Ausschalten")
    Local $Senden = GUICtrlCreateButton("Senden",100, 120, 85, 25)
    Local $Beenden = GUICtrlCreateButton("Beenden", 200, 120, 85, 25)




    GUISetState(@SW_SHOW, $hGUI)

Local $amtskarteread = ""
Local $amtsleitungread = ""
Local $amtsstatusread = ""


    While 1
        Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE, $Beenden
                ExitLoop


			case $amtskarte
			   $amtskarteread = GUICtrlRead($amtskarte)
			case $amtsleitung
			   $amtsleitungread = GUICtrlRead($amtsleitung)
			case $amtsstatus
			   $amtsstatusread = GUICtrlRead($amtsstatus)

			case $Senden
			   If NOT $amtskarteread = 0 Then
				  if $amtskarteread = "Karte 1 (links)" Then Local $idamtskarte = ("25")
				  if $amtskarteread = "Karte 2 (rechts)" Then Local $idamtskarte = ("31")
			   if not $amtsleitungread = 0 Then
				  if $amtsleitungread = "Amt 1" Then Local $idamtsleitung = ("0")
				  if $amtsleitungread = "Amt 2" Then Local $idamtsleitung = ("1")
				  if $amtsleitungread = "Amt 3" Then Local $idamtsleitung = ("2")
				  if $amtsleitungread = "Amt 4" Then Local $idamtsleitung = ("3")
			   if not $amtsstatusread = 0 Then
				  if $amtsstatusread = "Einschalten" Then Local $idamtsstatus = ("EINSCHALTEN")
				  if $amtsstatusread = "Ausschalten" Then Local $idamtsstatus = ("AUSSCHALTEN")
			   if $amtsleitungread = "Alle" Then
				if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
				  WinActivate($aktivesfenster)
				  WinWaitActive($aktivesfenster)
				  $sendcommand = $idamtsstatus&"-DSSU:TYP=LAGE,LAGE1=1-1-"&$idamtskarte&"-0;"
				  ClipPut($sendcommand)
				  $clip = ClipGet()
				  MouseMove(100,100,0)
					if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
					Send("^v")
					Else
					MouseClick($MOUSE_CLICK_RIGHT)
					EndIf
					Send("{Enter}")
					Sleep(100)
				  $sendcommand = $idamtsstatus&"-DSSU:TYP=LAGE,LAGE1=1-1-"&$idamtskarte&"-1;"
				  ClipPut($sendcommand)
				  $clip = ClipGet()
				  MouseMove(100,100,0)
					if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
					Send("^v")
					Else
					MouseClick($MOUSE_CLICK_RIGHT)
					EndIf
					Send("{Enter}")
					Sleep(100)
				  $sendcommand = $idamtsstatus&"-DSSU:TYP=LAGE,LAGE1=1-1-"&$idamtskarte&"-2;"
				  ClipPut($sendcommand)
				  $clip = ClipGet()
				  MouseMove(100,100,0)
					if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
					Send("^v")
					Else
					MouseClick($MOUSE_CLICK_RIGHT)
					EndIf
					Send("{Enter}")
					Sleep(100)
				  $sendcommand = $idamtsstatus&"-DSSU:TYP=LAGE,LAGE1=1-1-"&$idamtskarte&"-3;"
				  ClipPut($sendcommand)
				  $clip = ClipGet()
				  MouseMove(100,100,0)
					if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
					Send("^v")
					Else
					MouseClick($MOUSE_CLICK_RIGHT)
					EndIf
					Send("{Enter}")
					Sleep(100)
				  Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
				  Send("{Enter}")
				  WinActivate("Amtskarten")
				  WinWaitActive("Amtskarten")
			   Else
				  if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
				  WinActivate($aktivesfenster)
				  WinWaitActive($aktivesfenster)
				  $sendcommand = $idamtsstatus&"-DSSU:TYP=LAGE,LAGE1=1-1-"&$idamtskarte&"-"&$idamtsleitung&";"
				  ClipPut($sendcommand)
				  $clip = ClipGet()
				  MouseMove(100,100,0)
					if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
					Send("^v")
					Else
					MouseClick($MOUSE_CLICK_RIGHT)
					EndIf
					Send("{Enter}")
				  Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
				  Send("{Enter}")
				  WinActivate("Amtskarten")
				  WinWaitActive("Amtskarten")
			   endif

				  Else
			   MsgBox($MB_ICONERROR,"Fehler","Bitte Amtsstatus wählen.")
			   EndIf
				  Else
			   MsgBox($MB_ICONERROR,"Fehler","Bitte Amtsleitung wählen.")
			   EndIf
				  Else
			   MsgBox($MB_ICONERROR,"Fehler","Bitte Amtskarte wählen.")
			   endif
        EndSwitch
    WEnd
GUIDelete($hGUI)
WinActivate($hipathfenster)
WinWaitActive($hipathfenster)
EndFunc
Func DatumZeit()
    Local $datumzeit = GUICreate("Datum/Zeit", 300, 200)

    ; Create a combobox control.
	GUICtrlCreateLabel("Tag", 10, 10, 50, 20)
    Local $Tag = GUICtrlCreateCombo("01", 70, 10, 185, 20)
	GUICtrlCreateLabel("Monat", 10, 35, 50, 20)
    Local $Monat = GUICtrlCreateCombo("01", 70, 35, 185, 20)
	GUICtrlCreateLabel("Jahr", 10, 60, 50, 20)
    Local $Jahr = GUICtrlCreateCombo("2018", 70, 60, 185, 20)
    GUICtrlCreateLabel("Stunde", 10, 85, 50, 20)
    Local $Stunde = GUICtrlCreateCombo("00", 70, 85, 185, 20)
    GUICtrlCreateLabel("Minute", 10, 110, 50, 20)
    Local $Minute = GUICtrlCreateCombo("00", 70, 110, 185, 20)
    GUICtrlCreateLabel("Sekunde", 10, 135, 50, 20)
    Local $Sekunde = GUICtrlCreateCombo("00", 70, 135, 185, 20)

    Local $idSend = GUICtrlCreateButton("Senden", 120, 170, 85, 25)
    Local $idClose = GUICtrlCreateButton("Beenden", 210, 170, 85, 25)

    ; Add additional items to the combobox.
    GUICtrlSetData($Tag, "02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23|24|25|26|27|28|29|30|31")
    GUICtrlSetData($Monat, "02|03|04|05|06|07|08|09|10|11|12")
    GUICtrlSetData($Jahr, "2019|2020|2021|2022|2023")
    GUICtrlSetData($Stunde, "01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23|")
    GUICtrlSetData($Minute, "01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23|24|25|26|27|28|29|30|31|32|33|34|35|36|37|38|39|40|41|42|43|44|45|46|47|48|49|50|51|52|53|54|55|56|57|58|59")
    GUICtrlSetData($Sekunde, "01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23|24|25|26|27|28|29|30|31|32|33|34|35|36|37|38|39|40|41|42|43|44|45|46|47|48|49|50|51|52|53|54|55|56|57|58|59")
    ; Display the GUI.
    GUISetState(@SW_SHOW, $datumzeit)


    ; Loop until the user exits.
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idClose
                ExitLoop

            Case $idSend
                $ReadTag = GUICtrlRead($Tag)
                $ReadMonat = GUICtrlRead($Monat)
                $ReadJahr = GUICtrlRead($Jahr)
                $ReadStunde = GUICtrlRead($Stunde)
                $ReadMinute = GUICtrlRead($Minute)
                $ReadSekunde = GUICtrlRead($Sekunde)
				if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
				  WinActivate($aktivesfenster)
				  WinWaitActive($aktivesfenster)
				  $sendcommand = "AE-DATE:"&$ReadJahr&","&$ReadMonat&","&$ReadTag&","&$ReadStunde&","&$ReadMinute&","&$ReadSekunde&";"
				  ClipPut($sendcommand)
				  $clip = ClipGet()
				  MouseMove(100,100,0)
if $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
	Send("^v")
Else
MouseClick($MOUSE_CLICK_RIGHT)
EndIf
Send("{Enter}")
				  Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
				  Send("{Enter}")
				  WinActivate("Datum/Zeit")
				  WinWaitActive("Datum/Zeit")
        EndSwitch
    WEnd

    ; Delete the previous GUI and all controls.
    GUIDelete($datumzeit)
	WinActivate($hipathfenster)
    WinWaitActive($hipathfenster)
EndFunc
Func Konfiguration()
       Local $GUIkonfiguration = GUICreate("Konfiguration", 420, 450)
	GUICtrlCreateLabel("Netzwerk Grundeinstellungen", 30, 10)

GUICtrlCreateLabel("IP HiPath-Netzwerk", 10, 30)
	$iphicom = IniRead($configfilepath,"netzwerk","iphicom","default")
	GUICtrlCreateLabel($iphicom, 150, 30)
	$iphicominput = GUICtrlCreateInput("",250,26,130)
GUICtrlCreateLabel("Subnet HiPath-Netzwerk", 10, 50)
	$subnethicom = IniRead($configfilepath,"netzwerk","subnethicom","default")
	GUICtrlCreateLabel($subnethicom, 150, 50)
	$subnethicominput = GUICtrlCreateInput("",250,46,130)
GUICtrlCreateLabel("IP HiPath-Telnet", 10, 70)
    $telnethicomip = IniRead($configfilepath,"netzwerk","telnethicom","default")
	GUICtrlCreateLabel($telnethicomip, 150, 70)
    $telnethicomipinput = GUICtrlCreateInput("",250,66,130)
GUICtrlCreateLabel("IP SeaTel", 10, 90)
    $ipseatel = IniRead($configfilepath,"netzwerk","ipseatel","default")
	GUICtrlCreateLabel($ipseatel, 150, 90)
    $ipseatelinput = GUICtrlCreateInput("",250,86,130)

	$logoffdhcp = IniRead($configfilepath, "netzwerk", "dhcp", "default")
 if $logoffdhcp = 1 Then
    Local $DHCPCHECKBOX = GUICtrlCreateCheckbox("DHCP", 70, 86)
    GUICtrlSetState($DHCPCHECKBOX, $GUI_CHECKED)
 elseif $logoffdhcp = 0 Then
    Local $DHCPCHECKBOX = GUICtrlCreateCheckbox("DHCP", 70, 86)
    GUICtrlSetState($DHCPCHECKBOX, $GUI_UNCHECKED)
 EndIf

GUICtrlCreateLabel("Subnet SeaTel", 10, 110)
	$subnetseatel = IniRead($configfilepath,"netzwerk","subnetseatel","default")
    GUICtrlCreateLabel($subnetseatel, 150, 110)
	$subnetseatelinput = GUICtrlCreateInput("",250,106,130)
GUICtrlCreateLabel("Netzwerkadapter 1", 10, 130)
    $adapter1 = IniRead($configfilepath,"netzwerk","adapter1","default")
	GUICtrlCreateLabel($adapter1, 150, 130)
	$adapterinput1 = GUICtrlCreateInput("",250,126,130)
GUICtrlCreateLabel("Netzwerkadapter 2", 10, 150)
    $adapter2 = IniRead($configfilepath,"netzwerk","adapter2","default")
	GUICtrlCreateLabel($adapter2, 150, 150)
	$adapterinput2 = GUICtrlCreateInput("",250,146,130)

    GUICtrlCreateLabel("SHF-Auswahl Grundeinstellungen", 30, 170)

GUICtrlCreateLabel("Zielrufnummer für SHF 1", 10, 190)
	$shf1ziel = IniRead($configfilepath,"shfauswahl","shf1ziel","default")
	GUICtrlCreateLabel($shf1ziel, 150, 190)
	$shf1zielinput = GUICtrlCreateInput("",250,186,130)
GUICtrlCreateLabel("Rufnummer für SHF 1", 10, 210)
	$shf1nummer = IniRead($configfilepath,"shfauswahl","shf1nummer","default")
	GUICtrlCreateLabel($shf1nummer, 150, 210)
	$shf1nummerinput = GUICtrlCreateInput("",250,206,130)
GUICtrlCreateLabel("Zielrufnummer für SHF 2", 10, 230)
	$shf2ziel = IniRead($configfilepath,"shfauswahl","shf2ziel","default")
	GUICtrlCreateLabel($shf2ziel, 150, 230)
	$shf2zielinput = GUICtrlCreateInput("",250,226,130)
GUICtrlCreateLabel("Rufnummer für SHF 2", 10, 250)
	$shf2nummer = IniRead($configfilepath,"shfauswahl","shf2nummer","default")
	GUICtrlCreateLabel($shf2nummer, 150, 250)
	$shf2nummerinput = GUICtrlCreateInput("",250,246,130)
GUICtrlCreateLabel("Zielrufnummer für SHF 3", 10, 270)
	$shf3ziel = IniRead($configfilepath,"shfauswahl","shf3ziel","default")
	GUICtrlCreateLabel($shf3ziel, 150, 270)
	$shf3zielinput = GUICtrlCreateInput("",250,266,130)
GUICtrlCreateLabel("Rufnummer für SHF 3", 10, 290)
	$shf3nummer = IniRead($configfilepath,"shfauswahl","shf3nummer","default")
	GUICtrlCreateLabel($shf3nummer, 150, 290)
	$shf3nummerinput = GUICtrlCreateInput("",250,286,130)
GUICtrlCreateLabel("Zielrufnummer für SHF 4", 10, 310)
	$shf4ziel = IniRead($configfilepath,"shfauswahl","shf4ziel","default")
	GUICtrlCreateLabel($shf4ziel, 150, 310)
	$shf4zielinput = GUICtrlCreateInput("",250,306,130)
GUICtrlCreateLabel("Rufnummer für SHF 4", 10, 330)
	$shf4nummer = IniRead($configfilepath,"shfauswahl","shf4nummer","default")
	GUICtrlCreateLabel($shf4nummer, 150, 330)
	$shf4nummerinput = GUICtrlCreateInput("",250,326,130)

    GUICtrlCreateLabel("Berechtigungen", 30, 350)

GUICtrlCreateLabel("Rufnummern in Excel", 10, 370)
	$excelzeilen = IniRead($configfilepath,"berechtigung","excelzeilen","default")
	GUICtrlCreateLabel($excelzeilen, 150, 370)
	$excelzeileninput = GUICtrlCreateInput("",250,366,130)
Local $idspeichern = GUICtrlCreateButton("Speichern", 220, 410, 85, 25)
Local $idabbrechen = GUICtrlCreateButton("Beenden", 310, 410, 85, 25)


    GUISetState(@SW_SHOW, $GUIkonfiguration)
    While 1
        Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE, $idabbrechen
			   GUIDelete($GUIkonfiguration)
			   WinActivate($hipathfenster)
			   WinWaitActive($hipathfenster)
			   ExitLoop

			Case $idspeichern
			   $iphicomneu = GUICtrlRead($iphicominput)
			   if $iphicomneu = 0 Then $iphicomneu = $iphicom
			   IniWrite($configfilepath, "netzwerk", "iphicom", $iphicomneu)
			   $subnethicomneu = GUICtrlRead($subnethicominput)
			   if $subnethicomneu = 0 Then $subnethicomneu = $subnethicom
			   IniWrite($configfilepath, "netzwerk", "subnethicom", $subnethicomneu)
			   $telnethicomipneu = GUICtrlRead($telnethicomipinput)
			   if $telnethicomipneu = 0 Then $telnethicomipneu = $telnethicomip
			   IniWrite($configfilepath, "netzwerk", "telnethicom", $telnethicomipneu)


			   $logoffdhcpneu = GUICtrlRead($DHCPCHECKBOX)
			   if $logoffdhcpneu = 1 Then
			   IniWrite($configfilepath, "netzwerk", "dhcp", "1")
			   Else
			   IniWrite($configfilepath, "netzwerk", "dhcp", "0")
			   EndIf

			   $ipseatelneu = GUICtrlRead($ipseatelinput)
			   if $ipseatelneu = 0 Then $ipseatelneu = $ipseatel
			   IniWrite($configfilepath, "netzwerk", "ipseatel", $ipseatelneu)
			   $subnetseatelneu = GUICtrlRead($subnetseatelinput)
			   if $subnetseatelneu = 0 Then $subnetseatelneu = $subnetseatel
			   IniWrite($configfilepath, "netzwerk", "subnetseatel", $subnetseatelneu)
			   $adapterneu1 = GUICtrlRead($adapterinput1)
			   if $adapterneu1 = "" Then $adapterneu1 = $adapter1
			   IniWrite($configfilepath, "netzwerk", "adapter1", $adapterneu1)
			   $adapterneu2 = GUICtrlRead($adapterinput2)
			   if $adapterneu2 = "" Then $adapterneu2 = $adapter2
			   IniWrite($configfilepath, "netzwerk", "adapter2", $adapterneu2)
			   $shf1zielneu = GUICtrlRead($shf1zielinput)
			   if $shf1zielneu = 0 Then $shf1zielneu = $shf1ziel
			   IniWrite($configfilepath, "shfauswahl", "shf1ziel", $shf1zielneu)
			   $shf1nummerneu = GUICtrlRead($shf1nummerinput)
			   if $shf1nummerneu = 0 Then $shf1nummerneu = $shf1nummer
			   IniWrite($configfilepath, "shfauswahl", "shf1nummer", $shf1nummerneu)
			   $shf2zielneu = GUICtrlRead($shf2zielinput)
			   if $shf2zielneu = 0 Then $shf2zielneu = $shf2ziel
			   IniWrite($configfilepath, "shfauswahl", "shf2ziel", $shf2zielneu)
			   $shf2nummerneu = GUICtrlRead($shf2nummerinput)
			   if $shf2nummerneu = 0 Then $shf2nummerneu = $shf2nummer
			   IniWrite($configfilepath, "shfauswahl", "shf2nummer", $shf2nummerneu)
			   $shf3zielneu = GUICtrlRead($shf3zielinput)
			   if $shf3zielneu = 0 Then $shf3zielneu = $shf3ziel
			   IniWrite($configfilepath, "shfauswahl", "shf3ziel", $shf3zielneu)
			   $shf3nummerneu = GUICtrlRead($shf3nummerinput)
			   if $shf3nummerneu = 0 Then $shf3nummerneu = $shf3nummer
			   IniWrite($configfilepath, "shfauswahl", "shf3nummer", $shf3nummerneu)
			   $shf4zielneu = GUICtrlRead($shf4zielinput)
			   if $shf4zielneu = 0 Then $shf4zielneu = $shf4ziel
			   IniWrite($configfilepath, "shfauswahl", "shf4ziel", $shf4zielneu)
			   $shf4nummerneu = GUICtrlRead($shf4nummerinput)
			   if $shf4nummerneu = 0 Then $shf4nummerneu = $shf4nummer
			   IniWrite($configfilepath, "shfauswahl", "shf4nummer", $shf4nummerneu)
			   $excelzeilenneu = GUICtrlRead($excelzeileninput)
			   if $excelzeilenneu = 0 Then $excelzeilenneu = $excelzeilen
			   IniWrite($configfilepath, "berechtigung", "excelzeilen", $excelzeilenneu)
			   GUIDelete($GUIkonfiguration)
			   MsgBox($MB_OK, "Erfolfreich", "Änderungen wurden gespeichert")
			   konfiguration()
			   ExitLoop
        EndSwitch
    WEnd
    GUIDelete($GUIkonfiguration)
	#WinActivate($hipathfenster)
    #WinWaitActive($hipathfenster)
 EndFunc
Func Backup()
if WinExists("*Unbenannt") Then $aktivesfenster = "*Unbenannt"
WinActivate($aktivesfenster)
WinWaitActive($aktivesfenster)
Send("EXEC-UPDAT:MODUL=BP,SUSY=ALL;")
Send("{Enter}")
Sleep(200)
WinActivate($hipathfenster)
WinWaitActive($hipathfenster)
EndFunc
Func Beenden()
$logoffadapter = $logonadapter
$logoffstatus = IniRead($configfilepath, "netzwerk", "dhcp", "default")
If $logoffstatus = 1 Then
$sCMDlogoff = "netsh interface ipv4 set address name="&$logoffadapter&" source=dhcp"
Run($sCMDlogoff)
ElseIf $logoffstatus = 0 Then
$logoffip = IniRead($configfilepath, "netzwerk", "ipseatel", "default")
$logoffsubnet = IniRead($configfilepath, "netzwerk", "subnetseatel", "default")
$sCMDlogoff = "netsh interface ipv4 set address "&$logoffadapter&" static "&$logoffip&" "&$logoffsubnet
Run($sCMDlogoff)
EndIf
If $aktivesfenster = $telnethicom &" - PuTTY" Then
ProcessClose($PIDCOMWIN)
ElseIf $aktivesfenster = "Unbenannt" or $aktivesfenster = "*Unbenannt" Then
ProcessClose($PIDNOTEPAD)
EndIf
_Excel_Close($oAppl)
Exit
EndFunc