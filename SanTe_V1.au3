; ************************************************************************************************
; SanTe_V1 (Sanquin Tecan)
;
; For making joblist.twl files for EVoware use from manually inpuutted samplenames 
; OR from scan.csv file in EVOware outputfolder
;
; by Dion Methorst 
; Version 0.1 March 2012
; Version 0.5 May 2013
; Version 1.0 SanTe_V1.exe 17-02-2014
;
; ************************************************************************************************
; TO DO list:
;
; ini file read ?
; TPLASC combi function ?
; FTP function ?
; 
; ************************************************************************************************
; start of script:

#include <Array.au3>
#include <file.au3>
#include <Date.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <GuiStatusBar.au3>
#include <WinAPI.au3>
#include <Constants.au3>
#include <String.au3>
#include <ListBoxConstants.au3>
#include <ButtonConstants.au3>


; Global input for datahandling
Global $protocol, $joblist, $LIMS, $Barcode
; Tecan assays in use
Global $aC1inh , $C1inhP1 , $C1inhP2 , $C1inhD , $IgA , $MBL , $Vwfag , $OPEN
; Dilution Factors
Global $Xfactor1, $Xfactor2, $Xfactor3, $Xfactor4

main()

; ************************************************************************************************
Func main()
   
;_Ini() ;Func initialise, setting the stage
   
	  Opt("GUICoordMode", 1)
	  Opt("GUIResizeMode", 1)
	  Opt("GUIOnEventMode", 1)

	  Local $icofile = "C:\apps\EVO\setup\Ontw Sanq.ico"
	  $font = "calibri"

Global const $SanTe = GuiCreate("SanTe for Joblist TPL & ASC files", 375, 400)
   ;DllCall("user32.dll", "int", "AnimateWindow", "hwnd", $Sante, "int", 1000, "long", 0x00080000) ; fade-in
   GuiSetIcon($icofile, 0)
   GUICtrlSetDefColor(0x0000AB) ;blue
   GUISetBkColor(0xEfFFFF)  ; will change background color
   GUICtrlSetFont(-1, 9, 600, 4, $font)
   GUISetFont(9, 550, 1, $font)
   GUISetOnEvent($GUI_EVENT_CLOSE, "SpecialEvents")
   GUISetOnEvent($GUI_EVENT_MINIMIZE, "SpecialEvents")
   GUISetOnEvent($GUI_EVENT_RESTORE, "SpecialEvents")

;Write Barcodes
   GUICtrlCreateGroup("Joblist Barcode",20,11,340,133)
   $Barcode = GUICtrlCreateEdit("scan barcodes" & @CRLF & "or use Joblist scan.csv",30,31,319,100)
   GUICtrlCreateGroup("",-99,-99,1,1)

; radiobutton menu
   GUICtrlCreateGroup("Assay",20,180,165,210)
   Global Const $Bep1		= GUICtrlCreateRadio("aC1inh  (WL788)",30,195,120,15)
   GUICtrlSetOnEvent(-1, "_Bep")
   Global Const $Bep2		= GUICtrlCreateRadio("C1inh Prdkt  (WL751)",30,220,120,15)
   GUICtrlSetOnEvent(-1, "_Bep")
   Global Const $Bep3	 	= GUICtrlCreateRadio("C1inh conc  (WL841)",30,245,120,15)
   GUICtrlSetOnEvent(-1, "_Bep")
   Global Const $Bep4	 	= GUICtrlCreateRadio("C1inh PAT  (WL689)",30,270,120,15)
   GUICtrlSetOnEvent(-1, "_Bep")
   Global Const $Bep5		= GUICtrlCreateRadio("IgA  (WL670)",30,295,120,15)
   GUICtrlSetOnEvent(-1, "_Bep")
   Global Const $Bep6		= GUICtrlCreateRadio("MBL  (WL675)",30,320,120,15)
   GUICtrlSetOnEvent(-1, "_Bep")
   Global Const $Bep7		= GUICtrlCreateRadio("VWfAg  (WIP)",30,345,120,15)
   GUICtrlSetOnEvent(-1, "_Bep")
   Global Const $Bep8		= GUICtrlCreateRadio("Free Offset  ",30,370,120,15)
   GUICtrlSetOnEvent(-1, "_Bep")
   ;GUICtrlSetState($Bep5, $GUI_CHECKED)
   GUICtrlCreateGroup("",-99,-99,1,1)
   
; V dilutions menu
   GUICtrlCreateGroup("Dilution",195,180,165,55)
   Global Const $V1 = GUICtrlCreateCheckbox("V1",205,205,30,15)
   GUICtrlSetOnEvent(-1, "_V1")
   Global Const $V2	= GUICtrlCreateCheckbox("V2",245,205,30,15)
   GUICtrlSetOnEvent(-1, "_V2")
   Global Const $V3 = GUICtrlCreateCheckbox("V3",285,205,30,15)
   GUICtrlSetOnEvent(-1, "_V3")
   Global Const $V4	= GUICtrlCreateCheckbox("V4",325,205,30,15)
   GUICtrlSetOnEvent(-1, "_V4")
   GUICtrlCreateGroup("",-99,-99,1,1)

; MenuButtons
   $MakeJoblist = GUICtrlCreateButton("Joblist Barcode",195,245,165,25)
   GUICtrlSetOnEvent(-1, "WriteBC")
   $ReadBarcode = GUICtrlCreateButton("Joblist scan.csv",195,275,165,25)
   GUICtrlSetOnEvent(-1, "Make_Bjob")
   ;$TPLASC = GUICtrlCreateButton("TPL && ASC",195,305,165,25)
   ; GUICtrlSetOnEvent(-1, "_TPLASC")
   ;$HelpMenu = GUICtrlCreateButton("Need Help?",195,335,165,25)
   ;GUICtrlSetOnEvent(-1, "Help_Call")
   $ExitButton = GUICtrlCreateButton("Exit",195,305,165,25)
   GUICtrlSetOnEvent(-1, "Xsit")
	;$protocol = GUICtrlCreateInput("",30,370,140,20)
	;$SaveButton = GUICtrlCreateButton("Save Settings",195,285,165,25)

	GUISetState(@SW_SHOW)

   While 1
        Sleep(10)
   WEnd

EndFunc ; EndFunc Main()
;-------------------------------------------------------------------------------------------------
; Func _TPLASC()

Func _Ini()
; work in progress.... just an idea
  ; Local $icofile = IniRead("C:\apps\EVO\setup\SanTe.ini", "GUI", "icon", "NotFound")
   

Endfunc
;-------------------------------------------------------------------------------------------------

Func _Bep()

		Select
		case BitAnd(GUICtrlRead($Bep1),$GUI_CHECKED);GuiCtrlRead($Bep1) = $GUI_CHECKED
			$LIMS = "25"
;			case GuiCtrlRead($Bep1b) = $GUI_CHECKED
;			$LIMS = "26"
;			case GuiCtrlRead($Bep1c) = $GUI_CHECKED
;			$LIMS = "27"
		case BitAnd(GUICtrlRead($Bep2),$GUI_CHECKED);GuiCtrlRead($Bep2) = $GUI_CHECKED
			$LIMS = "88"
		case BitAnd(GUICtrlRead($Bep3),$GUI_CHECKED);GuiCtrlRead($Bep3) = $GUI_CHECKED
			$LIMS = "92"
		case BitAnd(GUICtrlRead($Bep4),$GUI_CHECKED);GuiCtrlRead($Bep4) = $GUI_CHECKED
			$LIMS = "74"
		case BitAnd(GUICtrlRead($Bep5),$GUI_CHECKED);GuiCtrlRead($Bep5) = $GUI_CHECKED
			$LIMS = "11"
		case BitAnd(GUICtrlRead($Bep6),$GUI_CHECKED);GuiCtrlRead($Bep6) = $GUI_CHECKED
			$LIMS = "12"
		case BitAnd(GUICtrlRead($Bep7),$GUI_CHECKED);GuiCtrlRead($Bep7) = $GUI_CHECKED
			$LIMS = "103"
		Case BitAnd(GUICtrlRead($Bep8),$GUI_CHECKED);GuiCtrlRead($Bep8) = $GUI_CHECKED
			$LIMS = inputbox("Tecan LIMS-number input", "Enter your Tecan EVO LIMS-number:" , "99","","400","220") ;_
			   ;& @CRLF & @CRLF & "IgA = 11, MBL = 12, C1inh = .." _
			   ;& @CRLF & @CRLF & "aC1inh = 25, aBSA = 26, CETOR= 27" _
			   ;& @CRLF & @CRLF & "of gebruik <CTRL H>, zoek en vervang functie", "99","","400","220")

	EndSelect

EndFunc ;Assay radiobutton

Func _V1() ;(Const ByRef $V1)
	If GUICtrlRead($V1) = $GUI_CHECKED Then
       $Xfactor1 = "V1"
    ElseIf GUICtrlRead($V1) = $GUI_UNCHECKED Then
        $Xfactor1 = ""
    EndIf
 EndFunc ;dilution checkboxes
 
Func _V2();(Const ByRef $V2)
    If GUICtrlRead($V2) = $GUI_CHECKED Then
       $Xfactor2 = "V2"
    ElseIf GUICtrlRead($V2) = $GUI_UNCHECKED Then
        $Xfactor2 = ""
    EndIf
 EndFunc ;dilution checkboxes
 
Func _V3();(Const ByRef $V3)
    If GUICtrlRead($V3) = $GUI_CHECKED Then
       $Xfactor3 = "V3"
    ElseIf GUICtrlRead($V3) = $GUI_UNCHECKED Then
        $Xfactor3 = ""
    EndIf
 EndFunc ;dilution checkboxes
 
Func _V4();(Const ByRef $V4)
    If GUICtrlRead($V4) = $GUI_CHECKED Then
       $Xfactor4 = "V4"
    ElseIf GUICtrlRead($V4) = $GUI_UNCHECKED Then
        $Xfactor4 = ""
    EndIf
EndFunc ;dilution checkboxes

Func  WriteBC()    
	  
;MsgBox(0, "Barcode input", GUICtrlRead($Barcode)) ; display the selected listbox entry
local $FolderPath = "C:\apps\EVO\CSV\"
;local $FolderPath = "C:\Program Files\Tecan\EVOware\output\"
local $FileNam2 = "SanTeBC.csv"
local $BCcsv = $FolderPath & $FileNam2
Local $FileOpen = FileOpen($BCcsv, 2) ; write mode 2 = overwrite, 1 = append
FileWrite($FileOpen, GUICtrlRead($Barcode))
	  
If FileExists($BCcsv) Then
   local $JOB2 = FileOpen("C:\apps\EVO\JOB\" & "joblist.twl", 2 + 8)
Else
MsgBox(4096, $BCcsv, "Scan.csv does NOT exist in folder:" & @CRLF & @CRLF & "C:\apps\EVO\CSV")
EndIf

Local $CountLines2 = _FileCountLines($BCcsv)
	  ;MsgBox(4096, "aantal regels =", $CountLines2, 2)

local $BCarray[$countlines2]
	  _FileReadToArray($BCcsv, $BCarray)
	  ;_ArrayDisplay($BCarray)

FileWriteLine($JOB2, "H;" & StringReplace(_NowCalcDate(), "/", "-") & ";" & _DateTimeFormat(_NowCalc(), 5))

For $Writ2 = 1 To Ubound($BCarray)-1 ;Ubound($Names)-2

			   FileWriteline($JOB2, "P;" & $BCarray[$Writ2] & @CRLF)
				If	$Xfactor1 = "V1" 	Then FileWriteline($JOB2, "O;" & $LIMS & ";1;" & $Xfactor1 & ";1;1.0"  & @CRLF)
				If	$Xfactor2 = "V2" 	Then FileWriteline($JOB2, "O;" & $LIMS & ";1;" & $Xfactor2 & ";1;1.0"  & @CRLF)
				If	$Xfactor3 = "V3" 	Then FileWriteline($JOB2, "O;" & $LIMS & ";1;" & $Xfactor3 & ";1;1.0"  & @CRLF)
				If	$Xfactor4 = "V4" 	Then FileWriteline($JOB2, "O;" & $LIMS & ";1;" & $Xfactor4 & ";1;1.0"  & @CRLF)

Next 

FileWriteLine($JOB2, "L;" & @CRLF)

fileclose($JOB2)
Fileclose($BCcsv)
MsgBox(4096, "Joblist", "Joblist.twl in" & @CRLF & "C:\APPS\EVO\JOB\", 3)
	  
EndFunc ;joblist from scanned barcodes

Func Make_Bjob()

;local $LIMS = inputbox("Tecan LIMS-nummer invoeren", "voer het Tecan EVO LIMS-nummer in:" _
;			   & @CRLF & @CRLF & "IgA = 11, MBL = 12, C1inh = .." _
;			   & @CRLF & @CRLF & "aC1inh = 25, aBSA = 26, CETOR= 27" _
;			   & @CRLF & @CRLF & "of gebruik <CTRL H>, zoek en vervang functie", "99","","400","220")

;local $FolderPath = "C:\Program Files\Tecan\EVOware\output\"
local $FolderPath = "C:\APPS\EVO\CSV\"
local $FileName = "scan.csv"
local $var = $FolderPath & $FileName
Local $CountLines = _FileCountLines($var)

If FileExists($var) Then
	  local $JOB = FileOpen("C:\apps\EVO\JOB\" & "joblist.twl", 2 + 8)
	  Else
	  MsgBox(4096, $var, "Scan.csv does NOT exist in folder:" & @CRLF & @CRLF & "C:\apps\EVO\CSV")
	  EndIf

local $Names = _ParseCSV($var, ";", '$', 4)

for $Trim = 0 to $Countlines-1

	$position = StringInStr($Names[$Trim][0], ";", 0, -1)
	$Names[$Trim][0] = StringTrimleft ( $Names[$Trim][0], $position )

next

FileWriteLine($JOB, "H;" & StringReplace(_NowCalcDate(), "/", "-") & ";" & _DateTimeFormat(_NowCalc(), 5))

For $Write = 1 To Ubound($Names)-2

			Select
			Case Stringright($Names[$Write][0], 3) <> "$$$"
			   FileWriteline($JOB, "P;" & $Names[$Write][0] & @CRLF)
				If	$Xfactor1 = "V1" 	Then FileWriteline($JOB, "O;" & $LIMS & ";1;" & $Xfactor1 & ";1;1.0"  & @CRLF)
				If	$Xfactor2 = "V2" 	Then FileWriteline($JOB, "O;" & $LIMS & ";1;" & $Xfactor2 & ";1;1.0"  & @CRLF)
				If	$Xfactor3 = "V3" 	Then FileWriteline($JOB, "O;" & $LIMS & ";1;" & $Xfactor3 & ";1;1.0"  & @CRLF)
				If	$Xfactor4 = "V4" 	Then FileWriteline($JOB, "O;" & $LIMS & ";1;" & $Xfactor4 & ";1;1.0"  & @CRLF)

			EndSelect
Next 

FileWriteLine($JOB, "L;" & @CRLF)

fileclose($JOB)
FileClose($var)

MsgBox(4096, "Joblist", "Joblist.twl in" & @CRLF & "C:\APPS\EVO\JOB\", 3)
$Countlines = ""
$var = ""
$filename = ""
$folderpath = ""
$filenam2 = ""
$JOB = ""
$Names 	= ""
$position = ""
$Write =""
$Trim = ""

Endfunc ; joblist from scan.csv file

Func Help_Call()

	Run('hh mk:@MSITStore:'&StringReplace(@AutoItExe,'.exe','.chm')&'::/html/tutorials/helloworld/helloworld.htm','',@SW_MAXIMIZE)

EndFunc ;call for HELP! :)

Func SpecialEvents()


    Select
        Case @GUI_CtrlId = $GUI_EVENT_CLOSE
            ;MsgBox(0, "Close Pressed", "ID=" & @GUI_CtrlId & " WinHandle=" & @GUI_WinHandle)
			;DllCall("user32.dll", "int", "AnimateWindow", "hwnd", $SanTe, "int", 1000, "long", 0x00090000)
            Exit

        Case @GUI_CtrlId = $GUI_EVENT_MINIMIZE
            ;MsgBox(0, "Window Minimized", "ID=" & @GUI_CtrlId & " WinHandle=" & @GUI_WinHandle)

        Case @GUI_CtrlId = $GUI_EVENT_RESTORE
            ;MsgBox(0, "Window Restored", "ID=" & @GUI_CtrlId & " WinHandle=" & @GUI_WinHandle)

    EndSelect

EndFunc

Func Xsit()
	;DllCall("user32.dll", "int", "AnimateWindow", "hwnd", $SanTe, "int", 1000, "long", 0x00090000)
	exit

EndFunc

Func _ParseCSV($sFile, $sDelimiters=';', $sQuote='"', $iFormat=0)
	Local Static $aEncoding[6] = [0, 0, 32, 64, 128, 256]
	If $iFormat < -1 Or $iFormat > 6 Then
		Return SetError(3,0,0)
	ElseIf $iFormat > -1 Then
		Local $hFile = FileOpen($sFile, $aEncoding[$iFormat]), $sLine, $aTemp, $aCSV[1], $iReserved, $iCount
		If @error Then Return SetError(1,@error,0)
		$sFile = FileRead($hFile)
		FileClose($hFile)
	EndIf
	If $sDelimiters = "" Or IsKeyword($sDelimiters) Then $sDelimiters = ';'
	If $sQuote = "" Or IsKeyword($sQuote) Then $sQuote = '"'
	$sQuote = StringLeft($sQuote, 1)
	Local $srDelimiters = StringRegExpReplace($sDelimiters, '[\\\^\-\[\]]', '\\\0')
	Local $srQuote = StringRegExpReplace($sQuote, '[\\\^\-\[\]]', '\\\0')
	Local $sPattern = StringReplace(StringReplace('(?m)(?:^|[,])\h*(["](?:[^"]|["]{2})*["]|[^,\r\n]*)(\v+)?',';', $srDelimiters, 0, 1),'"', $srQuote, 0, 1)
	Local $aREgex = StringRegExp($sFile, $sPattern, 3)
	If @error Then Return SetError(2,@error,0)
	$sFile = '' ; save memory
	Local $iBound = UBound($aREgex), $iIndex=0, $iSubBound = 1, $iSub = 0
	Local $aResult[$iBound][$iSubBound]
	For $i = 0 To $iBound-1
		Select
			Case StringLen($aREgex[$i])<3 And StringInStr(@CRLF, $aREgex[$i])
				$iIndex += 1
				$iSub = 0
				ContinueLoop
			Case StringLeft(StringStripWS($aREgex[$i], 1),1)=$sQuote
				$aREgex[$i] = StringStripWS($aREgex[$i], 3)
				$aResult[$iIndex][$iSub] = StringReplace(StringMid($aREgex[$i], 2, StringLen($aREgex[$i])-2), $sQuote&$sQuote, $sQuote, 0, 1)
			Case Else
				$aResult[$iIndex][$iSub] = $aREgex[$i]
		EndSelect
		$aREgex[$i]=0 ; save memory
		$iSub += 1
		If $iSub = $iSubBound Then
			$iSubBound += 1
			ReDim $aResult[$iBound][$iSubBound]
		EndIf
	Next
	If $iIndex = 0 Then $iIndex=1
	ReDim $aResult[$iIndex][$iSubBound]
	Return $aResult
 EndFunc ; End function Parse CSV
 
Func Bool(Const ByRef $checkbox)
    If GUICtrlRead($checkbox) = $GUI_CHECKED Then
        Return 1
    ElseIf GUICtrlRead($checkbox) = $GUI_UNCHECKED Then
        Return 0
    EndIf
EndFunc







