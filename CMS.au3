#AutoIt3Wrapper_icon=CMS.ico

; -------- Description --------
; CMS system on the interviewer PC.
; by  Richard Bruna - UIV 2.0.9.0


; -------- Includes --------


#Include <File.au3>                                                                             ; Needed for file operations.
#include <Array.au3>                                                                            ; Needed for array functions.
#include <GUIConstantsEx.au3>                                                                   ; Needed for GUI.
#include <WindowsConstants.au3>                                                                 ; Needed for GUI.
#Include <GuiTreeView.au3>                                                                      ; Needed for GUI.
#Include <GuiListView.au3>                                                                      ; Needed for GUI.
#include <GuiButton.au3>                                                                        ; Needed for GUI.
#include <Date.au3>                                                                             ; Needed for date functions.
#Include <ScreenCapture.au3>										   							; For screenshot
#Include <FTPEx.au3>																			; FTP support.
#include <ComboConstants.au3>																	; Combo box support.


; -------- Init variables --------


const $drive = StringLeft(@ScriptDir, 1)                                                        ; Get the drive letter the script is running from.
const $piaacpath = @scriptdir & "\"                                                             ; Set piaac dir.
const $vmexe ='"' & "c:\progra~1\VMware\vmware vix\vmrun.exe" & '"' & " -T player -gu root -gp poffpeng "		; Prefix to VM-commands.
const $vm =  $piaacpath & "vm\current\CZ.vmx"					 								; Full path to VM.
const $vmfilename  = "CZ.vmx"																	; VM file name
const $vmpref = @UserProfileDir & "\dataap~1\vmware\preferences.ini"
const $vm_current =  $piaacpath &  "vm\current\"												; Path to the current VM.
const $vm_backup = $piaacpath & "vm\backup\"                                                    ; Path the current damaged VM will be moved to.
const $vm_restore = $piaacpath & "vm\restore\"                                            		; Path to the working backup directory
const $random_file = $piaacpath & "random.txt"													; random record number file
const $oggdir = "c:\piaac\record\"																; recorder directory
const $ogg = "harddisk.exe"																		; Command for recording.
const $oggpreset = "recording.hdp"																; Preset settings for HarddiskOgg.
const $oggfilter = "normalization.hfs"															; Filter settings for HarddiskOgg.
const $cmsfilename = FileFindNextFile(FileFindFirstFile($piaacpath & "*.csv"))					; CSV master file name
const $cmsfile = $piaacpath & $cmsfilename														; CSV master file path
const $tcarray = StringSplit($cmsfilename,".",2)												; CAPI number array
const $tc = $tcarray[0]																			; CAPI number
const $inputpath = $piaacpath & "input\"														; XML ZIP import path.
const $outputpath = $piaacpath & "output\"														; XML ZIP output path.
const $ftpIP = '[removed]'																	; FTP IP address
const $ftpuser = '[removed]'																		; FTP username
const $ftppass = '[removed]'																		; FTP password
const $ftproot = "/prichozi/"																	; FTP root
$sortBy = 2                                                                                     ; Default column to sort by status
$showRetur = 0                                                                                  ; Default show finished on/off

; Column names and index in casefile. They all start with i.
const $iPersID = 0
const $iStatus = 1
const $iAssignDate = 2
const $iNextDateDay = 3
const $iNextDateMonth = 4
const $iNextDateYear = 5
const $iNextTime = 6
const $iSecondName = 7
const $iName = 8
const $iAge = 9
const $iYear = 10
const $iPhone = 11
const $iPhoneType = 12
const $iAddressStreet = 13
const $iAddressCp = 14
const $iCity = 15
const $iPeople = 16
const $iOrder = 17
const $iLastNote = 18
const $iLastDc = 19
const $iVisit1nextdate = 20
const $iVisit1nexttime = 21
const $iVisit1date = 22
const $iVisit1time = 23
const $iVisit1note = 24
const $iVisit1code = 25
const $iVisit1type = 26
const $iVisit2nextdate = 27
const $iVisit2nexttime = 28
const $iVisit2date = 29
const $iVisit2time = 30
const $iVisit2note = 31
const $iVisit2code = 32
const $iVisit2type = 33
const $iVisit3nextdate = 34
const $iVisit3nexttime = 35
const $iVisit3date = 36
const $iVisit3time = 37
const $iVisit3note = 38
const $iVisit3code = 39
const $iVisit3type = 40
const $iVisit4nextdate = 41
const $iVisit4nexttime = 42
const $iVisit4date = 43
const $iVisit4time = 44
const $iVisit4note = 45
const $iVisit4code = 46
const $iVisit4type = 47
const $iVisit5nextdate = 48
const $iVisit5nexttime = 49
const $iVisit5date = 50
const $iVisit5time = 51
const $iVisit5note = 52
const $iVisit5code = 53
const $iVisit5type = 54
const $iVisit6nextdate = 55
const $iVisit6nexttime = 56
const $iVisit6date = 57
const $iVisit6time = 58
const $iVisit6note = 59
const $iVisit6code = 60
const $iVisit6type = 61
const $iVisit7nextdate = 62
const $iVisit7nexttime = 63
const $iVisit7date = 64
const $iVisit7time = 65
const $iVisit7note = 66
const $iVisit7code = 67
const $iVisit7type = 68
const $iVisit8nextdate = 69
const $iVisit8nexttime = 70
const $iVisit8date = 71
const $iVisit8time = 72
const $iVisit8note = 73
const $iVisit8code = 74
const $iVisit8type = 75
const $iVisit9nextdate = 76
const $iVisit9nexttime = 77
const $iVisit9date = 78
const $iVisit9time = 79
const $iVisit9note = 80
const $iVisit9code = 81
const $iVisit9type = 82
const $iVisit10nextdate = 83
const $iVisit10nexttime = 84
const $iVisit10date = 85
const $iVisit10time = 86
const $iVisit10note = 87
const $iVisit10code = 88
const $iVisit10type = 89

; Statuses for $iStatus.
const $cStatus_new = 0
const $cStatus_paused = 9
const $cStatus_intervju = 1
const $cStatus_overfor = 2


; -------- Control --------


; Only allow once instance running.
$process = ProcessList(@ScriptName)
if(ubound($process) > 2) Then
	MsgBox(8192, "SC&C - PIAAC", "Aplikace CMS byla spuštìna.")
	Exit
EndIf

;master file control
if $cmsfilename = "" Then
	MsgBox(8192, "Soubor nenalezen", "Hlavní CSV soubor nebyl nalezen.")
	exit
endIf
if not StringRegExp($cmsfilename,'^(\d{1,6})(\.csv)$',0) then
	MsgBox(8192, "Nesprávný název souboru", "Nesprávný název hlavního CSV souboru.")
	exit
endIf

;check system time flashed by BIOS
local $sysTime = _Date_Time_GetLocalTime()
local $sysTimeArray = _Date_Time_SystemTimeToArray($sysTime)
if not ($sysTimeArray[2] > 2010) then
		MsgBox(8192,"SC&C - PIAAC", "Poèítaè má nesprávné datum.")
		Exit
EndIf

; patch the vmplayer configuration
patch()
if RegRead("HKEY_CURRENT_USER\Control Panel\Desktop","LowLevelHooksTimeout") <> "5000" then
	RegWrite("HKEY_CURRENT_USER\Control Panel\Desktop","LowLevelHooksTimeout","REG_DWORD","5000")
EndIf

;Check VM Player installation
if not FileExists(@ProgramFilesDir & "\VMware\vmware vix\vmrun.exe") then
	MsgBox(8192, "SC&C - PIAAC", "Nelze nalézt cestu k programu VM Player.")
	Exit
endif

;Check VM Image instaltion control
if not FileExists($vm) or not FileExists("c:\piaac\vm\current\CZ.nvram") or not FileExists("c:\piaac\vm\current\CZ.vmxf") then
	MsgBox(8192, "SC&C - PIAAC", "Nelze nalézt konfiguraèní soubory obrazu dotazníku.")
	exit
endif

;Check VM Image backup instaltion control
;if not FileExists($vm_restore & $vmfilename) then
;	MsgBox(8192, "SC&C - PIAAC", "Nelze nalézt záložní obraz dotazníku.")
;	exit
;endif

;Master file format check
if FileGetSize($cmsfile) <> 0 then
	local $fopen = FileOpen($cmsfile), $csvline = 0
	while 1
		$csvstring = FileReadLine($fopen)
		$csvline += 1
		if (@error = -1 or $csvstring = "") then exitLoop
		;separator control
		if StringInStr($csvstring,";",0,89) = 0 or StringInStr($csvstring,";",0,90) > 0 Then
			MsgBox(8192, "SC&C - PIAAC", "Nesprávný formát hlavního souboru. (" & $csvline & ")")
			exit
		endIf
	wend
	FileClose($fopen)
EndIf

;DC 04 patch
$cases = getAllCases($cmsfile)
for $p = 0 to UBound($cases) - 1
	if $cases[$p][19] = "4" then $cases[$p][19] = "04"
next
writeAllCases()


; -------- Main --------


;predefined global variables for compliation fix
global $cases,$casesView

global $winMain = GUICreate("SC&C - PIAAC 2.0.9.0", 830, 480,-1,-1,$WS_CAPTION+$WS_SYSMENU+$WS_MAXIMIZEBOX+$WS_MINIMIZEBOX+$WS_SIZEBOX,$WS_SIZEBOX)   ; Create main window.
global $btnStart = GUICtrlCreateButton("Start",95,440,80)
global $btnSort = GUICtrlCreateButton("Seøadit",187,440,80)
global $btnRetur = GUICtrlCreateButton("Filtr",279,440,80)
global $btnScreen = GUICtrlCreateButton("Screener",371,440,80)
global $btnVisitpad = GUICtrlCreateButton("Návštìvy",463,440,80)
global $btnExport = GUICtrlCreateButton("Export",555,440,80)
global $btnPatch = GUICtrlCreateButton("Záplata",647,440,80)
global $btnClose = GUICtrlCreateButton("Ukonèit",739,440,80)
global $listview = GUICtrlCreateListView("Pers.ID              |Status|Disp.k.|Datum|Èas|Pøíjmení     |Jméno|Vìk|Rok|Tel. èíslo|Ulice|Mìsto",10,10,810,420, $LVS_SHOWSELALWAYS, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT))
GUICtrlSetFont(-1,10)

GUISetState(@SW_SHOW, $winMain)
;set help
GUISetHelp("C:\WINDOWS\hh.exe " & $piaacpath & "cms.chm")

init()

While 1
	$persID = getPersID()
	switch GUIGetMsg()
		case $GUI_EVENT_CLOSE, $btnClose
			exit
		case $btnStart
			if($persID = "") Then
				msgbox(8192, "Žádný respondent nebyl vybrán", "Musíte nejprve vybrat respondenta.")
			elseif($cases[getCaseIndex($persID)][$iStatus] = $cStatus_intervju) Then
				MsgBox(8192, "Dokonèený dotazník", "Dotazník èíslo " & $PersID & " byl dokonèen.")
			else
				;get the name
				if $cases[getCaseIndex($persID)][$iName] <> "" Then
					$fullname = $cases[getCaseIndex($persID)][$iSecondName] & " " &  $cases[getCaseIndex($persID)][$iName]
				else
					$fullname = $cases[getCaseIndex($persID)][$iSecondName]
				endIf
				;confirmation
				if(MsgBox(8195, "Potvrdit dotazník", "Spustit dotazník èíslo " & $persID & " - " & $fullname & " ?") = 6) then
					cmsvisibility(0)
					;enable rescue mode
					HotKeySet("+!r", "escape")
					; disable repair mode
					HotKeySet("+!f")
 					recordingStart($persID)
					startCAPI($persID)
					waitForFinish()
					writeAllCases()
					sortCases($sortBy)
					recordingStop()
					;disable rescue mode
					HotKeySet("+!r")
					;enable repair mode
					HotKeySet("+!f", "repairVM")
					cmsvisibility(1)
					;make finished persid selected
					_GUICtrlListView_SetItemSelected($listview, getCaseSortIndex($persID))
				endif
			endif
		case $btnRetur
			if($showRetur = 0) then
				$showRetur = 1
				GUICtrlSetBkColor($btnRetur,0x6ccf56)
			elseif($showRetur = 1) Then
				$showRetur = 0
				GUICtrlSetBkColor($btnRetur,0xd4d0c8)
			endif
			; Fill list with cases
			fillList()
		case $btnSort
			sort()
		case $btnScreen
			if($persID = "") Then
				msgbox(8192, "Žádný respondent nebyl vybrán", "Musíte nejprve vybrat respondenta.")
			Else
				screener()
			EndIf
		case $btnVisitpad
			if($persID = "") Then
				msgbox(8192, "Žádný respondent nebyl vybrán", "Musíte nejprve vybrat respondenta.")
			Else
				visitpad()
			EndIf
		case $btnExport
			if(MsgBox(8195, "Potvrdit aktulizaci", "Opravdu chcete spustit import a export ?") = 6) then
				importexport()
			endif
		case $btnPatch
			if(MsgBox(8195, "Potvrdit záplatu", "Opravdu chcete instalovat záplatu ?") = 6) then
				taopatch()
			endif
	endswitch
WEnd
GUIDelete($winMain)


; -------- Functions --------


; Initialize the CMS.
func init()
	; Get all cases from cases-file.
	global $cases = getAllCases($cmsfile)
	; Used for sorting and viewing only.
	global $casesView = $cases
	;set repair hotkey [l-shft] + [r-alt] + [f]
	HotKeySet("+!f", "repairVM")
	;retrieve recording random file
	get_rnd_file()
	;IMPORT
	if FileExists($inputpath & $tc & ".csv") then
		if FileGetSize($inputpath & $tc & ".csv") <> 0 then
			local $importfailed = 0
			local $fopen = FileOpen($inputpath & $tc & ".csv")
			;import format control
			while 1
				$csvstring = FileReadLine($fopen)
				if(@error = -1 or $csvstring = "") then ExitLoop
				if StringInStr($csvstring,";",0,89) = 0 or StringInStr($csvstring,";",0,90) > 0 Then
					MsgBox(8192, "SC&C - PIAAC", "Nesprávný formát importovaného souboru.")
					$importfailed = 1
				endif
			wend
			FileClose($fopen)
			;control passed
			if $importfailed = 0 then
				;open master file for append to prevent new line bug
				$append = FileOpen($cmsfile,1)
				;import
				local $line,$var,$persid,$csvarray
				_FileReadToArray($inputpath & $tc & ".csv",$csvarray)
				;strip first line
				for $g = 2 to UBound($csvarray) - 1
					$var = StringSplit($csvarray[$g], ";", 2)
					$persID = $var[0]
					$case_exist = _ArrayFindAll($cases, $persID, 0, 0, 0, 0, $iPersID)
					if(not(UBound($case_exist) > 0)) then
						FileWriteLine($append, $csvarray[$g])
					endif
				next
			endif
			FileClose($append)
		endif
		;clean after self
		FileDelete($inputpath & $tc & ".csv")
	endif
	;get cases array again becase of previous import
	$cases = getAllCases($cmsfile)
	;generate XML after update
	XMLgenerate()
	;make the initial sort
	sortCases($sortBy)
endfunc

;Sort screen
func sort()
	cmsvisibility(0)

	$winSort = GUICreate("Výbìr parametru",240,194, -1, -1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU)
	GUICtrlCreateLabel("Vyberte zvolený sloupec:",10,10)
	$btnSortOk = GUICtrlCreateButton("OK",58,159,80)
	$btnSortClose = GUICtrlCreateButton("Zrušit",150,159,80)
	GUICtrlCreateGroup("",10,25,220,124)
	global $sortRadio1 = GUICtrlCreateRadio ("Pers.ID ",20,40)
	global $sortRadio2 = GUICtrlCreateRadio ("Status ",20,60)
	global $sortRadio3 = GUICtrlCreateRadio ("Dispozièní kód ",20,80)
	global $sortRadio4 = GUICtrlCreateRadio ("Datum / Èas ",20,100)
	global $sortRadio5 = GUICtrlCreateRadio ("Ulice / Mìsto ",20,120)

	switch $sortBy
		case 1
			GUICtrlSetState($sortRadio1, $GUI_CHECKED)
		case 2
			GUICtrlSetState($sortRadio2, $GUI_CHECKED)
		case 3
			GUICtrlSetState($sortRadio3, $GUI_CHECKED)
		case 4
			GUICtrlSetState($sortRadio4, $GUI_CHECKED)
		case 5
			GUICtrlSetState($sortRadio5, $GUI_CHECKED)
	endswitch

	GUISetState(@SW_SHOW, $winSort)

	While 1
		$msg = GUIGetMsg()
		if $msg = $GUI_EVENT_CLOSE or $msg = $btnSortClose then exitloop
		if $msg = $btnSortOk then
			if(bitand(GUICtrlRead($sortRadio1), $GUI_CHECKED)) then $sortBy = 1
			if(bitand(GUICtrlRead($sortRadio2), $GUI_CHECKED)) then $sortBy = 2
			if(bitand(GUICtrlRead($sortRadio3), $GUI_CHECKED)) then $sortBy = 3
			if(bitand(GUICtrlRead($sortRadio4), $GUI_CHECKED)) then $sortBy = 4
			if(bitand(GUICtrlRead($sortRadio5), $GUI_CHECKED)) then $sortBy = 5
			sortCases($sortBy)
			exitloop
		endif
	Wend
	GUIDelete($winSort)
	cmsvisibility(1)
Endfunc

; Get all cases in the cases-file.
func getAllCases($array)
	local $numlines = _FileCountLines($array)
	local $numvars = UBound(StringSplit(FileReadLine($array), ";", 2))

	if($numlines = 0 or $numvars = 0) then
		local $cases[1][1]
		return $cases
	endif
	local $cases[$numlines][$numvars]
	local $i = 0
	local $j = 0
	local $file = FileOpen($array, 0)
	local $line, $vars
	while 1
		$line = filereadline($file)
		if(@error = -1 or $line = "") then ExitLoop
		$vars = StringSplit($line, ";", 2)
		for $var in $vars
			if($j = 0) then
				$cases[$i][$j] = int($var)
			else
				$cases[$i][$j] = $var
			endif
			$j += 1
		next
		$j = 0
		$i += 1
	wend
	FileClose($file)
	return $cases
endfunc


; Write all cases the cases file.
func writeAllCases()
	local $file = FileOpen($cmsfile, 2)
	for $i = 0 to UBound($cases) - 1
		$line = ""
		for $j = 0 to UBound($cases, 2) - 1
			$line &= $cases[$i][$j]
			if($j < UBound($cases, 2) - 1) then
				$line &= ";"
			endif
		next
		FileWriteLine($file, $line)
	next
	FileClose($file)
endfunc


; Fill list with cases.
func fillList()
	_GUICtrlListView_DeleteAllItems($listview)
	if($casesView[0][0] = "") then return

	for $i = 0 to UBound($casesView) - 1
		switch $casesView[$i][$iStatus]
			Case 0
				$status = "Prázdný"
			case 9
				$status = "Nedokonèený"
			case 1
				$status = "Rozhovor"
			case 2
				$status = "Chyba"
;			case 3
;				$status = "Odmítnuto"
;			case 4
;				$status = "Odesláno"
		EndSwitch

		switch $casesView[$i][$iLastDc]
			Case 01
				$lastdc = "1"
			case 03
				$lastdc = "3"
			case 04
				$lastdc = "4"
			case 05
				$lastdc = "5"
			case 07
				$lastdc = "7"
			case 08
				$lastdc = "8"
			case 09
				$lastdc = "9"
			Case else
				$lastdc = $casesView[$i][$iLastDc]
		EndSwitch

		;join appointment date
		if $casesView[$i][$iNextDateDay] <> "" then
			$nextdate = $casesView[$i][$iNextDateDay] & "." & $casesView[$i][$iNextDateMonth] & "." & $casesView[$i][$iNextDateYear]
		else
			$nextdate = ""
		endif
		;join full address
		if  $casesView[$i][$iAddressStreet] <> "" then
			$address = $casesView[$i][$iAddressStreet] & " " & $casesView[$i][$iAddressCp]
		else
			$address = ""
		endif

		$str = ""
		$str &= $casesView[$i][$iPersID] & "|"
		$str &= $status & "|"
		$str &= $lastdc & "|"
		$str &= $nextdate & "|"
		$str &= $casesView[$i][$iNextTime] & "|"
		$str &= $casesView[$i][$iSecondName] & "|"
		$str &= $casesView[$i][$iName] & "|"
		$str &= $casesView[$i][$iAge] & "|"
		$str &= $casesView[$i][$iyear] & "|"
		$str &= $casesView[$i][$iPhone] & "|"
		$str &= $address & "|"
		$str &= $casesView[$i][$iCity] & "|"

		if($showRetur = 0) Then
			if($casesView[$i][$iStatus] = $cStatus_new or $casesView[$i][$iStatus] = $cStatus_paused) Then
				GUICtrlCreateListViewItem($str, $listview)
				GUICtrlSetBkColor($btnRetur,0xd4d0c8)
			endif
		elseif($showRetur = 1) then
			if($casesView[$i][$iStatus] = $cStatus_intervju or $casesView[$i][$iStatus] = $cStatus_overfor) then
				GUICtrlCreateListViewItem($str, $listview)
			;	GUICtrlSetBkColor($btnRetur,0x70c95c)
				GUICtrlSetBkColor($btnRetur,0x6ccf56)
			endif
		endif
	next
endfunc


; Sort the cases by a given column.
func sortCases($by)
	$casesView = $cases
	if($by = 1) Then
		_ArraySort($casesView, 0, 0, 0, $iPersID)
	elseif($by = 2) Then
		_ArraySort($casesView, 0, 0, 0, $iStatus)
	elseif($by = 3) Then
		_ArraySort($casesView, 0, 0, 0, $iLastDc)
	elseif($by = 4) Then
		_ArraySort($casesView, 0, 0, 0, $iNextDateYear)
		;sort by year and month
		for $a=0 to UBound($casesView)-2
			for $b=0 to UBound($casesView)-2
				if (($casesView[$b][$iNextDateYear] = $casesView[$b+1][$iNextDateYear]) and ($casesView[$b][$iNextDateMonth] > $casesView[$b+1][$iNextDateMonth])) Then
					_ArraySort($casesView, 0, $b, $b+1, $iNextDateMonth)
				endif
			next
		next
		;sort by year,month and day
		for $aa=0 to UBound($casesView)-2
			for $bb=0 to UBound($casesView)-2
				if (($casesView[$bb][$iNextDateYear] = $casesView[$bb+1][$iNextDateYear] and $casesView[$bb][$iNextDateMonth] = $casesView[$bb+1][$iNextDateMonth]) and ($casesView[$bb][$iNextDateDay] > $casesView[$bb+1][$iNextDateDay])) Then
					_ArraySort($casesView, 0, $bb, $bb+1, $iNextDateDay)
				endif
			next
		next
		;sort by year,month,day and time
		for $aaa=0 to UBound($casesView)-2
			for $bbb=0 to UBound($casesView)-2
				if (($casesView[$bbb][$iNextDateYear] = $casesView[$bbb+1][$iNextDateYear] and $casesView[$bbb][$iNextDateMonth] = $casesView[$bbb+1][$iNextDateMonth] and $casesView[$bbb][$iNextDateDay] = $casesView[$bbb+1][$iNextDateDay]) and ($casesView[$bbb][$iNextTime] > $casesView[$bbb+1][$iNextTime])) Then
					_ArraySort($casesView, 0, $bbb, $bbb+1, $iNextTime)
				endif
			next
		next
	elseif($by = 5) Then
		_ArraySort($casesView, 0, 0, 0, $iCity)
		;sort by city and street
		for $c=0 to UBound($casesView)-2
			for $d=0 to UBound($casesView)-2
				if (($casesView[$d][$iCity] = $casesView[$d+1][$iCity]) and ($casesView[$d][$iAddressStreet] > $casesView[$d+1][$iAddressStreet])) Then
					_ArraySort($casesView, 0, $d, $d+1, $iAddressStreet)
				endif
			next
		next
		;make all cp integerz
		for $xx=0 to UBound($casesView)-2
			$casesView[$xx][$iAddressCp] = int($casesView[$xx][$iAddressCp])
		next
		;sort by city,street and cp
		for $cc=0 to UBound($casesView)-2
			for $dd=0 to UBound($casesView)-2
				if (($casesView[$dd][$iCity] = $casesView[$dd+1][$iCity] and $casesView[$dd][$iAddressStreet] = $casesView[$dd+1][$iAddressStreet]) and ($casesView[$dd][$iAddressCp] > $casesView[$dd+1][$iAddressCp])) Then
					_ArraySort($casesView, 0, $dd, $dd+1, $iAddressCp)
				endif
			next
		next
	endif
	fillList()
endfunc

; This function starts the VM with a given persID.
func StartCAPI($PersID)
	local $status = $cases[getCaseIndex($PersID)][$iStatus]
	;if case exists, resume the case.
	if($status = $cStatus_paused) then
		RunWait(@comspec & " /c " & $piaacpath & "ResumeCapi.exe " & $persID, $piaacpath, @SW_HIDE)
		RunWait(@comspec & " /c " & $vmexe & "deleteFileInGuest " & $vm & " /var/www/piaac/Exchange/Export/" & $persID & ".zip", $piaacpath, @SW_HIDE)
	;if case doesn't exist, import new case.
	Elseif($status = $cStatus_new) then
		$cases[getCaseIndex($persID)][$iStatus] = $cStatus_paused
		RunWait(@comspec & " /c " & $piaacpath & "StartCapi.exe " & $persID, $piaacpath, @SW_HIDE)
	;if case failed get the previous state and restore it
	Elseif($status = $cStatus_overfor) then
		;start blank CAPI
		RunWait(@comspec & " /c " & $piaacpath & "StartCapi.exe", $piaacpath, @SW_HIDE)
		;getty
		gettyVM()
		;check for XML file
		restartTAOHard()
	EndIf
endfunc

; This function gets the global disposition code from the [PersID]-Var.xml file.
func getEndingCode($persid)
	local $gEndingCode
	DirRemove($outputpath & $persID, 1)
	RunWait('"' & "c:\progra~1\vmware\vmware player\unzip.exe" & '"' & " -o -q " & $persID & ".zip -l " & $persID & "-Var.xml -d " & $outputpath, $outputpath, @SW_HIDE)

	local $file = FileOpen($outputpath & $persID & "-Var.xml", 0)
	while 1
		local $line = FileReadLine($file)
		if(@error = -1 or $line = "") then ExitLoop
		if(StringInStr($line, "ENDING")) then
			$m = StringRegExp($line, '=\"(\d{1})\"', 1)
			if(UBound($m) > 0) Then
				$gEndingCode = $m[0]
				exitloop
			endif
		endif
	wend
	FileClose($file)
	FileDelete($outputpath & $persID & "-Var.xml")
	return $gEndingCode
endfunc

; This function will poll the VM for the export file. When the export file for persID is found, it will be copied to the host and then the VM is automatically closed.
func waitForFinish()
	while 1
		;check for export file
		RunWait(@comspec & " /c " & $vmexe & "FileExistsInGuest " & $vm & " /var/www/piaac/Exchange/Export/" & $persID & ".zip > case_exists.txt", $piaacpath, @SW_HIDE)
		$line = FileReadLine($piaacpath & "case_exists.txt")
		; failed ?
		$waitforfinishstatus = $cases[getCaseIndex($persID)][$iStatus]
		if($line = "The file exists.") then
			; Try to export the selected case.
			RunWait(@comspec & " /c " & $piaacpath & "ExportResult.exe " & $persID, $piaacpath, @SW_HIDE)
			global $EndingCode = getEndingCode($persID)
			; Incomplete interview.
			if(FileExists($outputpath & $persID & ".zip") and $EndingCode = 0) Then
				$cases[getCaseIndex($persID)][$iStatus] = $cStatus_paused
				StopVM()
				exitloop
			; Complete interview.
			elseif(FileExists($outputpath & $persID & ".zip") and $EndingCode > 0) Then
				$cases[getCaseIndex($persID)][$iStatus] = $cStatus_intervju
				StopVM()
				exitloop
			endif
		;failed interview
		elseif ($waitforfinishstatus = $cStatus_overfor) then
			StopVM()
			exitloop
		Else
			sleep(2000)
		endif
	wend
	;clean after self
	FileDelete($piaacpath & "case_exists.txt")
endfunc

;escape sequence function
Func escape()
	Global $failed = 0
	;escape menu
	$winEscape = GUICreate("SC&C - Záchranný mód", 445, 180, -1, -1)
	$btnEscapeRestartTAO = GUICtrlCreateButton("Krok 1", 96, 148, 75, 25)
	$btnEscapeRestartVM = GUICtrlCreateButton("Krok 2", 184, 148, 75, 25)
	$btnEscapeTerminate = GUICtrlCreateButton("Chyba", 272, 148, 75, 25)
	$btnEscapeExit = GUICtrlCreateButton("Storno", 360, 148, 75, 25)
	$EscapeLabel = GUICtrlCreateLabel("Spustili jste záchranný mód CMS:" & @CRLF & @CRLF & "Pokud nastaly technické problémy, zkuste Krok 1." & @CRLF & @CRLF & "Nebude-li úspìšní, pokraèujte Krokem 2." & @CRLF & @CRLF &"Pokud se problém nevyøešil, stisknìte tlaèítko Chyba.", 16, 16, 450, 100)

	GUISetState(@SW_SHOW, $winEscape)

	While 1
		$msg = GUIGetMsg()
		if($msg = $GUI_EVENT_CLOSE) Then
			ExitLoop
		EndIf
		if($msg = $btnEscapeRestartTAO) Then
			GUICtrlSetState($btnEscapeRestartTAO, $GUI_DISABLE)
			GUICtrlSetState($btnEscapeRestartVM, $GUI_DISABLE)
			GUICtrlSetState($btnEscapeTerminate, $GUI_DISABLE)
			GUICtrlSetState($btnEscapeExit, $GUI_DISABLE)
			;if player is running
			if ProcessExists("vmplayer.exe") > 0 then
				screenshot()
				;display the VM window full screen(fromStartCAPI.au3)
				WinActivate("[CLASS:VMPlayerFrame]", "")
				$hwnd = WinWaitActive("[CLASS:VMPlayerFrame]", "", 10)
				$size = WinGetPos("[active]")
				If $size <> 0 Then
					If $size[2] < @DesktopWidth Or $size[3] < @DesktopHeight Then
						If $hwnd <> 0 Then
							;send keys {LCTRL}{LALT}{ENTER}
							Send("^!{ENTER}")
						EndIf
					EndIf
				EndIf
				;fullscreen first
				restartTAO()
			endIf
			ExitLoop
		EndIf
		if($msg = $btnEscapeRestartVM) Then
			GUICtrlSetState($btnEscapeRestartTAO, $GUI_DISABLE)
			GUICtrlSetState($btnEscapeRestartVM, $GUI_DISABLE)
			GUICtrlSetState($btnEscapeTerminate, $GUI_DISABLE)
			GUICtrlSetState($btnEscapeExit, $GUI_DISABLE)
			;if player is running
			if ProcessExists("vmplayer.exe") > 0 then
				screenshot()
				;try to be nice on VM
				StopVM()
				;try to be hard if there is no lock on VM
				if ProcessExists("vmplayer.exe") > 0 then ProcessClose("vmplayer.exe")
				;reboot
				RunWait($piaacpath & "StartCAPI.exe", $piaacpath, @SW_HIDE)
				gettyVM()
				restartTAO()
			EndIf
			Exitloop
		EndIf
		if($msg = $btnEscapeTerminate) Then
			GUICtrlSetState($btnEscapeRestartTAO, $GUI_DISABLE)
			GUICtrlSetState($btnEscapeRestartVM, $GUI_DISABLE)
			GUICtrlSetState($btnEscapeTerminate, $GUI_DISABLE)
			GUICtrlSetState($btnEscapeExit, $GUI_DISABLE)
			screenshot()
			;try to get the last export from archive
			RunWait(@comspec & " /c " & $vmexe & "copyFileFromGuestToHost " & $vm & " /var/www/piaac/Exchange/Archive/" & $PersID & ".zip " & $outputpath & $persID & ".zip", $piaacpath, @SW_HIDE)
			;try to be nice on VM
			stopVM()
			;try to be hard if there is no lock on VM
			if ProcessExists("vmplayer.exe") > 0 then ProcessClose("vmplayer.exe")
			; try to kill all stacked *.exe
			if ProcessExists("StopVM.exe") > 0 then ProcessClose("StopVM.exe")
			if ProcessExists("ExportResult.exe") > 0 then ProcessClose("ExportResult.exe")
			if ProcessExists("StartCAPI.exe") > 0 then ProcessClose("StartCAPI.exe")
			if ProcessExists("ResumeCAPI.exe") > 0 then ProcessClose("ResumeCAPI.exe")
			;mark persid as failed
			$cases[getCaseIndex($persID)][$iStatus] = $cStatus_overfor
			Exitloop
		EndIf
		if($msg = $btnEscapeExit) Then
			Exitloop
		EndIf
	WEnd
	GUIDelete($winEscape)
EndFunc


;Enable/disable CMS buttons
func cmsvisibility($visible)
	if $visible = 1 then $setstate = $GUI_ENABLE
	If $visible = 0 then $setstate = $GUI_DISABLE
	GUICtrlSetState($btnStart, $setstate)
	GUICtrlSetState($btnSort, $setstate)
	GUICtrlSetState($btnRetur, $setstate)
	GUICtrlSetState($btnScreen, $setstate)
	GUICtrlSetState($btnVisitpad, $setstate)
	GUICtrlSetState($btnExport, $setstate)
	GUICtrlSetState($btnPatch, $setstate)
	GUICtrlSetState($btnClose, $setstate)
EndFunc


;Wait for VM getty
func gettyVM()
	while 1
		RunWait(@comspec & " /c " & $vmexe & "listProcessesInGuest " & $vm & " > " & $piaacpath & "gettyprocesslist.txt", $piaacpath, @SW_HIDE)
		if FileExists($piaacpath & "gettyprocesslist.txt") Then
			$proccessfile = FileOpen($piaacpath & "gettyprocesslist.txt", 0)
			While 1
				$proccessfileline = FileReadLine($proccessfile)
				$getty = StringInStr($proccessfileline, "getty")
				if ($getty > 0 ) Then ExitLoop
				if (@error or $proccessfileline = "") then ExitLoop
			WEnd
			FileClose($proccessfile)
		endif
		if ($getty > 0) then ExitLoop
	WEnd
	;clean after self
	FileDelete($piaacpath & "gettyprocesslist.txt")
	return
EndFunc

;Restart TAO by state
func restartTAO()
	RunWait(@comspec & " /c " & $vmexe & "runProgramInGuest " & $vm & " /usr/bin/killall firefox-bin", $piaacpath, @SW_HIDE)
	if ($cases[getCaseIndex($persID)][$iStatus] = 0) then
		RunWait(@comspec & " /c " & $vmexe & "runProgramInGuest " & $vm & " /root/startImportAndStartTAO.sh " & $persID, $piaacpath, @SW_HIDE)
	ElseIf ($cases[getCaseIndex($persID)][$iStatus] = 9) then
		RunWait(@comspec & " /c " & $vmexe & "runProgramInGuest " & $vm & " /root/startResumeTAO.sh " & $persID, $piaacpath, @SW_HIDE)
	EndIf
endfunc

;Restart TAO by XML
func restartTAOHard()
		RunWait(@comspec & " /c " & $vmexe & "FileExistsInGuest " & $vm & " /var/www/piaac/Exchange/Import/" & $persID & "-Var.xml > extract_exists.txt", $piaacpath, @SW_HIDE)
		$getline = FileReadLine($piaacpath & "extract_exists.txt")
		if $getline = "The file exists." then
			$cases[getCaseIndex($persID)][$iStatus] = $cStatus_paused
			RunWait(@comspec & " /c " & $vmexe & "runProgramInGuest " & $vm & " /root/startResumeTAO.sh " & $persID, $piaacpath, @SW_HIDE)
			RunWait(@comspec & " /c " & $vmexe & "deleteFileInGuest " & $vm & " /var/www/piaac/Exchange/Export/" & $persID & ".zip", $piaacpath, @SW_HIDE)
		else
			$cases[getCaseIndex($persID)][$iStatus] = $cStatus_paused
			RunWait(@comspec & " /c " & $vmexe & "copyFileFromHostToGuest " & $vm & " " & $inputpath & $persID & ".zip" & " /var/www/piaac/Exchange/Import/" & $PersID & ".zip", $piaacpath, @SW_HIDE)
			RunWait(@comspec & " /c " & $vmexe & "runProgramInGuest " & $vm & " /root/startImportAndStartTAO.sh " & $persID, $piaacpath, @SW_HIDE)
		EndIf
		;clean after self
		FileDelete($piaacpath & "extract_exists.txt")
EndFunc


; This function finds the array index in cases-array of a given persID.
func getCaseIndex($persID)
	return _ArraySearch($cases, $persID, 0, 0, 0, 0, 1, $iPersID)
endfunc

; This function finds the array index in cases-array of a given persID.
func getCaseSortIndex($persID)
	$casesSortView = $casesView
	;select status filter  and display filter by status
	if $cases[getCaseIndex($persID)][$iStatus] = 9 or $cases[getCaseIndex($persID)][$iStatus] = 0 then
			$st_1 = 1
			$st_2 = 2
			;display filter
			$showRetur = 0
		Else
			$st_1 = 0
			$st_2 = 9
			;display filter
			$showRetur = 1
	EndIf
	;filter pre-sorted casesview array
	$status_array=_ArrayFindAll($casesSortView, $st_1, 0, 0, 0, 0, $iStatus)
	$status_array_2=_ArrayFindAll($casesSortView, $st_2, 0, 0, 0, 0, $iStatus)
	for $ai = 0 to uBound($casesSortView)-UBound($status_array)-UBound($status_array_2)-1
		while $casesSortView[$ai][$iStatus] = $st_1 or $casesSortView[$ai][$iStatus] = $st_2
			_ArrayDelete($casesSortView,$ai)
		WEnd
	next
	;refill the list to match the right display table
	filllist()
	return _ArraySearch($casesSortView, $persID, 0, 0, 0, 0, 1, $iPersID)
endfunc

; This function finds the selected persID.
func getPersID()
	$selectedRow = _GUICtrlListView_GetItemTextArray($listview, -1)
	return $selectedRow[1]
endfunc


; This function stops the VM by running a shutdown command within the VM.
func stopVM()
	RunWait($piaacpath & "StopVM.exe", $piaacpath, @SW_HIDE)
	while 300
		;check lock rather then proccess
		if FileExists($vm_current & "*.lck") Then
			sleep(100)
		Else
			ExitLoop
		endif
	wend
endfunc

; Repair VM by saving the current VM and restoring a fresh backup.
func repairVM()
	cmsvisibility(0)
	if(msgbox(8195, "SC&C - Opravný mód", "Dostali jste se do opravného módu CMS" & @CRLF & @CRLF & "Chcete opravit obraz?") = 6) then
		if DirGetSize($vm_restore) = 0 Then
			MsgBox(0,"Obraz nelze opravit.", "Obraz nelze opravit, kontaktujte supervizora.")
		else
			;move the current
			FileMove($vm_current & "*.*", $vm_backup, 9)
			DirMove($vm_current & "caches", $vm_backup, 1)
			$dir = DirGetSize($vm_current, 1)
			if($dir[1] = 0) Then
				;restore from backup
				FileMove($vm_restore & "*.*", $vm_current,9)
				;move paused cases status
				for $i = 0 to UBound($cases) - 1
					if $cases[$i][$iStatus] = $cStatus_paused then
						$cases[$i][$iStatus] = $cStatus_intervju
					endif
				next
				;write the state refresh
				writeAllCases()
				;resort the CMS list
				sortCases($sortBy)
				MsgBox(8192, "Obraz opraven", "Obraz byl opraven.")
			endIf
		endif
	endif
	cmsvisibility(1)
endfunc


; Start recording audio.
func recordingStart($persID)
	local $tTime = _Date_Time_GetLocalTime()
	local $aTime = _Date_Time_SystemTimeToArray($tTime)
	$datetime = $atime[2] & "-" & $atime[0] & "-" & $atime[1] & "_" & $atime[3] & "." & $atime[4] & "." & $atime[5]
	;record random 3 from first six and not PERSID beginning with number 9 and only new case
	if stringleft($persid,1) <> "9" and $cases[getCaseIndex($persID)][$iStatus] = $cStatus_new then
		for $vrr=1 to 3
			if FileReadLine($random_file,$vrr) = FileReadLine($random_file,4) then
				Run(@ComSpec & " /c " & $ogg & " -record -minimize -silent -timelimit 1200 -preset " & $oggpreset & " -filter " & $oggfilter & " -output " & $outputpath & $persID & "_" & $datetime & ".ogg", $oggdir, @SW_HIDE)
			EndIf
		next
		_FileWriteToLine($random_file,4,FileReadLine($random_file,4)+1,1)
	endif
endfunc


; Stop recording audio.
func recordingStop()
	if stringleft($persid,1) <> "9" then
		for $vrr=1 to 3
			if FileReadLine($random_file,$vrr) = FileReadLine($random_file,4)-1 then
				Run(@ComSpec & " /c " & $ogg & " -quit", $oggdir, @SW_HIDE)
			EndIf
		next
	endif
endfunc


;This function will generate ziped XML templates
func XMLgenerate()
	$file = FileOpen($cmsfile, 0)

	while 1
		$line = FileReadLine($file)
		if(@error = -1 or $line = "") then ExitLoop
			$vars = StringSplit($line, ";", 2)
		if(StringIsDigit($vars[0])) Then
			if(not(FileExists($inputpath & $vars[0] & ".zip"))) Then
			; Write to xml-file.
			createFile($vars[0], $vars[13], $vars[14])
			; Zip the xml-file.
			RunWait('"' & "c:\progra~1\vmware\vmware player\zip.exe" & '"' & " -q " & $vars[0] & ".zip " & $vars[0] & "-Var.xml", $inputpath, @SW_HIDE)
			FileDelete($inputpath & $vars[0] & "-Var.xml")
			endif
		endif
	WEnd
	FileClose($file)
Endfunc

; This function writes the xml-file.
func createFile($id, $addressStreet, $addressCp)
	; Opens the file as UTF-8 (128) and write (2).
	local $file = FileOpen($inputpath & $id & "-Var.xml", 130)
	filewrite($file, "<?xml version=""1.0""?>" & @CRLF)
	filewrite($file, "<hyperClassExport>" & @CRLF)
	filewrite($file, "  <hyperClass >" & @CRLF)
	filewrite($file, "    <hyperObject>" & @CRLF)
	filewrite($file, "      <property label=""skipBQLang"" code=""skipBQLang"" isValid=""1"">" & @CRLF)
	filewrite($file, "        <resource label=""BQLANG question should be skipped"" code=""01""/>" & @CRLF)
	filewrite($file, "      </property>" & @CRLF)
	filewrite($file, "      <property label=""skipCILang"" code=""skipCILang"" isValid=""1"">" & @CRLF)
	filewrite($file, "        <resource label=""CILANG question should be skipped"" code=""01""/>" & @CRLF)
	filewrite($file, "      </property>" & @CRLF)
	filewrite($file, "      <property label=""CI_PERSID"" code=""CI_PERSID"" isValid=""1"">" & @CRLF)
	filewrite($file, "        <literal value=""" & $id & """/>" & @CRLF)
	filewrite($file, "      </property> " & @CRLF)
	filewrite($file, "      <property label=""CI_PERSID"" code=""CI_PERSID"" isValid=""1"">" & @CRLF)
	filewrite($file, "        <resource label=""CI_PERSID"" code=""00""/>" & @CRLF)
	filewrite($file, "      </property> " & @CRLF)
	filewrite($file, "      <property label=""CI_Address"" code=""CI_Address"" isValid=""1"">" & @CRLF)
	filewrite($file, "        <literal value=""" & $addressStreet & " " & $addressCp & """/>" & @CRLF)
	filewrite($file, "      </property>" & @CRLF)
	filewrite($file, "      <property label=""CI_Address"" code=""CI_Address"" isValid=""1"">" & @CRLF)
	filewrite($file, "        <resource label=""CI_Address"" code=""00""/>" & @CRLF)
	filewrite($file, "      </property>" & @CRLF)
    filewrite($file, "      <property label=""CILANG"" code=""CILANG"" isValid=""1"">" & @CRLF)
    filewrite($file, "        <resource label=""Èeský"" code=""07""/>" & @CRLF)
    filewrite($file, "      </property>" & @CRLF)
	filewrite($file, "      <property label=""BQLANG"" code=""BQLANG"" isValid=""1"">" & @CRLF)
	filewrite($file, "        <resource label=""Èeský"" code=""09""/>" & @CRLF)
	filewrite($file, "      </property>" & @CRLF)
	filewrite($file, "    </hyperObject>" & @CRLF)
	filewrite($file, "  </hyperClass>" & @CRLF)
	filewrite($file, "</hyperClassExport>" & @CRLF)
	fileclose($file)
endfunc

; Take a screanshot function
Func screenshot()
	GUISetState(@SW_HIDE)
	sleep(100)
    Local $hBmp
	; Capture full screen
    $hBmp = _ScreenCapture_Capture("")
    ; Save bitmap to file
    _ScreenCapture_SaveImage($outputpath & $persID & ".jpg", $hBmp)
	sleep(5)
	GUISetState(@SW_SHOW)
EndFunc


;------- import/export --------


;Import / export function
;server side directory structure
;
;\  .......(root; TC-appointments.csv(import only), TC.csv(export only))
;\TC\ .....(TC directory; import/PERSID.csv(import only), PERSID.zip (exportonly)PERSID_.ogg(export only), PERSID_screenshot.jpg(export only))
;
Func importexport()
	;create parsed csv file from template
	template()
	$transferproblem = ""
	cmsvisibility(0)
	;create session
	$session = _FTP_Open('FTP session')
	$conn = _FTP_Connect($session,$ftpIP, $ftpuser, $ftppass,1)
	if @error then
		;error solution GUI
		ftp_error()
	else
	;IMPORT
		;import TC.csv
		; set CAPI/import directory
		_Ftp_DirSetCurrent($conn, $ftproot & $tc & "/import")
		if @error then
			$transferproblem &= @CRLF & "Nastavení adresáøe " & $ftproot & $tc & "/import" & " selhalo"
		else
			;get file list
			$remotecvsarray = _FTP_ListToArray($conn,2,$INTERNET_FLAG_RELOAD)
			if not @error then
				ProgressOn("Importuji Váš nový TÚ", "Importuji Váš nový TÚ", "Èekejte...")
				;import list
				for $r = 1 to UBound($remotecvsarray) - 1
					ProgressSet(round($r / (UBound($remotecvsarray) - 1) * 100), "Soubor: " & $remotecvsarray[$r])
					_FTP_FileGet($Conn, $remotecvsarray[$r], $inputpath & $remotecvsarray[$r])
					if @error then $transferproblem &= @CRLF & " Import TÚ selhal"
					;delete remote file
					_FTP_FileDelete($Conn, $remotecvsarray[$r])
					if @error then $transferproblem &= @CRLF & "Odstranìní TÚ selhalo"
				next
			endif
			ProgressOff()
		endif
	;EXPORT
		;export CAPI.csv
		;set root directory
		_Ftp_DirSetCurrent($conn, $ftproot)
		if @error then
			$transferproblem &= @CRLF & "Nastavení adresáøe " & $ftproot & " selhalo"
		Else
			;export file
			ProgressOn("Exportuji Vaše hlášení", "Exportuji Vaše hlášení", "Èekejte...")
			ProgressSet(50, "Soubor: " & $tc & ".csv")
			;export the master file
			_FTP_FilePut($Conn, $piaacpath & $tc & ".parsed", $tc & ".csv")
			if @error then $transferproblem &= @CRLF & "Export hlášení selhal"
			ProgressOff()
		endif
		;export PERSID.zip
		;set tc directory
		_Ftp_DirSetCurrent($conn, $ftproot & $tc)
		if @error then
			$transferproblem &= @CRLF & "Nastavení adresáøe " & $ftproot & $tc & " selhalo"
		else
			;get file list
			$exportziparray = _FileListToArray($outputpath, "*.zip", 1)
			if ( not( @error = 4 )) then
				ProgressOn("Exportuji data", "Exportuji data", "Èekejte...")
				;export list
				for $s = 1 to Ubound($exportziparray) - 1
					ProgressSet(round($s / (UBound($exportziparray) - 1) * 100), "Soubor: " & $exportziparray[$s])
					;then update
					_FTP_FilePut($conn, $outputpath & $exportziparray[$s], $exportziparray[$s])
					if @error then $transferproblem &= @CRLF & "Export dat " & $exportziparray[$s]  & " selhal"
				next
				ProgressOff()
			endif
			;export PESID*.jpg
			;get file list
			$exportjpgarray = _FileListToArray($outputpath, "*.jpg", 1)
			if ( not( @error = 4 )) then
				ProgressOn("Exportuji snímky", "Exportuji snímky", "Èekejte...")
				;export list
				for $v = 1 to Ubound($exportjpgarray) - 1
					ProgressSet(round($v / (UBound($exportjpgarray) - 1) * 100), "Soubor: " & $exportjpgarray[$v])
					;then update
					_FTP_FilePut($conn, $outputpath & $exportjpgarray[$v], $exportjpgarray[$v])
					if @error then
						$transferproblem &= @CRLF & "Export snímku " & $exportjpgarray[$v] & " selhal"
					else
						;mark as exported
						FileMove($outputpath & $exportjpgarray[$v], $outputpath & $exportjpgarray[$v] & ".exported")
					EndIf
				next
				ProgressOff()
			EndIf
			;export PESID*.ogg
			;get file list
			$exportoggarray = _FileListToArray($outputpath, "*.ogg", 1)
			if ( not( @error = 4 )) then
				ProgressOn("Exportuji zvukové nahrávky", "Exportuji zvukové nahrávky", "Èekejte...")
				;export list
				for $u = 1 to Ubound($exportoggarray) - 1
					ProgressSet(round($u / (UBound($exportoggarray) - 1) * 100), "Soubor: " & $exportoggarray[$u])
					;then update
					_FTP_FilePut($conn, $outputpath & $exportoggarray[$u], $exportoggarray[$u])
					if @error then
						$transferproblem &= @CRLF & "Export nahrávky " & $exportoggarray[$u] & " selhal"
					else
						;mark as exported
						FileMove($outputpath & $exportoggarray[$u], $outputpath & $exportoggarray[$u] & ".exported")
					EndIf
				next
				ProgressOff()
			EndIf
		endif
		;success /failed message
		if ($transferproblem = "") then
			MsgBox(0,"Import / Export", "Operace probìhla v poøádku. Zmìny se projeví po novém spuštìní CMS.")
		elseif ($transferproblem <> "" ) then
			MsgBox(16,"Chyba","Pøi pøenosu došlo k následující chybì:" & @CRLF & $transferproblem & @CRLF & @CRLF & "Opakujte pøenos..")
		endIf
	endif
	_FTP_Close($session)
	;drop the pre-exported master file for export
	FileDelete($piaacpath & $tc & ".parsed")
	cmsvisibility(1)
EndFunc


;------ Screener ---------


global $btnScreenerSave,$btnScreenerExit

;screener
func screener()
	cmsvisibility(0)
	;get actual array from file
	$cases = getAllCases($cmsfile)
	;get year for control
	$scrTime = _Date_Time_GetLocalTime()
	$scraTime = _Date_Time_SystemTimeToArray($scrTime)
	;check box marker
	$smarker = 0
	;get prefix
	$prefix = ""
	$phone = $cases[getCaseIndex($persID)][$iPhone]
	if StringLen($cases[getCaseIndex($persID)][$iPhone]) = 12 Then
		$prefix = StringLeft($cases[getCaseIndex($persID)][$iPhone],3)
		$phone =  StringTrimLeft($cases[getCaseIndex($persID)][$iPhone],3)
	endif

	$winScreener = GUICreate("SC&C - Screener", 335, 366, -1, -1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU)
	$ScreenerLabel1 = GUICtrlCreateLabel("PersID:", 12, 11, 36, 17)
	$ScreenerLabel2 = GUICtrlCreateLabel($persID, 53, 11, 200, 17)
	$ScreenerLabel3 = GUICtrlCreateLabel("Ulice:", 12, 52, 36, 17)
	$ScreenerInput1 = GUICtrlCreateInput($cases[getCaseIndex($persID)][$iAddressStreet] & " " & $cases[getCaseIndex($persID)][$iAddressCp], 196, 48, 89, 21)
	GUICtrlSetLimit(-1,25,0)
	GUICtrlSetTip(-1,"5-20 písmen, mezera, 1-4 èíslice 1-9999")
	GUICtrlSetState(-1,$GUI_DISABLE)
	$ScreenerCheckbox = GUICtrlCreateCheckbox("", 298, 50, 17, 17)
	$ScreenerLabel4 = GUICtrlCreateLabel("Poèet vhodných osob v domácnosti:", 12, 84, 190, 25)
	$ScreenerInput2 = GUICtrlCreateInput($cases[getCaseIndex($persID)][$iPeople], 196, 80, 25, 21)
	GUICtrlSetLimit(-1,2,0)
	GUICtrlSetTip(-1,"èíslice 1-19")
	$ScreenerLabel5 = GUICtrlCreateLabel("Poøadí vybrané osoby:", 12, 116, 130, 17)
	$ScreenerInput3 = GUICtrlCreateInput($cases[getCaseIndex($persID)][$iOrder], 196, 112, 25, 21)
	GUICtrlSetLimit(-1,2,0)
	GUICtrlSetTip(-1,"èíslice 1-19")
	$ScreenerLabel6 = GUICtrlCreateLabel("Pøíjmení:", 12, 148, 50, 17)
	$ScreenerInput4 = GUICtrlCreateInput($cases[getCaseIndex($persID)][$iSecondName], 196, 144, 129, 21)
	GUICtrlSetLimit(-1,20,0)
	GUICtrlSetTip(-1,"2-20 písmen")
	$ScreenerLabel7 = GUICtrlCreateLabel("Rodné jméno:", 12, 180, 70, 17)
	$ScreenerInput5 = GUICtrlCreateInput($cases[getCaseIndex($persID)][$iName], 196, 176, 129, 21)
	GUICtrlSetLimit(-1,20,0)
	GUICtrlSetTip(-1,"2-20 písmen")
	$ScreenerLabel8 = GUICtrlCreateLabel("Rok narození:", 12, 212, 100, 17)
	$ScreenerInput6 = GUICtrlCreateInput($cases[getCaseIndex($persID)][$iYear], 196, 208, 57, 21)
	GUICtrlSetLimit(-1,4,0)
	if $cases[getCaseIndex($persID)][$iAge] = "ST" then
		if $scraTime[2] = 2011 then
			GUICtrlSetTip(-1,"4 èíslice 1945-1995")
		elseif $scraTime[2] = 2012 then
			GUICtrlSetTip(-1,"4 èíslice 1946-1996")
		endIf
	ElseIf $cases[getCaseIndex($persID)][$iAge] = "ML" then
		if $scraTime[2] = 2011 then
			GUICtrlSetTip(-1,"4 èíslice 1981-1995")
		elseif $scraTime[2] = 2012 then
			GUICtrlSetTip(-1,"4 èíslice 1982-1996")
		endIf
	endif
	$ScreenerInput8 = GUICtrlCreateInput($prefix, 196, 240, 30, 21)
	GUICtrlSetLimit(-1,3,0)
	GUICtrlSetTip(-1,"3 èíslice 100-999")
	$ScreenerLabel9 = GUICtrlCreateLabel("[ Pøedvolba ] / Telefonní èíslo:", 12, 244, 150, 17)
	$ScreenerInput7 = GUICtrlCreateInput($phone, 236, 240, 89, 21)
	GUICtrlSetLimit(-1,9,0)
	GUICtrlSetTip(-1,"9 èíslic 100000000-999999999")
	$ScreenerLabel10 = GUICtrlCreateLabel("Specifikace:", 12, 276, 100, 17)
	$ScreenerCombo = GUICtrlCreateCombo("", 196, 272, 129, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	GUICtrlSetData(-1, "||Mobilní telefon|Domù|Práce|Soused", $cases[getCaseIndex($persID)][$iPhoneType])
global	$btnScreenerSave = GUICtrlCreateButton("Uložit", 161, 328, 75, 25)
global	$btnScreenerExit = GUICtrlCreateButton("Konec", 249, 328, 75, 25)
	$btnScreenerHelper = GUICtrlCreateButton("?", 228, 82, 19, 17)

	GUISetState(@SW_SHOW, $winScreener)

	While 1

		$msg = GUIGetMsg()
	;checkbox control
		if GUICtrlRead($ScreenerCheckbox) = 1 and GUICtrlGetState($ScreenerInput1) <> 80 then
			GUICtrlSetState($ScreenerInput1,$GUI_ENABLE)
			$smarker = 1
		EndIf
		if GUICtrlRead($ScreenerCheckbox) = 4 and GUICtrlRead($ScreenerInput1) = $cases[getCaseIndex($persID)][$iAddressStreet] & " " & $cases[getCaseIndex($persID)][$iAddressCp] and $smarker = 1 then
			GUICtrlSetState($ScreenerInput1,$GUI_DISABLE)
			$smarker = 0
		EndIf
	;screener input control
		$cscrfailed = 0
		if ($msg = $GUI_EVENT_CLOSE) then exitloop
		if ($msg = $btnScreenerHelper ) then
			screenability(0)
			MsgBox(32,"Poèet vhodných osob","2011:" & @CRLF & @CRLF & "ST (starší  ) 1945 - 1995" & @CRLF &  "ML (mladší ) 1981 - 1995" & @CRLF & @CRLF & "2012:" & @CRLF & @CRLF & "ST (starší  ) 1946 - 1996"  & @CRLF & "ML (mladší ) 1982 - 1996")
			screenability(1)
		endif
		if ($msg = $btnScreenerSave) then
			;phone spec. VS phone control VS prefix
			if GUICtrlRead($ScreenerInput7) <> "" and GUICtrlRead($ScreenerCombo) = "" then
					MsgBox(0, "Chyba", "Hodnota specifikace je prázdná.")
					$cscrfailed += 1
			endif
			if (GUICtrlRead($ScreenerInput7) = "" and GUICtrlRead($ScreenerCombo) <> "") or (GUICtrlRead($ScreenerInput8) <> "" and GUICtrlRead($ScreenerInput7) = "") then
					MsgBox(0, "Chyba", "Hodnota telefonní èíslo je prázdná.")
					$cscrfailed += 1
			endif
			;order <= people control
			if int(GUICtrlRead($ScreenerInput3)) > int(GUICtrlRead($ScreenerInput2)) Then
					MsgBox(0, "Chyba", "Poøadí vybrané osoby je vìtší než poèet osob v domácnosti.")
					$cscrfailed += 1
			EndIf
			;address
			if GUICtrlGetState($ScreenerInput1) = 80 Then
				if inputcontrol(GUICtrlRead($ScreenerInput1),9) = 1 and GUICtrlRead($ScreenerInput1) <> "" then
					$addressgex = StringRegExp(GUICtrlRead($ScreenerInput1),'^(\D{5,20})\040([1-9]\d{0,3})$',1)
					$cases[getCaseIndex($persID)][$iAddressStreet] = $addressgex[0]
					$cases[getCaseIndex($persID)][$iAddressCp] = $addressgex[1]
				Else
					MsgBox(0, "Chyba", "Ulice: nesprávná hodnota." )
					$cscrfailed += 1
				EndIf
			EndIf
			;people
			if inputcontrol(GUICtrlRead($ScreenerInput2),1) = 1 or GUICtrlRead($ScreenerInput2) = "" then
				$cases[getCaseIndex($persID)][$iPeople] = GUICtrlRead($ScreenerInput2)
			Else
				MsgBox(0, "Chyba", "Poèet vhodných osob v domácnosti: nesprávná hodnota.")
				$cscrfailed += 1
			EndIf
			;people order
			if inputcontrol(GUICtrlRead($ScreenerInput3),1) = 1 or GUICtrlRead($ScreenerInput3) = "" then
				$cases[getCaseIndex($persID)][$iOrder] = GUICtrlRead($ScreenerInput3)
			Else
				MsgBox(0, "Chyba", "Poøadí vybrané osoby:  nesprávná hodnota.")
				$cscrfailed += 1
			EndIf
			;second name
			if inputcontrol(GUICtrlRead($ScreenerInput4),2) = 1  or GUICtrlRead($ScreenerInput4) = "" then
				$cases[getCaseIndex($persID)][$iSecondName] = GUICtrlRead($ScreenerInput4)
			Else
				MsgBox(0, "Chyba", "Pøíjmení: nesprávná hodnota.")
				$cscrfailed += 1
			EndIf
			;first name
			if inputcontrol(GUICtrlRead($ScreenerInput5),2) = 1 or GUICtrlRead($ScreenerInput5) = "" then
			$cases[getCaseIndex($persID)][$iName] = GUICtrlRead($ScreenerInput5)
			Else
				MsgBox(0, "Chyba", "Rodné jméno: nesprávná hodnota.")
				$cscrfailed += 1
			EndIf
			;year
			if $scraTime[2] = 2011 and $cases[getCaseIndex($persID)][$iAge] = "ST" then
				if inputcontrol(GUICtrlRead($ScreenerInput6),3) = 1 or GUICtrlRead($ScreenerInput6) = "" then
					$cases[getCaseIndex($persID)][$iYear] = GUICtrlRead($ScreenerInput6)
				Else
					MsgBox(0, "Chyba", "Rok narození: nesprávná hodnota.")
					$cscrfailed += 1
				EndIf
			elseif $scraTime[2] = 2012 and $cases[getCaseIndex($persID)][$iAge] = "ST" then
				if inputcontrol(GUICtrlRead($ScreenerInput6),5) = 1 or GUICtrlRead($ScreenerInput6) = "" then
					$cases[getCaseIndex($persID)][$iYear] = GUICtrlRead($ScreenerInput6)
				Else
					MsgBox(0, "Chyba", "Rok narození: nesprávná hodnota.")
					$cscrfailed += 1
				EndIf
			elseif $scraTime[2] = 2011 and $cases[getCaseIndex($persID)][$iAge] = "ML" then
				if inputcontrol(GUICtrlRead($ScreenerInput6),10) = 1 or GUICtrlRead($ScreenerInput6) = "" then
					$cases[getCaseIndex($persID)][$iYear] = GUICtrlRead($ScreenerInput6)
				Else
					MsgBox(0, "Chyba", "Rok narození: nesprávná hodnota.")
					$cscrfailed += 1
				EndIf
			elseif $scraTime[2] = 2012 and $cases[getCaseIndex($persID)][$iAge] = "ML" then
				if inputcontrol(GUICtrlRead($ScreenerInput6),11) = 1 or GUICtrlRead($ScreenerInput6) = "" then
					$cases[getCaseIndex($persID)][$iYear] = GUICtrlRead($ScreenerInput6)
				Else
					MsgBox(0, "Chyba", "Rok narození: nesprávná hodnota.")
					$cscrfailed += 1
				EndIf
			endif
			;prefix
			if GUICtrlRead($ScreenerInput7) <> "" then
				if inputcontrol(GUICtrlRead($ScreenerInput8),12) = 1 or GUICtrlRead($ScreenerInput8) = "" then
					$prefix = GUICtrlRead($ScreenerInput8)
				else
					MsgBox(0, "Chyba", "Telefonní pøedvolba: nesprávná hodnota.")
					$cscrfailed += 1
				endIf
			endif
			;phone number
			if inputcontrol(GUICtrlRead($ScreenerInput7),4) = 1 or GUICtrlRead($ScreenerInput7) = "" then
				$cases[getCaseIndex($persID)][$iPhone] = GUICtrlRead($ScreenerInput7)
				if GUICtrlRead($ScreenerInput7) <> "" and GUICtrlRead($ScreenerInput8) <> "" then
					$cases[getCaseIndex($persID)][$iPhone] = $prefix & GUICtrlRead($ScreenerInput7)
				endif
			else
				MsgBox(0, "Chyba", "Telefonní èíslo: nesprávná hodnota.")
				$cscrfailed += 1
			EndIf
			;phone specification
			if inputcontrol(GUICtrlRead($ScreenerCombo),13) = 1 or GUICtrlRead($ScreenerCombo) = "" then
				$cases[getCaseIndex($persID)][$iPhoneType] = GUICtrlRead($ScreenerCombo)
			else
				MsgBox(0, "Chyba", "Specifikace: nesprávná hodnota.")
				$cscrfailed += 1
			endif
		;save
			if $cscrfailed = 0 then
				;write all ceses
				writeAllCases()
				;sort by last checked
				sortCases($sortBy)
				;make finished persid selected
				_GUICtrlListView_SetItemSelected($listview,getCaseSortIndex($persID))
				ExitLoop
			endif
		endif
		if($msg = $btnScreenerExit) Then ExitLoop
	WEnd
	GUIDelete($winScreener)
	cmsvisibility(1)
EndFunc

;shadow screener buttons when in helper
func screenability($screenvisible)
	if $screenvisible = 1 then $setstate = $GUI_ENABLE
	If $screenvisible = 0 then $setstate = $GUI_DISABLE
	GUICtrlSetState($btnScreenerSave, $setstate)
	GUICtrlSetState($btnScreenerExit, $setstate)
EndFunc


;------ Visit pad ---------


;predefined global variables for compilation fix
global $VisitpadListView,$btnVisitpadEdit,$btnVisitpadPreview,$btnVisitpadExit
global $btnVisiteditSave,$btnVisiteditExit

;visitpad
func visitpad()
	cmsvisibility(0)
	;get actual array from file
	$cases = getAllCases($cmsfile)

	$winVisitpad = GUICreate("SC&C - Návštìvy", 537, 292, -1, -1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU)
	$VisitpadLabel1 = GUICtrlCreateLabel("PersID:", 8, 8, 36, 17)
	$VisitpadLabel2 = GUICtrlCreateLabel($persID, 48, 8, 196, 17)
global	$VisitpadListView = GUICtrlCreateListView("Poøadí|Datum|Èas|Disp. kód|Poznámka|Datum další návštìvy|Èas další návštìvy", 9, 32, 519, 209)
global	$btnVisitpadEdit = GUICtrlCreateButton("Upravit", 273, 256, 75, 25)
global	$btnVisitpadPreview = GUICtrlCreateButton("Náhled", 365, 256, 75, 25)
global	$btnVisitpadExit = GUICtrlCreateButton("Konec", 453, 256, 75, 25)

	GUISetState(@SW_SHOW, $winVisitpad)
	fillvisitpadlist()
	While 1
		$msg = GUIGetMsg()
		if ($msg = $GUI_EVENT_CLOSE) then exitloop
		if($msg = $btnVisitpadEdit) Then
					;get the line number
					$selectedline = _GUICtrlListView_GetItemTextArray($VisitpadListView, -1)
					$linenumber = $selectedline[1]
					if ($linenumber = "") then
						msgbox(0, "Návštìva", "Žádná návštìva nebyla vybrána.")
					elseif (($cases[getCaseIndex($persID)][$linenumber*7+18] <> "")) or $linenumber = 0 then
						msgbox(0, "Návštìva", "Návštìvu nelze upravit.")
					else
						visitedit($linenumber)
					endif
		endif
		if($msg = $btnVisitpadPreview) Then
				padibility(0)
					;get the line number
					$selectedline = _GUICtrlListView_GetItemTextArray($VisitpadListView, -1)
					$linenumber = $selectedline[1]
					if ($linenumber = "") then
						msgbox(0, "Náhled", "Žádná návštìva nebyla vybrána.")
					elseif (($cases[getCaseIndex($persID)][$linenumber*7+18] <> "")) or $linenumber = 0 then
						msgbox(32, "Náhled","Dispozièní kód: "& @CRLF & @CRLF & $cases[getCaseIndex($persID)][$linenumber*7+18] & @CRLF & @CRLF & "Datum a èas návštìvy: " & @CRLF & @CRLF & $cases[getCaseIndex($persID)][$linenumber*7+15] & " " & $cases[getCaseIndex($persID)][$linenumber*7+16] & @CRLF & @CRLF & "Poznámka: " & @CRLF & @CRLF & $cases[getCaseIndex($persID)][$linenumber*7+17] & @CRLF & @CRLF & "Datum a èas další navštìvy: " & @CRLF & @CRLF & $cases[getCaseIndex($persID)][$linenumber*7+13]  & " " & $cases[getCaseIndex($persID)][$linenumber*7+14] )
					else
						msgbox(0, "Náhled", "Náhled nelze zobrazit.")
					EndIf
				padibility(1)
		EndIf
		if($msg = $btnVisitpadExit) Then ExitLoop
	WEnd
	GUIDelete($winVisitpad)
	cmsvisibility(1)
EndFunc

;shadow editor buttons when editting
func padibility($padvisible)
	if $padvisible = 1 then $setstate = $GUI_ENABLE
	If $padvisible = 0 then $setstate = $GUI_DISABLE
	GUICtrlSetState($btnVisitpadEdit, $setstate)
	GUICtrlSetState($btnVisitpadPreview, $setstate)
	GUICtrlSetState($btnVisitpadExit, $setstate)
EndFunc

;Generate vists list
func fillvisitpadlist()
	_GUICtrlListView_DeleteAllItems($VisitpadListView)
	;generate lines
	for $number = 0 to $cases[getCaseIndex($persID)][$iLastNote]
		if $number <= 9 then
			$str1 = $number+1 & "|"
			$str1 &= $cases[getCaseIndex($persID)][$number*7+22] & "|"  ;date
			$str1 &= $cases[getCaseIndex($persID)][$number*7+23] & "|"	;time
			$str1 &= $cases[getCaseIndex($persID)][$number*7+25] & "|"	;code
			$str1 &= $cases[getCaseIndex($persID)][$number*7+24] & "|"	;note
			$str1 &= $cases[getCaseIndex($persID)][$number*7+20] & "|"	;next date
			$str1 &= $cases[getCaseIndex($persID)][$number*7+21]		;next time
			$editline = GUICtrlCreateListViewItem($str1, $VisitpadListView)
			;grey the not empty one
			if ($cases[getCaseIndex($persID)][$number*7+25] <> "") then GUICtrlSetBkColor($editline, 0xc9c9c9)
		endif
	next
endfunc

;shadow editor editing buttons when helper is open
func padibility_helper($padeditorvisible)
	if $padeditorvisible = 1 then $setstate = $GUI_ENABLE
	If $padeditorvisible = 0 then $setstate = $GUI_DISABLE
	GUICtrlSetState($btnVisiteditSave, $setstate)
	GUICtrlSetState($btnVisiteditExit, $setstate)
EndFunc

;Visits line editor
func visitedit($linenumber)
	padibility(0)
	;check box marker
	$marker = 0

	$winVisitedit = GUICreate("Návštìva - Upravit", 557, 137, -1, -1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU)
global $btnVisiteditSave = GUICtrlCreateButton("Uložit", 386, 104, 75, 25)
global $btnVisiteditExit = GUICtrlCreateButton("Storno", 474, 104, 75, 25)
	$btnVisiteditHelper = GUICtrlCreateButton("?", 135, 10, 19, 17)
	$VisiteditLabel1 = GUICtrlCreateLabel("Dispozièní kód:", 8, 14, 80, 17)
	$VisiteditCombo = GUICtrlCreateCombo("", 88, 8, 40, 21, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	GUICtrlSetData($VisiteditCombo, "AP|CB|NH|RB|RL|IL|OT|1|3|4|5|7|8|9|12|13|14|15|16|17|18|19|20|21|22|24|26|27|28|90", "")
	$VisiteditLabel7 = GUICtrlCreateLabel("Osobnì:", 41, 45, 40, 17)
	$VisiteditCheckbox1 = GUICtrlCreateCheckbox("", 88, 43, 17, 17)
;	GUICtrlSetState(-1, $GUI_CHECKED)
	$VisiteditLabel8 = GUICtrlCreateLabel("Telefonicky:", 21, 77, 60, 17)
	$VisiteditCheckbox2 = GUICtrlCreateCheckbox("", 88, 75, 17, 17)
	$VisiteditLabel2 = GUICtrlCreateLabel("Datum:", 227, 14, 38, 17)
	$VisiteditInput1 = GUICtrlCreateInput("", 269, 8, 65, 21);note date
	GUICtrlSetLimit(-1,10,0)
	GUICtrlSetTip(-1,"DD.MM.RRRR")
	$VisiteditLabel3 = GUICtrlCreateLabel("Èas:", 453, 14, 25, 17)
	$VisiteditInput2 = GUICtrlCreateInput("", 482, 8, 65, 21);note time
	GUICtrlSetLimit(-1,5,0)
	GUICtrlSetTip(-1,"HH:MM")
	$VisiteditLabel4 = GUICtrlCreateLabel("Datum další návštìvy:", 156, 45, 108, 17)
	$VisiteditInput3 = GUICtrlCreateInput("", 269, 40, 65, 21);next note date
	GUICtrlSetLimit(-1,10,0)
	GUICtrlSetTip(-1,"DD.MM.RRRR")
	$VisiteditLabel5 = GUICtrlCreateLabel("Èas další návštìvy:", 382, 45, 95, 17)
	$VisiteditInput4 = GUICtrlCreateInput("", 482, 40, 65, 21);next note time
	GUICtrlSetLimit(-1,5,0)
	GUICtrlSetTip(-1,"HH:MM")
	$VisiteditLabel6 = GUICtrlCreateLabel("Poznámka:", 207, 77, 57, 17);note text
	$VisiteditInput5 = GUICtrlCreateInput("", 268, 72, 279, 21)
	GUICtrlSetLimit(-1,250,0)
	GUICtrlSetTip(-1,"2-250 znakù")

	GUISetState(@SW_SHOW,$winVisitedit)

	While 1
		$msg = GUIGetMsg()
		;checkbox switch
		if GUICtrlRead($VisiteditCheckbox2) = 1 and  $marker = 0 then
			GUICtrlSetState($VisiteditCheckbox1, $GUI_UNCHECKED)
			$marker = 1
		EndIf
		if GUICtrlRead($VisiteditCheckbox1) = 1 and $marker = 1 then
			GUICtrlSetState($VisiteditCheckbox2, $GUI_UNCHECKED)
			$marker = 0
		EndIf
		;input control
		$cvisitfailed = 0
		if ($msg = $GUI_EVENT_CLOSE) then exitloop
		if ($msg = $btnVisiteditHelper ) then
			padibility_helper(0)
			msgbox(32,"Návštìvy - Dispozièní kód","Kód - Definice" & @CRLF  & @CRLF & "Pøechodné kódy:" & @CRLF & @CRLF & "AP - Domluvena schùzka. Zaznamenejte podrobnosti o schùzce do poznámky." & @CRLF & "CB - Domácnost kontaktována bez závazného termínu schùzky. Zaznamenej-" & @CRLF & "        te si obecnou informaci o tom, kdy se pøíštì máte zastavit, do poznámky." & @CRLF & "NH - Nikdo nezastižen." & @CRLF & "RB - Poèáteèní odmítnutí. Èlen domácnosti se odmítl zúèastnit nebo pøed" & @CRLF & "        dokonèením dotazník pøerušil." & @CRLF & "RL - Odmítnutí + dopis. Èlen domácnosti nebo respondent se odmítl zúèastnit" & @CRLF & "        nebo pøed dokonèením dotazník pøerušil. Byl by vhodný pøesvìdèovací" & @CRLF & "        dopis. Podrobnosti zapište do sekce pro poznámky." & @CRLF & " IL - Nemoc/Indispozice." & @CRLF & "OT - Jiné. Jakákoliv další situace, která vyžaduje sledování, jako nemoc" & @CRLF & "        respondenta, jazykové potíže, nemožnost lokalizovat místo pobytu," & @CRLF & "        nebo zamèený objekt." & @CRLF & @CRLF & "Finální kódy:" & @CRLF & @CRLF & "  1 - Dokonèení dotazníku." & @CRLF & "  3 - Èásteèné dokonèení/Pøerušení dotazníku." & @CRLF & "  4 - Odmítnutí – respondent/èlen domácnosti." & @CRLF & "  5 - Odmítnutí – jiný." & @CRLF & "  7 - Jazykový problém." & @CRLF & "  8 - Potíže se ètením a psaním." & @CRLF & "  9 - Mentální handicap/Potíže s uèením." & @CRLF & "12 - Vada sluchu." & @CRLF & "13 - Slepota/Vada zraku." & @CRLF & "14 - Vada øeèi." & @CRLF & "15 - Fyzický handicap." & @CRLF & "16 - Jiný handicap." & @CRLF & "17 - Jiné (nespecifikované), jako nemoc nebo neoèekávané okolnosti." & @CRLF & "18 - Úmrtí respondenta." & @CRLF & "19 - Žádný respondent ve vìku 16-65." & @CRLF & "20 - Obytná jednotka nenalezena." & @CRLF & "21 - Maximální poèet návštìv." & @CRLF & "22 - Obytná jednotka v rekonstrukci." & @CRLF & "24 - Doèasná nepøítomnost/Nedostupnost bìhem doby terénního výzkumu." & @CRLF & "26 - Prázdná obytná jednotka." & @CRLF & "27 - Duplicita – již dotazován." & @CRLF & "28 - Na adrese není obytná jednotka." & @CRLF & "90 - Technický problém.")
			padibility_helper(1)
		endif
		if ($msg = $btnVisiteditSave ) then
			;check box empty control
				if GUICtrlRead($VisiteditCheckbox1) = $GUI_UNCHECKED and GUICtrlRead($VisiteditCheckbox2) = $GUI_UNCHECKED then
					msgbox (0, "Chyba", "Hodnota typ návštìvy musí být oznaèena.")
					$cvisitfailed += 1
				endif
			;empty control
			if GUIctrlread($VisiteditCombo) = "" or GUICtrlRead($VisiteditInput1) = "" or GUICtrlRead($VisiteditInput2) = "" then
				msgbox (0, "Chyba", "Hodnota dispozièní kód, datum a èas musí být vyplnìna.")
				$cvisitfailed += 1
			EndIf
			;disp. code control
			if GUICtrlRead($VisiteditCombo) = "AP" then
				if GUICtrlRead($VisiteditInput3) = "" or GUICtrlRead($VisiteditInput4) = "" or GUICtrlRead($VisiteditInput5) = "" then
					MsgBox(0, "Chyba", "Hodnota datum, èas další návštìvy a poznámka musí být vyplnìna.")
					$cvisitfailed += 1
				EndIf
			endif
			if GUICtrlRead($VisiteditCombo) = "CB" or GUICtrlRead($VisiteditCombo) = "RL" then
				if GUICtrlRead($VisiteditInput5) = "" then
					MsgBox(0, "Chyba", "Poznámka musí být vyplnìna.")
					$cvisitfailed += 1
				EndIf
			EndIf
			;format control
			if GUIctrlread($VisiteditCombo) <> "" and inputcontrol(GUIctrlread($VisiteditCombo),14) <> 1 then
				MsgBox(0, "Chyba", "Dispozièní kód: nesprávná hodnota.")
				$cvisitfailed += 1
			endif
			if GUICtrlRead($VisiteditInput1) <> "" and inputcontrol(GUICtrlRead($VisiteditInput1),6) <> 1 then
				MsgBox(0, "Chyba", "Datum: nesprávná hodnota.")
				$cvisitfailed += 1
			endif
			if GUICtrlRead($VisiteditInput2) <> "" and inputcontrol(GUICtrlRead($VisiteditInput2),7) <> 1 then
				MsgBox(0, "Chyba", "Èas: nesprávná hodnota.")
				$cvisitfailed += 1
			endif
			if GUICtrlRead($VisiteditInput3) <> "" and inputcontrol(GUICtrlRead($VisiteditInput3),6) <> 1 then
				MsgBox(0, "Chyba", "Datum další návštìvy: nesprávná hodnota.")
				$cvisitfailed += 1
			endif
			if GUICtrlRead($VisiteditInput4) <> "" and inputcontrol(GUICtrlRead($VisiteditInput4),7) <> 1 then
				MsgBox(0, "Chyba", "Èas další návštìvy: nesprávná hodnota.")
				$cvisitfailed += 1
			endif
			if GUICtrlRead($VisiteditInput5) <> "" and inputcontrol(GUICtrlRead($VisiteditInput5),8) <> 1 then
				MsgBox(0, "Chyba", "Poznámka: nesprávná hodnota.")
				$cvisitfailed += 1
			endif
			; semicolon bypass bug check
			if GUICtrlRead($VisiteditInput5) <> "" and StringRegExp(GUICtrlRead($VisiteditInput5),'^.*;.*$',0) = 1 then
				MsgBox(0, "Chyba", "Poznámka: nepovolený znak ';'.")
				$cvisitfailed += 1
			endif
			;single time and single date control
			if GUIctrlread($VisiteditInput1) <> "" and GUIctrlread($VisiteditInput2) = "" and GUICtrlRead($VisiteditCombo) <> "AP" then
				MsgBox(0, "Chyba", "Hodnota èas není vyplnìna.")
				$cvisitfailed += 1
			endif
			if GUIctrlread($VisiteditInput1) = "" and GUIctrlread($VisiteditInput2) <> "" and GUICtrlRead($VisiteditCombo) <> "AP" then
				MsgBox(0, "Chyba", "Hodnota datum není vyplnìna.")
				$cvisitfailed += 1
			endif
			;exclude IL,CB
			if not (GUICtrlRead($VisiteditCombo) <> "IL" or GUICtrlRead($VisiteditCombo) <> "CB") then
				if GUIctrlread($VisiteditInput3) <> "" and GUIctrlread($VisiteditInput4) = "" and GUICtrlRead($VisiteditCombo) <> "AP" then
					MsgBox(0, "Chyba", "Hodnota èas další návštìvy není vyplnìna.")
					$cvisitfailed += 1
				endif
			endif
			if GUIctrlread($VisiteditInput3) = "" and GUIctrlread($VisiteditInput4) <> "" and GUICtrlRead($VisiteditCombo) <> "AP" then
				MsgBox(0, "Chyba", "Hodnota datum další návštìvy není vyplnìna.")
				$cvisitfailed += 1
			endif
			;date flow control
			$dategex = StringRegExp(GUIctrlread($VisiteditInput1),'^([0][1-9]|[12][0-9]|3[01])\.(0[1-9]|1[012])\.(2011|2012)$',1)
			if $linenumber = 1 then
				$csvdategex = ""
				else
				$csvdategex = StringRegExp($cases[getCaseIndex($persID)][($linenumber-1)*7+15],'^([0][1-9]|[12][0-9]|3[01])\.(0[1-9]|1[012])\.(2011|2012)$',1)
			EndIf
			If	$csvdategex <> "" and $cvisitfailed = 0 then
				;year must be the same or higher
				if int($csvdategex[2]) > int($dategex[2]) then
					MsgBox(0, "Chyba", "Hodnota datum a èas je stejná, nebo starší než predchozí.")
					$cvisitfailed += 1
				endif
				;if the year is the same the month must be the same or higher
				if int($csvdategex[2]) = int($dategex[2]) and int($csvdategex[1]) > int($dategex[1]) then
					MsgBox(0, "Chyba", "Hodnota datum a èas je stejná, nebo starší než predchozí.")
					$cvisitfailed += 1
				endif
				;if year is the same and month is the same the day must be the same or higher
				if int($csvdategex[2]) = int($dategex[2]) and int($csvdategex[1]) = int($dategex[1]) and int($csvdategex[0]) > int($dategex[0]) then
					MsgBox(0, "Chyba", "Hodnota datum a èas je stejná, nebo starší než predchozí.")
					$cvisitfailed += 1
				endif
				;if the year is the same the month is the same the day is the same the time must be higher
				if int($csvdategex[2]) = int($dategex[2]) and int($csvdategex[1]) = int($dategex[1]) and int($csvdategex[0]) = int($dategex[0]) and $cases[getCaseIndex($persID)][($linenumber-1)*7+16] >= GUIctrlread($VisiteditInput2) then
					MsgBox(0, "Chyba", "Hodnota datum a èas je stejná, nebo starší než predchozí.")
					$cvisitfailed += 1
				endif
			endif
			;next date flow control
			;$nextdategex = StringRegExp(GUIctrlread($VisiteditInput3),'^([0][1-9]|[12][0-9]|3[01])\.(0[1-9]|1[012])\.(2011|2012)$',1)
			;If	$nextdategex <> "" and $cvisitfailed = 0 then
			;	;year must be the same or higher
			;	if int($cases[getCaseIndex($persID)][$iNextDateYear]) > int($nextdategex[2]) then
			;		MsgBox(0, "Chyba", "Hodnota datum a èas další návštìvy je stejná, nebo starší než predchozí.")
			;		$cvisitfailed += 1
			;	endif
			;	;if the year is the same the month must be the same or higher
			;	if int($cases[getCaseIndex($persID)][$iNextDateYear]) = int($nextdategex[2]) and int($cases[getCaseIndex($persID)][$iNextDateMonth]) > int($nextdategex[1]) then
			;		MsgBox(0, "Chyba", "Hodnota datum a èas další návštìvy je stejná, nebo starší než predchozí.")
			;		$cvisitfailed += 1
			;	endif
			;	;if year is the same and month is the same the day must be the same or higher
			;	if int($cases[getCaseIndex($persID)][$iNextDateYear]) = int($nextdategex[2]) and int($cases[getCaseIndex($persID)][$iNextDateMonth]) = int($nextdategex[1]) and int($cases[getCaseIndex($persID)][$iNextDateDay]) > int($nextdategex[0]) then
			;		MsgBox(0, "Chyba", "Hodnota datum a èas další návštìvy je stejná, nebo starší než predchozí.")
			;		$cvisitfailed += 1
			;	endif
			;	;if the year is the same the month is the same the day is the same the time must be higher
			;	if int($cases[getCaseIndex($persID)][$iNextDateYear]) = int($nextdategex[2]) and int($cases[getCaseIndex($persID)][$iNextDateMonth]) = int($nextdategex[1]) and int($cases[getCaseIndex($persID)][$iNextDateDay]) = int($nextdategex[0]) and $cases[getCaseIndex($persID)][$iNextTime] >= GUIctrlread($VisiteditInput4) then
			;		MsgBox(0, "Chyba", "Hodnota datum a èas další návštìvy je stejná, nebo starší než predchozí.")
			;		$cvisitfailed += 1
			;	endif
			;endif
			;date VS  next date
			if $cvisitfailed = 0 then
				$dategex = StringRegExp(GUIctrlread($VisiteditInput1),'^([0][1-9]|[12][0-9]|3[01])\.(0[1-9]|1[012])\.(2011|2012)$',1)
				$nextdategex = StringRegExp(GUIctrlread($VisiteditInput3),'^([0][1-9]|[12][0-9]|3[01])\.(0[1-9]|1[012])\.(2011|2012)$',1)
				if	$nextdategex <> "" and $dategex <> "" then
					;year must be the same or higher
					if int($dategex[2]) > int($nextdategex[2]) then
						MsgBox(0, "Chyba", "Hodnota datum a èas další návštìvy je stejná, nebo starší než datum a èas návštìvy.")
						$cvisitfailed += 1
					endif
					;if the year is the same the month must be the same or higher
					if int($dategex[2]) = int($nextdategex[2]) and int($dategex[1]) > int($nextdategex[1]) then
						MsgBox(0, "Chyba", "Hodnota datum a èas další návštìvy je stejná, nebo starší než datum a èas návštìvy.")
						$cvisitfailed += 1
					endif
					;if year is the same and month is the same the day must be the same or higher
					if int($dategex[2]) = int($nextdategex[2]) and int($dategex[1]) = int($nextdategex[1]) and int($dategex[0]) > int($nextdategex[0]) then
						MsgBox(0, "Chyba", "Hodnota datum a èas další návštìvy je stejná, nebo starší než datum a èas návštìvy.")
						$cvisitfailed += 1
					endif
					;if the year is the same the month is the same the day is the same the time must be higher
					if int($dategex[2]) = int($nextdategex[2]) and int($dategex[1]) = int($nextdategex[1]) and int($dategex[0]) = int($nextdategex[0]) and GUIctrlread($VisiteditInput2) >= GUIctrlread($VisiteditInput4) then
						MsgBox(0, "Chyba", "Hodnota datum a èas další návštìvy je stejná, nebo starší než datum a èas návštìvy.")
						$cvisitfailed += 1
					endif
				endif
			endif
			;save
			if $cvisitfailed = 0 then
				;save visit type
				if GUICtrlRead($VisiteditCheckbox2) = 1 then
				$cases[getCaseIndex($persID)][$linenumber*7+19] = "T" ;type
					else
				$cases[getCaseIndex($persID)][$linenumber*7+19] = "O" ;type
				endif
				;save dates,times,notes and codes
				$cases[getCaseIndex($persID)][$linenumber*7+13] = GUIctrlread($VisiteditInput3)	;next date
				$cases[getCaseIndex($persID)][$linenumber*7+14] = GUIctrlread($VisiteditInput4)	;next time
				$cases[getCaseIndex($persID)][$linenumber*7+15] = GUIctrlread($VisiteditInput1)	;date
				$cases[getCaseIndex($persID)][$linenumber*7+16] = GUIctrlread($VisiteditInput2)	;time
				$cases[getCaseIndex($persID)][$linenumber*7+17] = GUICtrlRead($VisiteditInput5)	;note
				$cases[getCaseIndex($persID)][$linenumber*7+18] = GUIctrlread($VisiteditCombo)	;code
				;update last index and dc
				if $cases[getCaseIndex($persID)][$iLastNote] <= 9 then
					$cases[getCaseIndex($persID)][$iLastNote] += 1
					if GUIctrlread($VisiteditCombo) = 1 or GUIctrlread($VisiteditCombo) = 3 or GUIctrlread($VisiteditCombo) = 4 or GUIctrlread($VisiteditCombo) = 5 or GUIctrlread($VisiteditCombo) = 7 or GUIctrlread($VisiteditCombo) = 8 or GUIctrlread($VisiteditCombo) = 9 then
						$cases[getCaseIndex($persID)][$iLastDc] = "0"
						$cases[getCaseIndex($persID)][$iLastDc] &= GUIctrlread($VisiteditCombo)
					else
						$cases[getCaseIndex($persID)][$iLastDc] = GUIctrlread($VisiteditCombo)
					endif
				endif
				;update $iNextDate/$iNextTime
				$nextdategex = StringRegExp(GUIctrlread($VisiteditInput3),'^([0][1-9]|[12][0-9]|3[01])\.(0[1-9]|1[012])\.(2011|2012)$',1)
				if $nextdategex <> "" then
					$cases[getCaseIndex($persID)][$iNextDateDay] = $nextdategex[0]
					$cases[getCaseIndex($persID)][$iNextDateMonth] = $nextdategex[1]
					$cases[getCaseIndex($persID)][$iNextDateYear] = $nextdategex[2]
					$cases[getCaseIndex($persID)][$iNextTime] = GUIctrlread($VisiteditInput4)
				endif
				;write all ceses
				writeAllCases()
				;refill the list
				fillvisitpadlist()
				;resort the CMS list
				sortCases($sortBy)
				;make finished persid selected
				_GUICtrlListView_SetItemSelected($listview, getCaseSortIndex($persID))
				ExitLoop
			Endif
		EndIf
		if($msg = $btnVisiteditExit) Then ExitLoop
	Wend
	GUIDelete($winVisitedit)
	padibility(1)
EndFunc


;------ functions  ---------


;check input string by type, return 1 if match elese 0
func inputcontrol($string, $type)
	if $string <> "" then
		if $type > 0 and $type < 15 then
			Switch $type
				case $type = 1
					if StringRegExp($string, '^[1-9]$|^1[0-9]$', 0) = 1 and StringLen($string) < 3 then return 1 ; people,order = 1 digit
				case $type = 2
					if StringRegExp($string, '^\D{2,20}$', 0) = 1 and StringLen($string) < 20  and StringLen($string) > 1 then return 1 ;name/second name = 2-20 literal
				case $type = 3
					if StringRegExp($string, '^19\d{2}$', 0) = 1 and StringLen($string) = 4 and $string > 1944 and $string < 1996 then return 1 ;ST year(2011) = 4 digit
				case $type = 4
					if StringRegExp($string, '^[1-9]{1}\d{8}$', 0) = 1 and StringLen($string) = 9 then return 1 ; phone = 9 digit
				case $type = 5
					if (StringRegExp($string, '^19\d{2}$', 0) = 1 and StringLen($string) = 4 and $string > 1945 and $string < 1997)  or $string = "1945" then return 1 ;ST year(2012) = 4 digit
				case $type = 6
					if StringRegExp($string, '^([0][1-9]|[12][0-9]|3[01])\.(0[1-9]|1[012])\.(2011|2012)$', 0) = 1 and StringLen($string) = 10 then return 1 ; date = 2 dig . 2 dig . 4dig
				case $type = 7
					if StringRegExp($string, '^([01][0-9]|2[0-4])\:([0-5][0-9])$', 0) = 1  and StringLen($string) = 5 then return 1 ; time = 2 dig. : 2 dig.
				case $type = 8
					if StringRegExp($string, '^.{1,250}$', 0) = 1 and StringLen($string) > 1  and StringLen($string) < 250 then return 1 ; note = 1-250 dig./lit.
				case $type = 9
					if StringRegExp($string, '^\D{5,20}\040[1-9]\d{0,3}$', 0) = 1 and StringLen($string) > 5 and StringLen($string) < 20 then return 1 ; address = 5-20 lit. space 1-4 dig.
				case $type = 10
					if StringRegExp($string, '^19\d{2}$', 0) = 1 and StringLen($string) = 4 and $string > 1980 and $string < 1996 then return 1 ;ML year(2011) = 4 digit
				case $type = 11
					if (StringRegExp($string, '^19\d{2}$', 0) = 1 and StringLen($string) = 4 and $string > 1981 and $string < 1997) or $string = "1981" then return 1 ;ML year(2012) = 4 digit
				case $type = 12
					if StringRegExp($string, '^\d{3}$', 0) = 1 and StringLen($string) = 3 then return 1 ;prefix = 3 digits
				case $type = 13
					if StringRegExp($string, '^(Mobilní telefon|Domù|Práce|Soused)$', 0) = 1 then return 1 ;screener mob. spec.
				case $type = 14
					if StringRegExp($string, '^(AP|CB|NH|RB|RL|IL|OT|1|3|4|5|7|8|9|12|13|14|15|16|17|18|19|20|21|22|24|26|27|28|90)$', 0) = 1 then return 1 ;visitpad dc. spec.
			EndSwitch
		else
			return 0
		endIf
	else
		return 0
	EndIf
EndFunc

;get random file for recording
Func get_rnd_file()
	if not FileExists($random_file) then
		;get random 3 numbers to file
		$rndfile = FileOpen($random_file,1)
		while 1
			if (_FileCountLines($random_file) = 3) then ExitLoop
			$rndnum = Random(1,6,1)
			if FileReadLine($random_file,1) <> $rndnum then
				if FileReadLine($random_file,2) <> $rndnum then
					FileWrite($random_file,$rndnum & @CRLF)
				endif
			endIf
		WEnd
		;set the record mark number
		FileWrite($random_file,1)
		FileClose($rndfile)
	EndIf
EndFunc

;get all disp codes
func getDispCodes($persid)
	local $gDispCodes
	RunWait('"' & "c:\progra~1\vmware\vmware player\unzip.exe" & '"' & " -o -q " & $persID & ".zip -l " & $persID & "-Var.xml -d " & $outputpath, $outputpath, @SW_HIDE)

	local $file = FileOpen($outputpath & $persID & "-Var.xml", 0)

	if not(FileExists($outputpath & $persid & "-Var.xml")) then
			$gDispCodes = ";;;;;;;;"
	else
	while 1
		; Read one line from the xml-file.
		$dcline = FileReadLine($file)
		; If end of file, exit the loop.
		if(@error = -1 or $dcline = "") then ExitLoop
		if(StringInStr($dcline, '"DISP_CI"')) then
			$n = StringRegExp($dcline, '=\"(\d{1})\"', 1)
			$dcline = FileReadLine($file)
			if StringRegExp($dcline, 'resource', 0) = 1  or StringRegExp($dcline, 'literal', 0) = 1 then
				$m = StringRegExp($dcline, '=\"(\d{2})\"', 1)
				if(UBound($m) > 0) Then
					$gDispCodes &=";" & $m[0] & ";"
				endif
			else
				$gDispCodes &= ";" & ";"
			;else
			;	if(UBound($n) > 0) Then
			;		$gDispCodes &= ";" & $n[0] & ";"
			;	endif
			EndIf
		endif
		if(StringInStr($dcline, '"DISP_BQ"')) then
			$n = StringRegExp($dcline, '=\"(\d{1})\"', 1)
			$dcline = FileReadLine($file)
			if StringRegExp($dcline, 'resource', 0) = 1  or StringRegExp($dcline, 'literal', 0) = 1 then
				$m = StringRegExp($dcline, '=\"(\d{2})\"', 1)
				if(UBound($m) > 0) Then
					$gDispCodes &= $m[0] & ";"
				endif
			else
				$gDispCodes &= ";"
			;else
			;	if(UBound($n) > 0) Then
			;		$gDispCodes &= $n[0] & ";"
			;	endif
			EndIf
		endif
		if(StringInStr($dcline, '"DISP_CORE"')) then
			$n = StringRegExp($dcline, '=\"(\d{1})\"', 1)
			$dcline = FileReadLine($file)
			if StringRegExp($dcline, 'resource', 0) = 1  or StringRegExp($dcline, 'literal', 0) = 1 then
				$m = StringRegExp($dcline, '=\"(\d{2})\"', 1)
				if(UBound($m) > 0) Then
					$gDispCodes &= $m[0] & ";"
				endif
			else
				$gDispCodes &= ";"
			;else
			;	if(UBound($n) > 0) Then
			;		$gDispCodes &= $n[0] & ";"
			;	endif
			endif
		endif
		if(StringInStr($dcline, '"DISP_CBA"')) then
			$n = StringRegExp($dcline, '=\"(\d{1})\"', 1)
			$dcline = FileReadLine($file)
			if StringRegExp($dcline, 'resource', 0) = 1  or StringRegExp($dcline, 'literal', 0) = 1 then
				$m = StringRegExp($dcline, '=\"(\d{2})\"', 1)
				if(UBound($m) > 0) Then
					$gDispCodes &= $m[0] & ";"
				endif
			else
				$gDispCodes &= ";"
			;else
			;	if(UBound($n) > 0) Then
			;		$gDispCodes &= $n[0] & ";"
			;	endif
			endif
		endif
		if(StringInStr($dcline, '"DISP_PP"')) then
			$n = StringRegExp($dcline, '=\"(\d{1})\"', 1)
			$dcline = FileReadLine($file)
			if StringRegExp($dcline, 'resource', 0) = 1  or StringRegExp($dcline, 'literal', 0) = 1 then
				$m = StringRegExp($dcline, '=\"(\d{2})\"', 1)
				if(UBound($m) > 0) Then
					$gDispCodes &= $m[0] & ";"
				endif
			else
				$gDispCodes &= ";"
			;else
			;	if(UBound($n) > 0) Then
			;		$gDispCodes &= $n[0] & ";"
			;	endif
			EndIf
		endif
		if(StringInStr($dcline, '"DISP_PRC"')) then
			$n = StringRegExp($dcline, '=\"(\d{1})\"', 1)
			$dcline = FileReadLine($file)
			if StringRegExp($dcline, 'resource', 0) = 1  or StringRegExp($dcline, 'literal', 0) = 1 then
				$m = StringRegExp($dcline, '=\"(\d{2})\"', 1)
				if(UBound($m) > 0) Then
					$gDispCodes &= $m[0] & ";"
				endif
			else
				$gDispCodes &= ";"
			;else
			;	if(UBound($n) > 0) Then
			;		$gDispCodes &= $n[0] & ";"
			;	endif
			EndIf
		endif
		if(StringInStr($dcline, '"GLOBALDISPCODE"')) then
			$n = StringRegExp($dcline, '=\"(\d{1})\"', 1)
			$dcline = FileReadLine($file)
			if StringRegExp($dcline, 'resource', 0) = 1  or StringRegExp($dcline, 'literal', 0) = 1 then
				$m = StringRegExp($dcline, '=\"(\d{2})\"', 1)
				if(UBound($m) > 0) Then
					$gDispCodes &= $m[0] & ";"
				endif
			else
				$gDispCodes &= ";"
			;else
			;	if(UBound($n) > 0) Then
			;		$gDispCodes &= $n[0] & ";"
			;	endif
			endIf
		endif
	wend
	EndIf

	FileClose($file)
	FileDelete($outputpath & $persID & "-Var.xml")
	return $gDispCodes
endfunc

;fullscreen mode configuration patch
func patch()
	if not FileExists(@UserProfileDir & "\dataap~1\vmware\") then DirCreate(@UserProfileDir & "\dataap~1\vmware\")
	if not FileExists($vmpref) then _FileCreate($vmpref)

	$vmpreffile = FileOpen($vmpref,1)

	if FileGetSize($vmpref) = 0 then
		FileWriteLine($vmpreffile,".encoding = " & '"windows-1250"')
		FileWriteLine($vmpreffile,"pref.trayicon.enabled = " & '"true"')
		FileWriteLine($vmpreffile,"pref.usbDev.maxDevs = " & '"0"')
		FileWriteLine($vmpreffile,"pref.keyboardAndMouse.maxProfiles = " & '"0"')
		FileWriteLine($vmpreffile,"hint.cui.toolsInfoBar.suppressible = " & '"FALSE"')
;		FileWriteLine($vmpreffile,"hints.hideAll = " & '"TRUE"')
		FileWriteLine($vmpreffile,"pref.vmplayer.componentDownloadPermission = " & '"deny"')
;		FileWriteLine($vmpreffile,"pref.vmplayer.autoSoftwareUpdatePermission = " & '"deny"')
;		FileWriteLine($vmpreffile,"pref.vmplayer.dataCollectionEnabled = " & '"no"')
		FileWriteLine($vmpreffile,"pref.vmplayer.exit.vmAction = " & '"poweroff"')
;		FileWriteLine($vmpreffile,"pref.vmplayer.fullscreen.autohide = " & '"TRUE"')
		FileWriteLine($vmpreffile,"pref.vmplayer.fullscreen.nobar = " & '"TRUE"')
		FileWriteLine($vmpreffile,"pref.fullscreen.v5 = "& '"TRUE"')
		FileWriteLine($vmpreffile,"pref.fullscreen.nobar = " & '"TRUE"')
		FileWriteLine($vmpreffile,"pref.eula.0.appName = " & '"VMware Player"')
		FileWriteLine($vmpreffile,"pref.eula.0.buildNumber = " & '"261024"')
		FileClose($vmpreffile)
	else
		for $o = 1 to _FileCountLines($vmpref)
			if StringRegExp(FileReadLine($vmpref,$o),'pref\.fullscreen\.v5',0) = 1 then ExitLoop
			if $o = _FileCountLines($vmpref) then
				$vmpreffile = FileOpen($vmpref,1)
				FileWriteLine($vmpreffile,"pref.vmplayer.fullscreen.nobar = " & '"TRUE"')
				FileWriteLine($vmpreffile,"pref.fullscreen.v5 = "& '"TRUE"')
				FileWriteLine($vmpreffile,"pref.fullscreen.nobar = " & '"TRUE"')
				FileClose($vmpreffile)
			endif
		next
	endif
EndFunc

;get interview timing and duration
func getepoch($persid)
	RunWait('"' & "c:\progra~1\vmware\vmware player\unzip.exe" & '"' & " -o -q " & $persID & ".zip -l " & $persID & "-Log.xml -d " & $outputpath, $outputpath, @SW_HIDE)

	local $logfile = FileOpen($outputpath & $persID & "-Log.xml", 0)
	local $logcontrol = 0
	local $epochstr = ""

	if FileExists($outputpath & $persid & "-Log.xml") then
		;interview end control
		for $e = 1 to _FileCountLines($outputpath & $persid & "-Log.xml")
			if StringRegExp(FileReadLine($outputpath & $persid & "-Log.xml",$e),'INTERVIEW_END',0) = 1 then
				$logcontrol = 1
			Endif
		next
		;get ecpoch and calculate the duration
		if $logcontrol = 1 then
			while 1
				; Read one line from the xml-file.
				$epochline = FileReadLine($logfile)
				; If end of file, exit the loop.
				if(@error = -1 or $epochline = "") then ExitLoop
				if(StringInStr($epochline, '"INTERVIEW_START"')) then
					$f = StringRegExp($epochline, '=\"(\d{10})\"', 1)
					if(UBound($f) > 0) Then
						$epochstr = $f[0] & " "
					endif
				endIf
				if(StringInStr($epochline, '"INTERVIEW_PAUSE"')) then
					$f = StringRegExp($epochline, '=\"(\d{10})\"', 1)
					if(UBound($f) > 0) Then
						$epochstr &= $f[0] & " "
					endif
				endif
				if(StringInStr($epochline, '"INTERVIEW_RESUME"')) then
					$f = StringRegExp($epochline, '=\"(\d{10})\"', 1)
					if(UBound($f) > 0) Then
						$epochstr &= $f[0] & " "
					endif
				endif
				if(StringInStr($epochline, '"INTERVIEW_END"')) then
					$f = StringRegExp($epochline, '=\"(\d{10})\"', 1)
					if(UBound($f) > 0) Then
						$epochstr &= $f[0]
					endif
				endif
			wend
		EndIf
	endIf

	FileClose($logfile)
	FileDelete($outputpath & $persID & "-Log.xml")

	if $epochstr <> "" then
		$epocharray = StringSplit($epochstr," ")
		$epochstartdate = DllCall("msvcrt.dll", "str:cdecl", "ctime", "int*", $epocharray[1])
		$epochstopdate = DllCall("msvcrt.dll", "str:cdecl", "ctime", "int*", $epocharray[$epocharray[0]])
		$durationsec = 0

		while UBound($epocharray) > 1
			$durationsec = $durationsec + int($epocharray[2]) - int($epocharray[1])
			_ArrayDelete($epocharray,1)
			_ArrayDelete($epocharray,1)
		wend

		$duration = sec2time($durationsec)
		$epoch = stringreplace($epochstartdate[0] & ";" & $duration & ";" & $epochstopdate[0] & ";",@CR,"")
		return stringreplace($epoch,@LF,"")
	else
		return ";;;"
	endIf
endfunc

;convertseconds to time
func sec2time($fr_sec)
   $sec2time_hour = Int($fr_sec / 3600)
   $sec2time_min = Int(($fr_sec - $sec2time_hour * 3600) / 60)
   $sec2time_sec = $fr_sec - $sec2time_hour * 3600 - $sec2time_min * 60
   Return StringFormat('%02d:%02d:%02d', $sec2time_hour, $sec2time_min, $sec2time_sec)
endfunc

;create master file copy for export and append disp. codes + timing
func template()
	cmsvisibility(0)
	;set splash on
	SplashTextOn("Poèítám hodnoty", "Poèítám výsledné hodnoty." & @CRLF & "Prosím poèkejte...",240,65,-1,-1,-1,"",10)
	;get acctual case state array
	$cases = getAllCases($cmsfile)
	;get export time
	$expTime = _Date_Time_GetLocalTime()
	$expaTime = _Date_Time_SystemTimeToArray($expTime)
	$expdatetime = $expatime[1] & "." & $expatime[0] & "." & $expatime[2] & "," & $expatime[3] & ":" & $expatime[4]
	;create template/parsed file if not exist
	if not FileExists($piaacpath & $tc & ".template") then _FileCreate($piaacpath & $tc & ".template")
	;re-create parsed file
	;FileDelete($piaacpath & $tc & ".parsed")
	;if not FileExists($piaacpath & $tc & ".parsed") then _FileCreate($piaacpath & $tc & ".parsed")
	_FileCreate($piaacpath & $tc & ".parsed")
	;get master file line array
	local $line,$var,$pid,$pindex
	local $tplarray[1][1],$persidarray[1][1]

	_FileReadToArray($cmsfile,$persidarray)
	if FileGetSize($cmsfile) <> 0 Then
		for $line = 1 to UBound($persidarray)-1
			$var = StringSplit($persidarray[$line], ";", 2)
			$pid = $var[0]
			;get status = finished;
			if $cases[getCaseIndex($pid)][$iStatus] = 1 then
				$template = getAllCases($piaacpath & $tc & ".template")
				$pindex = _ArraySearch($template,$pid, 0, 0, 0, 0, 1, 0)
				if  $pindex >= 0 then
					;write from array
					FileWriteLine($piaacpath & $tc & ".parsed", $persidarray[$line] & ";" & $template[$pindex][1] & ";" & $template[$pindex][2] & ";" & $template[$pindex][3] & ";" & $template[$pindex][4] & ";" & $template[$pindex][5] & ";" & $template[$pindex][6] & ";" & $template[$pindex][7] & ";" & $template[$pindex][8] & ";" & $template[$pindex][9] & ";" & $template[$pindex][10] & ";" & $expdatetime)
				else
					;write to template
					local $thecode = getDispCodes($pid)
					local $thepoch = getepoch($pid)
					FileWriteLine($piaacpath & $tc & ".template", $pid & $thecode & $thepoch)
					;write previous to file
					FileWriteLine($piaacpath & $tc & ".parsed", $persidarray[$line] & $thecode & $thepoch & $expdatetime)
				endif
			else
				;recalculate disp. codez, ;;; ,time and write
				FileWriteLine($piaacpath & $tc & ".parsed", $persidarray[$line] & getDispCodes($pid) & ";;;" & $expdatetime)
			endif
		next
	endIf
	;write the first line
	_FileWriteToLine($piaacpath & $tc & ".parsed",1,"PersID;Status;AssignDate;NextDateDay;NextDateMonth;NextDateYear;NextTime;SecondName;FirstName;Age;Year;Phone;PhoneType;AddressStreet;AddressCp;City;People;Order;LastNote;LastDc;Visit1nextdate;Visit1nexttime;Visit1date;Visit1time;Visit1note;Visit1code;Visit1type;Visit2nextdate;Visit2nexttime;Visit2date;Visit2time;Visit2note;Visit2code;Visit2type;Visit3nextdate;Visit3nexttime;Visit3date;Visit3time;Visit3note;Visit3code;Visit3type;Visit4nextdate;Visit4nexttime;Visit4date;Visit4time;Visit4note;Visit4code;Visit4type;Visit5nextdate;Visit5nexttime;Visit5date;Visit5time;Visit5note;Visit5code;Visit5type;Visit6nextdate;Visit6nexttime;Visit6date;Visit6time;Visit6note;Visit6code;Visit6type;Visit7nextdate;Visit7nexttime;Visit7date;Visit7time;Visit7note;Visit7code;Visit7type;Visit8nextdate;Visit8nexttime;Visit8date;Visit8time;Visit8note;Visit8code;Visit8type;Visit9nextdate;Visit9nexttime;Visit9date;Visit9time;Visit9note;Visit9code;Visit9type;Visit10nextdate;Visit10nexttime;Visit10date;Visit10time;Visit10note;Visit10code;Visit10type;Disp_ci;Disp_bq;Disp_core;Disp_cba;Disp_pp;Disp_prc;Globaldispcode;Epoch_start;Duration;Epoch_stop;Expdatetime",0)
	;set splash on
	SplashOff()
	cmsvisibility(1)
EndFunc

Func taopatch()
	cmsvisibility(0)
	RunWait(@comspec & " /c " & $piaacpath & "PatchVM.exe", $piaacpath, @SW_HIDE)
	ProcessWaitClose("vmware-vmx.exe")
	msgBox(0,"Záplata", "Operace probìhla v poøádku. Zmìny se projeví po novém spuštìní dotazníku.")
	;remove the patch
	FileDelete($piaacpath & "patch\*.gpg")
	cmsvisibility(1)
EndFunc

func ftp_error()
	cmsvisibility(0)

	local $ttime = _Date_Time_GetLocalTime()
	local $atime = _Date_Time_SystemTimeToArray($ttime)
	$ftp_error_time = $atime[2] & "-" & $atime[0] & "-" & $atime[1]

	$winFtp_error = GUICreate("Chyba", 340, 186,-1, -1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU)
	$ftp_errorIcon = GUICtrlCreateIcon("shell32.dll", 11, 11, 5, 32, 32)
	$ftp_errorLabel1 = GUICtrlCreateLabel("Pøipojení selhalo", 50, 12, 316, 17)
	GUICtrlCreateGroup("",10,35,321,106)
	$ftp_errorRadio1 = GUICtrlCreateRadio("Pøipojení zkontrolováno. Opakovat pøenos.", 16, 52, 303, 17)
	GUICtrlSetState($ftp_errorRadio1, $GUI_CHECKED)
	$ftp_errorRadio2 = GUICtrlCreateRadio("Pøipojení nefunguje. Vygenerovat hlášení k odeslání.", 16, 83, 303, 17)
	$ftp_errorRadio3 = GUICtrlCreateRadio("Pøipojení nefunguje. Vygenerovat hlášení a data k odeslání.", 16, 115, 303, 17)
	$btnftp_errorPotvrdit = GUICtrlCreateButton("OK", 168, 152, 75, 25)
	$btnftp_errorStorno = GUICtrlCreateButton("Storno", 256, 152, 75, 25)

	GUISetState(@SW_SHOW)

	while 1
		$msg = GUIGetMsg()
		if $msg = $GUI_EVENT_CLOSE or $msg = $btnftp_errorStorno then exitloop
		if $msg = $btnftp_errorPotvrdit then
			if(bitand(GUICtrlRead($ftp_errorRadio1), $GUI_CHECKED)) then exitloop
			if(bitand(GUICtrlRead($ftp_errorRadio2), $GUI_CHECKED)) then
				GUICtrlSetState($ftp_errorRadio1, $GUI_DISABLE)
				GUICtrlSetState($ftp_errorRadio2, $GUI_DISABLE)
				GUICtrlSetState($ftp_errorRadio3, $GUI_DISABLE)
				GUICtrlSetState($btnftp_errorPotvrdit, $GUI_DISABLE)
				GUICtrlSetState($btnftp_errorStorno, $GUI_DISABLE)
				;move parse to final
				FileMove($piaacpath & $tc & ".parsed", $outputpath & $tc & ".csv",1)
				;create export pack
				RunWait('"' & "c:\progra~1\vmware\vmware player\zip.exe" & '"' & " -q " & '"' & @DesktopDir & "\" & "export_" & $tc & "_" & $ftp_error_time & ".zip" & '" ' & "*.zip " & $tc & ".csv", $outputpath, @SW_HIDE)
				exitloop
			endif
			if(bitand(GUICtrlRead($ftp_errorRadio3), $GUI_CHECKED)) then
				GUICtrlSetState($ftp_errorRadio1, $GUI_DISABLE)
				GUICtrlSetState($ftp_errorRadio2, $GUI_DISABLE)
				GUICtrlSetState($ftp_errorRadio3, $GUI_DISABLE)
				GUICtrlSetState($btnftp_errorPotvrdit, $GUI_DISABLE)
				GUICtrlSetState($btnftp_errorStorno, $GUI_DISABLE)
				;move parse to final
				FileMove($piaacpath & $tc & ".parsed", $outputpath & $tc & ".csv",1)
				;create full export pack
				RunWait('"' & "c:\progra~1\vmware\vmware player\zip.exe" & '"' & " -q " & '"' & @DesktopDir & "\" & "export_" & $tc & "_" & $ftp_error_time & ".zip" & '" ' & "*.zip " & "*.ogg " & "*.jpg " & $tc & ".csv", $outputpath, @SW_HIDE)
				;export ogg , jpg only once
				$oggmovearray = _FilelistToArray($outputpath,"*.ogg",1)
				$jpgmovearray = _FilelistToArray($outputpath,"*.jpg",1)
				for $k = 1 to Ubound($oggmovearray) - 1
					FileMove($outputpath & $oggmovearray[$k], $outputpath & $oggmovearray[$k] & ".exported")
				next
				for $l = 1 to Ubound($jpgmovearray) - 1
					FileMove($outputpath & $jpgmovearray[$l], $outputpath & $jpgmovearray[$l] & ".exported")
				next
				exitloop
			endif
		EndIf
	wend
	;clean after self
	FileDelete($outputpath & $tc & ".csv")
	GUIDelete($winFtp_error)
	cmsvisibility(1)
EndFunc
