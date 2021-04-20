#NoEnv
#SingleInstance Force
SetWorkingDir %A_ScriptDir%
if not A_IsAdmin
	Run *RunAs "%A_ScriptFullPath%"

Global CurrentDate
	, currentVersion := "20210413_01"
	, versionSource := "https://raw.githubusercontent.com/rambertm/diymanual/master/BE"
FormatTime, CurrentDate,, yyyyMMdd
Global EMRPath := "C:\RMX_2.0\EXE\RMXLDR.EXE"
	, filePath := "C:\Users\" . A_UserName . "\Documents\Rounding" . CurrentDate
	, xlFile := filePath . ".xlsx"
	, wordFileVert := filepath . "(Vert).docx"
	, txtFile := filepath . ".txt"
	, currentYear := Floor(SubStr(CurrentDate, 1, 4))
	, currentMonth := Floor(SubStr(CurrentDate, 5, 2))
	, currentDay := Floor(SubStr(CurrentDate, 7, 2))
    , UpdateProgress
	, callnum1
	, callnum2
	, targetRowArray := []
	, ptNumBTArr := []
	, ptNumDrainArr := []
	, feverCountDay := 2
	, ppAlignLeft := 1
	, ppAlignCenter := 2
	, ppAlignRight := 3
	, myGUI
	, isBE
	, isBEStatus
	, ptMemoX := 875
	, ptMemoY := 570
	, initialExcelX := 620
	, initialExcelY := 560
	, DrNameArray := []

Gui, New, +HwndmyGUI, BE_DailyAuto_FA
Gui, Add, Text, x20 y25 w120 h22, % "version: " currentVersion
Gui, Add, Progress, x200 y10 w80 h10 cB9DCFA vUpdateProgress
Gui, Add, Button, x200 y20 w80 h22 gUpdate, UPDATE

Gui, Add, Button, x10 y55 w120 h25 gisBECheck, 진료과 체크 설정
Gui, Add, Text, x220 y62 w45 h25 visBEStatus,
Gui, Add, Text, x10 y165 w70 h22, 주치의:
Gui, Add, Edit, x65 y160 w70 h22 vDrName
Gui, Add, Button, x145 y160 w35 h24 gDrNameInput, 추가
Gui, Add, Button, x190 y160 w45 h24 gDrNameReset, 초기화
Gui, Add, Button, x240 y160 w70 h24 gDrNameCheck, 등록상태

Func_CheckRegistry()
Sleep, 1000
Func_ClearEnvironment()
Sleep, 10000
Func_RunEMR()
Sleep, 10000
Func_RSGenerate()
ExitApp

isBECheck:
{
	RegRead, isBE, HKCU, Software\BEDailyAutoFA, isBE
	if isBE {
		RegWrite, REG_BINARY, HKCU\Software\BEDailyAutoFA, isBE, 00
		GuiControl, +c000000, isBEStatus
		GuiControl, ,isBEStatus, 해제됨
	}else{
		RegWrite, REG_BINARY, HKCU\Software\BEDailyAutoFA, isBE, 01
		GuiControl, +cDA4F49, isBEStatus
		GuiControl, ,isBEStatus, 설정됨
	}
	return
}

Update:
{
	GuiControl, %myGUI%:, UpdateProgress, 10
	Sleep, 250
	oHttp := ComObjCreate("WinHttp.Winhttprequest.5.1")
	oHttp.open("GET", versionSource)
	oHttp.send()
	rawData = % oHttp.responseText
	versionDataEnd := InStr(rawData, "]BE_AUTO_VERSION", 1)
	versionDataStart := InStr(rawData, "BE_AUTO_VERSION[", 1)
	versionData := SubStr(rawData, versionDataStart + 16, versionDataEnd - versionDataStart - 16)
	GuiControl, %myGUI%:, UpdateProgress, 30
	Sleep, 250
	if (versionData == currentVersion){
		GuiControl, %myGUI%:, UpdateProgress, 70
		Sleep, 150
		GuiControl, %myGUI%:, UpdateProgress, 100
		Sleep, 100
		GuiControl, %myGUI%:, UpdateProgress, 0
		MsgBox, 최신 버전입니다.
	}else{
		GuiControl, %myGUI%:, UpdateProgress, 50
		dlPath = https://github.com/rambertm/diymanual/raw/master/BE_AUTO_%versionData%.exe?raw=true
		versionFileName = BE_AUTO_%versionData%.exe
		IfNotExist, %versionFileName%
			URLDownloadToFile, %dlPath%, %versionFileName%
		GuiControl, %myGUI%:, UpdateProgress, 100
		Sleep, 100
		GuiControl, %myGUI%:, UpdateProgress, 0
		IfExist, "beautorenamer.bat"
			FileDelete, "beautorenamer.bat"
		MsgBox, 업데이트 완료
		FileAppend,
		(
			Del %A_ScriptName%
			ren %versionFileName% BE_AUTO.exe
			Del beautorenamer.bat
		), beautorenamer.bat
		Run beautorenamer.bat
		ExitApp
	}
	return
}


DrNameInput:
{
	Gui, Submit, noHide
	if (DrName <> ""){
		current_DrName := ""
		RegRead, current_DrName, HKCU, Software\BEDailyAutoFA, DrName
		if ErrorLevel = 1
		{
			RegWrite, REG_SZ, HKCU\Software\BEDailyAutoFA, DrName, % DrName
		}else{
			RegWrite, REG_SZ, HKCU\Software\BEDailyAutoFA, DrName, % current_DrName . "," . DrName
		}
		MsgBox,,알림, % "주치의: " . DrName . " 등록되었습니다."
		DrName := ""
		GuiControl,,DrName,
	}
	return
}

DrNameReset:
{
	DrName := ""
	GuiControl,,DrName,
	RegDelete, HKCU\Software\BEDailyAutoFA, DrName
	return
}

DrNameCheck:
{
	current_DrName := ""
	RegRead, current_DrName, HKCU, Software\BEDailyAutoFA, DrName
	if ErrorLevel = 1
	{
		MsgBox,,알림, % "현재 등록되어 있는 주치의가 없습니다.`n모든 환자리스트가 출력됩니다."
	}else{
		MsgBox,,알림, % "현재 등록되어 있는 주치의: " . current_DrName
	}
	return
}

^+d::
{
    if (!WinExist("ahk_id " myGUI)){
		Gui %myGUI%: Show
		Pause
	}
	return
}

Func_RSGenerate(){
	WinActivate ahk_class TRMXMAINF
	Sleep, 250
	WinActivate ahk_class TMDP600F1
	Sleep, 1000
	MouseClick, left, 155, 75, 1, 0 ; 병동환자
	Sleep, 3000
	if isBE {
		MouseClick, left, 200, 125, 1, 0 ; 진료과
		Sleep, 3000
	}
	WinGetPos, , , ptSelect_W, ptSelect_H, ahk_class TMDP600F1
	if (ptSelect_W < 930){
		MsgBox, , 경고, 통합환자선택화면의 너비가 좁습니다.`n'환자메모'가 보이도록 창을 늘린 후 다시 시도해주세요.
		return
	}
	ptMemoX := 875
	ptMemoY := Floor(ptSelect_H - 35)
	initialExcelX := 620
	initialExcelY := Floor(ptSelect_H - 50)
	WinActivate ahk_class TMDP600F1
	MouseClick, left, ptMemoX, ptMemoY, 1, 0
	Sleep, 500
	WinActivate ahk_class TMDP600F1
	MouseClick, left, initialExcelX, initialExcelY, 1, 0
	Func_checkExcelReason()
	Sleep, 20000

	Xl := Excel_Get()
	Wb := Xl.ActiveWorkBook
	St := Wb.ActiveSheet

	columnMemo := 8

	Xl.range(Xl.rows(1),Xl.rows(3)).delete
	Xl.Application.ActiveWindow.FreezePanes := False
	numEndRow := St.UsedRange.Rows.Count
	Xl.range(Xl.rows(numEndRow),Xl.rows(numEndRow)).delete
	Xl.range(Xl.columns("E"),Xl.columns("G")).delete
	Xl.range(Xl.columns("G"),Xl.columns("H")).delete
	Xl.range(Xl.columns("H"),Xl.columns("U")).delete

	rowCount := 1
	numEndRow := St.UsedRange.Rows.Count
	While rowCount <= numEndRow
	{
		ptMemoVal := Xl.cells(rowCount, 8).value
		if (InStr(ptMemoVal, "drain") or Func_isOPAdmin(ptMemoVal)){
			ptNumDrainArr[rowCountDrain] := rowCount
			rowCountDrain++
		}
		ptNumBTArr[rowCount] := Xl.cells(rowCount, 3).value
		ptRoomRawValue := Xl.cells(rowCount, 1).value
		val_isICU := SubStr(ptRoomRawValue, 4, 1)
		isICU := false
		isEW := false
		If ((val_isICU = "S") or (val_isICU = "M") or (val_isICU = "C") or (val_isICU = "E") or (val_isICU = "N")){
			isICU := true
		}
		IfInString, ptRoomRawValue, EW
			isEW := true
		If isICU {
			Xl.cells(rowCount, 1).value := Substr(ptRoomRawValue, -1)
		}else if isEW {
			Xl.cells(rowCount, 1).value := Substr(ptRoomRawValue, -2)
		}else{
			if (ptRoomRawValue > 11000){
				Xl.cells(rowCount, 1).value := Substr(ptRoomRawValue, -3)
			}
			else{
				Xl.cells(rowCount, 1).value := Substr(ptRoomRawValue, -2)
			}
		}
		Xl.cells(rowCount,columnMemo).value := StrReplace(Xl.cells(rowCount, columnMemo).value, "`r", "")
		rowCount++
	}
	St.Columns(columnMemo).NumberFormat := General
	St.Columns(columnMemo).ColumnWidth := 45
	St.Columns(columnMemo).VerticalAlignment:= -4160
	St.Columns(columnMemo).HorizontalAlignment:=-4131

	Sleep, 5000
	WinActivate ahk_class TRMXMAINF
	Sleep, 250
	WinActivate ahk_class TMDP600F1
	Sleep, 250
	WinClose ahk_class TMDP600F1
	Sleep, 250
	WinActivate ahk_class TRMXMAINF
	Sleep, 250
	
	;;;;;;;;;;; V/S check module ;;;;;;;;;;;
	MouseClick, left, 280, 40, 1, 0 ;처방결과 조회
	Sleep, 250
	Loop, 12
		{
			Send, {Down}
			Sleep, 200
		}	
	Sleep, 250
	Send, {Right}
	Sleep, 250
	Send, {Down}
	Sleep, 250
	Send, {Enter}
	Sleep, 250
	WinActivate, ahk_class TMDF120F2
	WinMove, ahk_class TMDF120F2, , 0, 0
	Sleep, 250
	needleRegEx = (\d\d:\d\d)
	FormatTime, TodayDate,, yyMMdd
	Loop, % numEndRow
	{
		vsData := ""
		Clipboard := ""
		p := 1, md := ""
		feverEventCount := 0
		msgDate := ""
		ptNumBT := ptNumBTArr[A_Index]
		MouseClick, Left, 80, 57, 2, 0
		Sleep, 500
		SendInput, % ptNumBT
		Sleep, 250
		Send, {Enter}
		Sleep, 500
		MouseClick, Left, 50, 155, 1, 0
		Send, ^a
		Send, ^c
		Sleep, 250
		vsData := Clipboard
		Sleep, 250
		while p := RegExMatch(vsData, needleRegEx, md, p + StrLen(md))
		{
			vsCheckDate := SubStr(vsData, p - 7, 6)
			vsCheckDateNum := Floor(vsCheckDate)
			gapDate := TodayDate - vsCheckDateNum, Days
			if (gapDate > feverCountDay){
				break
			}
			btPos := InStr(vsData, "`t", , p, 6)
			vsBT := Round(SubStr(vsData, btPos + 1, 4), 1)
			if (vsBT > 37.9) and (vsBT < 43.0)
			{
				vsCheckTime := md
				if (msgDate <> ""){
					msgDate := msgDate . "`n"
				}
				vsCheckDateMo := SubStr(vsCheckDate, 3, 2)
				vsCheckDateDay := SubStr(vsCheckDate, 5, 2)
				msgDate := msgDate . vsCheckDateMo  . "/" . vsCheckDateDay . " " . vsCheckTime . "   " . vsBT
				feverEventCount++
				if (feverEventCount = 2){
					break
				}
			}
		}
		if (msgDate <> ""){
			Xl.cells(A_Index, 12).value := msgDate
		}
	}
	Sleep, 500
	WinClose ahk_class TMDF120F2
	Sleep, 250
	;;;;;;;;;;; End of V/S check module ;;;;;;;;;;;
	
	;;;;;;;;;;; Drain Check module ;;;;;;;;;;;
	WinActivate ahk_class TRMXMAINF
	Sleep, 250
	MouseClick, left, 280, 40, 1, 0 ;처방결과 조회
	Sleep, 250
	Loop, 12
		{
			Send, {Down}
			Sleep, 100
		}	
	Sleep, 150
	Send, {Right}
	Sleep, 100
	Send, {Down}
	Sleep, 100
	Send, {Down}
	Sleep, 100
	Send, {Enter}
	Sleep, 500
	WinActivate ahk_class TMDF130F1
	WinMove, ahk_class TMDF130F1, , 0, 0
	Sleep, 250
	needleRegEx = (: )(\d?\d?\d?)(,)
	Loop, % ptNumDrainArr.length()
	{
		drainData := ""
		Clipboard := ""
		md := ""
		msgDate := "`nDrain: "
		ptNumDrain := Xl.cells(ptNumDrainArr[A_Index], 3).value
		WinActivate ahk_class TMDF130F1
		MouseClick, Left, 85, 48, 2, 0
		Sleep, 500
		SendInput, % ptNumDrain
		Sleep, 250
		Send, {Enter}
		Sleep, 500
		MouseClick, Left, 426, 578, 1, 0
		Sleep, 750
		WinActivate ahk_class TMDF130F2		;;I/O기록창
		Sleep, 250
		Send, % currentYear
		Sleep, 100
		SendInput, {Right}
		Sleep, 100
		Send, % currentMonth
		Sleep, 100
		SendInput, {Right}
		Sleep, 100
		Send, % currentDay - 1
		Sleep, 350
		MouseClick, Left, 410, 335, 1, 0	;;시간탭 클릭
		Sleep, 250
		MouseClick, Left, 410, 355, 1, 0
		Sleep, 100
		MouseClick, Right, 410, 355, 1, 0	;;첫번째 줄 클릭
		Sleep, 100
		SendInput, {Down}
		Sleep, 100
		SendInput, {Enter}
		Sleep, 100
		drainData := Clipboard
		if (RegExMatch(drainData, needleRegEx, md, 1)){
			msgDate := msgDate . md2
			isNextDrain := false
			IfInString, drainData, Lt
			{
				msgDate := msgDate . "(Lt)"
				isNextDrain := true
			}
			IfInString, drainData, Rt
			{
				msgDate := msgDate . "(Rt)"
				isNextDrain := true
			}
			IfInString, drainData, (1)
			{
				msgDate := msgDate . "(1)"
				isNextDrain := true
			}
			if isNextDrain
			{
				drainData := ""
				Clipboard := ""
				md := ""
				isNextDrain := false
				MouseClick, Left, 410, 375, 1, 0
				Sleep, 100
				MouseClick, Right, 410, 375, 1, 0	;;두번째 줄 클릭
				Sleep, 100
				SendInput, {Down}
				Sleep, 100
				SendInput, {Enter}
				Sleep, 100
				drainData := Clipboard
				if (RegExMatch(drainData, needleRegEx, md, 1)){
					msgDate := msgDate . ", " . md2
				}
				IfInString, drainData, Lt
				{
					msgDate := msgDate . "(Lt)"
				}
				IfInString, drainData, Rt
				{
					msgDate := msgDate . "(Rt)"
				}
				IfInString, drainData, (2)
				{
					msgDate := msgDate . "(2)"
					isNextDrain := true
				}
				if isNextDrain
				{
					isNextDrain := false
					drainData := ""
					Clipboard := ""
					MouseClick, Left, 410, 395, 1, 0
					Sleep, 100
					MouseClick, Right, 410, 395, 1, 0	;;세번째 줄 클릭
					Sleep, 100
					SendInput, {Down}
					Sleep, 100
					SendInput, {Enter}
					Sleep, 100
					drainData := Clipboard
					IfInString, drainData, (3)
					{
						if (RegExMatch(drainData, needleRegEx, md, 1)){
							msgDate := msgDate . ", " . md2 . "(3)"
						}
					}
				}
			}
			prevData := ""
			prevData := Xl.cells(ptNumDrainArr[A_Index], 12).value
			if (prevData = ""){
				Xl.cells(ptNumDrainArr[A_Index], 12).value := msgDate
			}else{
				Xl.cells(ptNumDrainArr[A_Index], 12).value := prevData . "`n" . msgDate
			}
		}
		Sleep, 250
		WinClose ahk_class TMDF130F2
	}
	Sleep, 250
	WinClose ahk_class TMDF130F1
	Sleep, 250
	;;;;;;;;;;; End of Drain Check module ;;;;;;;;;;;
	
	
	IfExist, %xlFile%
		FileDelete, %xlFile%
	Sleep, 500
	Wb.SaveAs(xlFile)
	Wb.Close(false)
	Xl.Quit
	Sleep, 250
	
	WinActivate ahk_class TRMXMAINF
	Sleep, 250
	MouseClick, left, 280, 40, 1, 0 ;처방결과 조회
	Sleep, 250
	Loop, 15
		{
			Send, {Down}
			Sleep, 250
		}	
	Sleep, 250
	Send, {Enter}
	Sleep, 5000
	WinActivate ahk_class TMMR020F1
	Sleep, 250
	MouseClick, left, 20, 55, 1, 0 ;small LU box
	Sleep, 1500
	MouseClick, left, 185, 80, 1, 0	;진료과
	Sleep, 1000
	MouseClick, left, 495, 50, 1, 0 ;적용
	Sleep, 1500
	MouseClick, left, 210, 50, 1, 0 ;진료과
	Sleep, 1500
	MouseClick, left, 905, 50, 1, 0	;엑셀
	Func_checkExcelReason()
	Sleep, 15000

	Xl := Excel_Get()
	Wb := Xl.ActiveWorkBook
	St := Wb.ActiveSheet

	columnCBCAfter := 10
	columnLabAfter := 11
	columnLabName := 2

	numEndRow := St.UsedRange.Rows.Count - 1
	rowWBC := Func_LabRowCheck(Xl, numEndRow, "WBC count")
	rowHb := Func_LabRowCheck(Xl, numEndRow, "Hb")
	rowPlt := Func_LabRowCheck(Xl, numEndRow, "Platelet count")
	rowANC := Func_LabRowCheck(Xl, numEndRow, "Neutrophil count")
	rowGlu := Func_LabRowCheck(Xl, numEndRow, "Glucose (FBS, 응급)")
	rowCa := Func_LabRowCheck(Xl, numEndRow, "Calcium, total (응급)")
	rowP := Func_LabRowCheck(Xl, numEndRow, "Phosphorus (응급)")
	rowUA := Func_LabRowCheck(Xl, numEndRow, "Uric acid (응급)")
	rowBUN := Func_LabRowCheck(Xl, numEndRow, "BUN (응급)")
	rowCr := Func_LabRowCheck(Xl, numEndRow, "Creatinine (serum)")
	rowAlb := Func_LabRowCheck(Xl, numEndRow, "Albumin (응급)")
	rowAST := Func_LabRowCheck(Xl, numEndRow, "AST (응급)")
	rowALT := Func_LabRowCheck(Xl, numEndRow, "ALT (응급)")
	rowGGT := Func_LabRowCheck(Xl, numEndRow, "GGT(응급)")
	rowALP := Func_LabRowCheck(Xl, numEndRow, "ALP (응급)")
	rowCK := Func_LabRowCheck(Xl, numEndRow, "CK (응급)")
	rowLD := Func_LabRowCheck(Xl, numEndRow, "LD (응급)")
	rowAmyl := Func_LabRowCheck(Xl, numEndRow, "Amylase (응급)")
	rowNa := Func_LabRowCheck(Xl, numEndRow, "Na (응급)")
	rowK := Func_LabRowCheck(Xl, numEndRow, "K (응급)")
	rowCl := Func_LabRowCheck(Xl, numEndRow, "Cl (응급)")
	rowCRP := Func_LabRowCheck(Xl, numEndRow, "CRP(응급실용)")
	rowPTH := Func_LabRowCheck(Xl, numEndRow, "PTH, intact")

	arrPtNum := []
	arrLabCBC := []
	arrLabAbn := []

	columnLabCount := 4
	numEndColumn := St.UsedRange.Columns.Count
	While columnLabCount <= numEndColumn
	{
		labCBC := Func_GetCBC(Xl, rowWBC, rowHb, rowPlt, rowANC, columnLabCount)
		labGlu := Func_GetLabData(Xl, rowGlu, 60, 126, "[ FBS: ", columnLabCount)
		labCa := Func_GetLabData(Xl, rowCa, 8.8, 10.6, "[ Ca: ", columnLabCount)
		labP := Func_GetLabData(Xl, rowP, 2.5, 4.6, "[ P: ", columnLabCount)
		labUA := Func_GetLabData(Xl, rowUA, 2.9, 7.3, "[ UA: ", columnLabCount)
		labAlb := Func_GetLabData(Xl, rowAlb, 3.8, 5.3, "[ Alb: ", columnLabCount)
		labBUN := Func_GetLabData(Xl, rowBUN, 5, 23, "[ BUN: ", columnLabCount)
		labCr := Func_GetLabRawData(Xl, rowCr, "[ Cr: ", columnLabCount)
		labAST := Func_GetLabRawData(Xl, rowAST, "[ AST: ", columnLabCount)
		labALT := Func_GetLabRawData(Xl, rowALT, "[ ALT: ", columnLabCount)
		labGGT := Func_GetLabData(Xl, rowGGT, 11, 75, "[ GGT: ", columnLabCount)
		labALP := Func_GetLabData(Xl, rowALP, 35, 105, "[ ALP: ", columnLabCount)
		labCK := Func_GetLabData(Xl, rowCK, 22, 269, "[ CK: ", columnLabCount)
		labLD := Func_GetLabData(Xl, rowLD, 0, 249, "[ LD: ", columnLabCount)
		labAmyl := Func_GetLabData(Xl, rowAmyl, 28, 100, "[ Amyl: ", columnLabCount)
		labCRP := Func_GetLabRawData(Xl, rowCRP, "[ CRP: ", columnLabCount)
		labNa := Func_GetLabData(Xl, rowNa, 135, 145, "[ Na: ", columnLabCount)
		labK := Func_GetLabData(Xl, rowK, 3.6, 5.5, "[ K: ", columnLabCount)
		labCl := Func_GetLabData(Xl, rowCl, 101, 111, "[ Cl: ", columnLabCount)
		labPTH := Func_GetLabRawData(Xl, rowPTH, "[ iPTH: ", columnLabCount)
		labAbn = %labGlu%%labCa%%labP%%labUA%
		labNext = %labBUN%%labCr%
		labAbn := Func_AddLabLineLab(labAbn, labNext)
		labNext = %labAlb%%labAST%%labALT%%labGGT%
		labAbn := Func_AddLabLineLab(labAbn, labNext)
		labNext = %labALP%%labCK%%labLD%%labAmyl%
		labAbn := Func_AddLabLineLab(labAbn, labNext)
		labNext = %labNa%%labK%%labCl%
		labAbn := Func_AddLabLineLab(labAbn, labNext)
		labAbn := Func_AddLabLineLab(labAbn, labCRP)
		labAbn := Func_AddLabLineLab(labAbn, labPTH)
		labPtCount := columnLabCount - 3
		arrPtNum[labPtCount] := Xl.cells(3, columnLabCount).value
		arrLabCBC[labPtCount] := labCBC
		arrLabAbn[labPtCount] := labAbn
		columnLabCount++
	}
	Wb.Close(False)
	Xl.Quit
	Sleep, 3000

	Xl := ComObjCreate("Excel.Application")
	Xl.Workbooks.Open(xlFile)
	Xl.Visible := true
	Wb := Xl.ActiveWorkBook
	St := Wb.ActiveSheet
	Sleep, 3000

	St.Columns(columnCBCAfter).NumberFormat := "@"
	St.Columns(columnCBCAfter).ColumnWidth := 22
	St.Columns(columnCBCAfter).VerticalAlignment:= -4160
	St.Columns(columnCBCAfter).HorizontalAlignment:=-4131
	St.Columns(columnLabAfter).NumberFormat := "@"
	St.Columns(columnLabAfter).ColumnWidth := 22
	St.Columns(columnLabAfter).VerticalAlignment:= -4160
	St.Columns(columnLabAfter).HorizontalAlignment:=-4131

	i := 1
	newEndRow := St.UsedRange.Rows.Count
	while i <= labPtCount
	{
		targetPtNum := arrPtNum[i]
		j := 1
		while j <= newEndRow
		{
			if (Xl.cells(j, 3).value == targetPtNum){
				Xl.cells(j, columnCBCAfter).value := arrLabCBC[i]
				Xl.cells(j, columnLabAfter).value := arrLabAbn[i]
				break
			}
			j++		
		}
		i++
	}
	Wb.Close(true)
	Xl.Quit
	Sleep, 3000

	WinActivate ahk_class TRMXMAINF
	WinActivate ahk_class TMMR020F1
	Sleep, 1000

	MouseClick, left, 580, 50, 1, 0 ;previous day total lab
	Sleep, 2000
	MouseClick, left, 905, 50, 1, 0	;print to excel
	Func_checkExcelReason()
	Sleep, 15000

	Xl := Excel_Get()
	Wb := Xl.ActiveWorkBook
	St := Wb.ActiveSheet

	columnCBCAfter := 9

	lengthOfArray := arrPtNum.Length()
	if (lengthOfArray > 0){
		arrPtNum.Delete(1, lengthOfArray)
	}
	lengthOfArray := arrLabCBC.Length()
	if (lengthOfArray > 0){
		arrLabCBC.Delete(1, lengthOfArray)
	}

	numEndRow := St.UsedRange.Rows.Count - 1
	rowWBC := Func_LabRowCheck(Xl, numEndRow, "WBC count")
	rowHb := Func_LabRowCheck(Xl, numEndRow, "Hb")
	rowPlt := Func_LabRowCheck(Xl, numEndRow, "Platelet count")
	rowANC := Func_LabRowCheck(Xl, numEndRow, "Neutrophil count")
	rowCr := Func_LabRowCheck(Xl, numEndRow, "Creatinine (serum)")
	rowAST := Func_LabRowCheck(Xl, numEndRow, "AST (응급)")
	rowALT := Func_LabRowCheck(Xl, numEndRow, "ALT (응급)")
	rowCRP := Func_LabRowCheck(Xl, numEndRow, "CRP(응급실용)")
	rowPTH := Func_LabRowCheck(Xl, numEndRow, "PTH, intact")

	columnLabCount := 4
	numEndColumn := St.UsedRange.Columns.Count
	While columnLabCount <= numEndColumn
	{
		labCBC := Func_GetCBC(Xl, rowWBC, rowHb, rowPlt, rowANC, columnLabCount)
		labCr := Func_GetLabRawData(Xl, rowCr, "[ Cr: ", columnLabCount)
		labAST := Func_GetLabRawData(Xl, rowAST, "[ AST: ", columnLabCount)
		labALT := Func_GetLabRawData(Xl, rowALT, "[ ALT: ", columnLabCount)
		labCRP := Func_GetLabRawData(Xl, rowCRP, "[ CRP: ", columnLabCount)
		labPTH := Func_GetLabRawData(Xl, rowPTH, "[ iPTH: ", columnLabCount)
		labPtCount := columnLabCount - 3
		arrPtNum[labPtCount] := Xl.cells(3, columnLabCount).value
		labNext = %labAST%%labALT%
		lab := Func_AddLabLineLab(labCBC, labNext)
		lab := Func_AddLabLineLab(lab, labCr)
		lab := Func_AddLabLineLab(lab, labCRP)
		lab := Func_AddLabLineLab(lab, labPTH)
		arrLabCBC[labPtCount] := lab
		columnLabCount++
	}
	Wb.Close(False)
	Xl.Quit
	Sleep, 500

	Xl := ComObjCreate("Excel.Application")
	Xl.Workbooks.Open(xlFile)
	Xl.Visible := true
	Wb := Xl.ActiveWorkBook
	St := Wb.ActiveSheet
	Sleep 3000
		
	St.Columns(columnCBCAfter).NumberFormat := "@"
	St.Columns(columnCBCAfter).ColumnWidth := 22
	St.Columns(columnCBCAfter).VerticalAlignment:= -4160
	St.Columns(columnCBCAfter).HorizontalAlignment:=-4131
		
	i := 1
	newEndRow := St.UsedRange.Rows.Count
	while i <= labPtCount
	{
		targetPtNum := arrPtNum[i]
		j := 1
		while j <= newEndRow
		{
			if (Xl.cells(j, 3).value == targetPtNum){
				Xl.cells(j, columnCBCAfter).value := arrLabCBC[i]
				break
			}
			j++		
		}
		i++
	}
	Sleep, 250
	Wb.Save
	numEndRow := St.UsedRange.Rows.Count
	roundingRoomNumbers := Func_GetRoundingRoomNumber(Xl, numEndRow)
	Func_GetDrNameArray()
	Func_GetTargetRows(Xl, numEndRow)

	WordApp := ComObjCreate("Word.Application")
	WordApp.Visible := true
	oWord := WordApp.Documents.Add
	Sleep, 250

	Func_WdPageSetup(oWord, 7, 0, 45, 30, 45, 30)
	Func_WdFormatSetup(WordApp, "굴림", 10.5, 0)

	wordHeader = RoundingSheet: %CurrentDate%`n%roundingRoomNumbers%
	Func_WdCreateCalendar(WordApp, oWord, wordHeader)
	WinActivate ahk_class OpusApp
	Send, ^{End}
	Send, {Enter}
	lengthOfTargetRowArray := targetRowArray.Length()
	objTable := Func_WdCreateTable(WordApp, lengthOfTargetRowArray, 3, 125, 125, 275)
	Func_WdFillTable(Xl, lengthOfTargetRowArray, objTable, 12, 9, 9, 10, 10, 10, 9)

	IfExist, %wordFileVert%
		FileDelete, %wordFileVert%
	Sleep, 500
	oWord.SaveAs(wordFileVert)
	Wb.Close(true)
	Xl.Quit
	SLeep, 500

	WinClose, ahk_class TRMXMAINF
	Sleep, 5000
	Send, {Enter}
	Sleep, 250
	return
}

Func_CheckRegistry(){
	RegRead, isBE, HKCU, Software\BEDailyAutoFA, isBE
	if ErrorLevel = 1
	{
		RegWrite, REG_BINARY, HKCU\Software\BEDailyAutoFA, isBE, 00
		isBE := 0
	}
	if isBE {
		GuiControl, +cDA4F49, isBEStatus
		GuiControl, ,isBEStatus, 설정됨
	}else{
		GuiControl, +c000000, isBEStatus
		GuiControl, ,isBEStatus, 해제됨
	}
	RegRead, callnum1_temp, HKCU, Software\BEDailyAutoFA, callnum1
	if (ErrorLevel = 0){
		prefix1 := SubStr(callnum1_temp, 1, 2)
		afterfix1 := SubStr(callnum1_temp, 4, 4)
		callnum1 := prefix1 afterfix1
	}else{
		MsgBox, 로그인 정보가 없습니다. 프로그램을 종료합니다.
		ExitApp
	}
	RegRead, callnum2_temp, HKCU, Software\BEDailyAutoFA, callnum2
	if (ErrorLevel = 0){
		prefix2 := SubStr(callnum2_temp, 5, 3)
		afterfix2 := SubStr(callnum2_temp, 9, 4)
		callnum2 := prefix2 afterfix2
	}else{
		MsgBox, 로그인 정보가 없습니다. 프로그램을 종료합니다.
		ExitApp
	}
	return
}

Func_ClearEnvironment(){
	Xl := ComObjCreate("Excel.Application")
	Xl.Visible := true
	Sleep, 20000
	Xl.Workbooks.Add()
	Sleep, 10000
	While WinExist("ahk_class XLMAIN")
	{
		oExcel := Excel_Get()
		oExcel.ActiveWorkbook.Close(false)
		oExcel.Quit
		Sleep, 250
	}
	Sleep, 1000
	While WinExist("ahk_class OpusApp")
	{
		oWord := Word_Get()
		oWord.Documents.Close(false)
		oWord.Quit
		Sleep, 250
	}
	return
}

Func_RunEMR(){
	Run, %EMRPath%
	Sleep, 30000
	WinActivate, ahk_class TRMXLOGF
	Sleep, 500
	SendInput, % callnum1
	Sleep, 500
	Send, {Tab}
	Sleep, 250
	SendInput, % callnum2
	Sleep, 500
	Send, {Enter}
	Sleep, 500
	Send, {Enter}
	Sleep, 500
	Send, {Enter}
	Loop, 4
	{
		Sleep, 5000
		WinClose ahk_class TMRE015F1
		WinClose ahk_class TSRA380F2
		WinClose ahk_class TMDO970F3
		WinClose ahk_class TMDC200F1
	}
	return
}

Func_LabRowCheck(Xl, endRow, searchText){
	i := 6
	while i <= endRow
	{
		if (Xl.cells(i, 2).value = searchText){
			return i
		}
		i++
	}
	return 0
}

Func_GetCBC(Xl, rowWBC, rowHb, rowPlt, rowANC, columnLabCount){
	tempWBC := Xl.cells(rowWBC, columnLabCount).value * 1000
	labWBC := Round(tempWBC)
	labHb := Xl.cells(rowHb, columnLabCount).value
	labPlt := Xl.cells(rowPlt, columnLabCount).value
	if (rowANC = 0){
		if((tempWBC = "") and (labHb = "") and (labPlt = "")){
			labCBC = 
		}
		else{
			labCBC = %labWBC%/%labHb%/%labPlt%k/(-)
		}
	}
	else{
		tempANC := Xl.cells(rowANC, columnLabCount).value
		if (tempANC = ""){
			labANC := "(-)"
			if((tempWBC = "") and (labHb = "") and (labPlt = "")){
				labCBC = 
			}
			else{
				labCBC = %labWBC%/%labHb%/%labPlt%k/%labANC%
			}
		}
		else{
			labANC := Round(tempANC * 1000)
			if((tempWBC = "") and (labHb = "") and (labPlt = "")){
				labCBC = (-)/(-)/(-)/%labANC%
			}
			else{
				labCBC = %labWBC%/%labHb%/%labPlt%k/%labANC%
			}
		}
	}
	return labCBC
}

Func_GetLabData(Xl, rowLab, labMin, labMax, labName, columnLabCount){
	if (rowLab = 0){
		labData := ""
	}
	else{
		labDataRaw := Xl.cells(rowLab, columnLabCount).value
		if ((labDataRaw != "") and ((labDataRaw<labMin) or (labDataRaw >labMax))){
			labData := labName . labDataRaw . " ]"
		}
		else{
			labData := ""
		}
	}
	return labData
}

Func_AddLabLineLab(labAbn, labNext){
	if ((labAbn != "") and (labNext != "")){
		labAbn = %labAbn%`n%labNext%
	}else{
		labAbn = %labAbn%%labNext%
	}
	return labAbn
}

Func_GetDrNameArray(){
	current_DrName := ""
	RegRead, current_DrName, HKCU, Software\BEDailyAutoFA, DrName
	if ErrorLevel = 0
	{
		Loop, parse, current_DrName, `,
		{
			DrNameArray[A_Index] := A_LoopField
		}
	}	
	return
}

Func_GetTargetRows(Xl, numEndRow){
	lengthOfArray := targetRowArray.Length()
	if (lengthOfArray > 0){
		targetRowArray.Delete(1, lengthOfArray)
	}
	targetRowCount := 1	
	floorCount := 13
	
	Loop, % numEndRow
	{
		roomNumber := Xl.cells(A_Index, 1).value
		val_isICU := SubStr(roomNumber, 1, 1)
		If (val_isICU = "S") or (val_isICU = "M") or (val_isICU = "C") or (val_isICU = "E") or (val_isICU = "N")
		{
			targetRowArray[targetRowCount] := A_Index
			targetRowCount++
		}
	}
	while floorCount > 3
	{
		currentFloor := floorCount * 100
		nextFloor := (floorCount + 1) * 100
		i := 1
		while i <= numEndRow
		{
			roomNumber := Xl.cells(i, 1).value
			if ((roomNumber > currentFloor) and (roomNumber < nextFloor)){
				drNameValue := Xl.cells(i, 5).value
				isTarget := false
				if drs := DrNameArray.length() {
					Loop, % drs
					{
						IfInString, drNameValue, % DrNameArray[A_Index]
						{
							isTarget := true
							break
						}
					}
				}else{
					isTarget := true
				}
				If isTarget
				{
					targetRowArray[targetRowCount] := i
					targetRowCount++
				}
			}
			i++
		}
		floorCount := floorCount - 1
	}
	return
}

Func_GetRoundingRoomNumber(Xl, numEndRow){
	i := numEndRow
	Arr_RoomNumbers := []
	while i > 0 
	{
		isNew := true
		if roomNumber := Floor(Xl.cells(i, 1).value)
		{
			if ArrCount := Arr_RoomNumbers.length()
			{
				if (Arr_RoomNumbers[ArrCount] == roomNumber){
				isNew := false
				}
			}
			if isNew {
				Arr_RoomNumbers.push(roomNumber)
			}
		}
		i := i - 1
	}
	n := 1
	Loop, % Arr_RoomNumbers.length() - 1
	{
		compare_L := Arr_RoomNumbers[n]
		compare_R := Arr_RoomNumbers[n+1]
		if ((compare_L < compare_R) and (n < Arr_RoomNumbers.length())){
			isChanged_F := true
			Arr_RoomNumbers[n] := compare_R
			Arr_RoomNumbers[n+1] := compare_L
		}else{
			isChanged_F := false
		}
		if isChanged_F {
			isChanged_B := true
			nn := n
			while isChanged_B
			{
				compare_L := Arr_RoomNumbers[nn-1]
				compare_R := Arr_RoomNumbers[nn]
				if ((compare_L < compare_R) and (nn > 1)){
					isChanged_B := true
					Arr_RoomNumbers[nn-1] := compare_R
					Arr_RoomNumbers[nn] := compare_L
					nn := nn - 1
				}else{
					isChanged_B := false
				}
			}
		}
		n := n + 1
	}
	
	n := 1
	roundingRoomNumbers := Arr_RoomNumbers[1]
	Loop, % Arr_RoomNumbers.length() - 1
	{
		if (Floor(Arr_RoomNumbers[n] / 100) == Floor(Arr_RoomNumbers[n+1] / 100)){
			roundingRoomNumbers := roundingRoomNumbers . ", " . Arr_RoomNumbers[n+1]
		}else{
			roundingRoomNumbers := roundingRoomNumbers . "/ " . Arr_RoomNumbers[n+1]
		}
		n++
	}
	
	Loop, % numEndRow
	{
		roomNumber := Xl.cells(A_Index, 1).value
		val_isICU := SubStr(roomNumber, 1, 1)
		If (val_isICU = "S") or (val_isICU = "M") or (val_isICU = "C") or (val_isICU = "E") or (val_isICU = "N")
		{
			IfNotInString, roundingRoomNumbers, %val_isICU%
				roundingRoomNumbers = %roundingRoomNumbers% / %val_isICU%
		}
	}
	return roundingRoomNumbers
}

Func_WdPageSetup(oWord, size, ori, TopMargin, BotMargin, LtMargin, RtMargin){
	oWord.PageSetup.PaperSize := size
	oWord.PageSetup.Orientation := ori
	oWord.PageSetup.TopMargin := TopMargin
	oWord.PageSetup.BottomMargin := BotMargin
	oWord.PageSetup.LeftMargin := LtMargin
	oWord.PageSetup.RightMargin := RtMargin
	return
}

Func_WdFormatSetup(WordApp, FontName, LineSpacing, SpaceAfter){
	WordApp.Selection.Font.Name := FontName
	WordApp.Selection.ParagraphFormat.LineSpacingRule := 1
	WordApp.Selection.ParagraphFormat.LineSpacing := LineSpacing
	WordApp.Selection.ParagraphFormat.SpaceAfter := SpaceAfter
	return
}

Func_WdCreateTable(WordApp, numEndRow, numColumns, firstColWidth, secondColWidth, thirdColWidth){
	numRows := numEndRow * 3
	tbl := WordApp.ActiveDocument.tables.Add(WordApp.Selection.Range, numRows, numColumns)
	tbl.Borders.Enable := true
	tbl.Columns(1).Width := firstColWidth
	tbl.Columns(2).Width := secondColWidth
	tbl.Columns(3).Width := thirdColWidth
	return tbl
}

Func_WdFillTable(Xl, lengthOfTargetRowArray, objTable, fs11, fs12, fs13, fs21, fs22, fs31, fs32){
	k := 1
	while k <= lengthOfTargetRowArray
	{
		percentage := Floor((k / lengthOfTargetRowArray) * 100)
		Progress, % percentage, 작업 중.. 컴퓨터 사용 중단 권장합니다., % k . " / " lengthOfTargetRowArray . " 번째 진행중...",알림
		i := targetRowArray[k]
		ptRoomNumber := Xl.cells(i, 1).value
		ptName := Xl.cells(i, 2).value
		ptNumber := Xl.cells(i, 3).value
		ptAS := Xl.cells(i, 4).value
		ptDoctor := Xl.cells(i, 5).value
		ptAdmDate := Xl.cells(i, 6).value
		ptHD := "HD # " . Xl.cells(i, 7).value
		ptMemoRaw := Xl.cells(i, 8).value
		ptYesterdayCBC := Xl.cells(i, 9).value
		if (ptYesterdayCBC != ""){
			ptYesterdayCBC = `n%ptYesterdayCBC%`n(어제)
		}
		ptTodayCBC := Xl.cells(i, 10).value
		if (ptTodayCBC != ""){
			ptTodayCBC = `n%ptTodayCBC%`n(오늘)
		}
		ptTodayLab := Xl.cells(i, 11).value
		if (ptRoomNumber > 1000){
			ptCValue = `n%ptRoomNumber%  %ptName%`n       %ptAS%
		}else{
			ptCValue = `n%ptRoomNumber%  %ptName%`n      %ptAS%
		}
		ptVSData := Xl.cells(i, 12).value
		ptMValue = `n%ptNumber%  %ptDoctor%`n%ptAdmDate%  %ptHD%
		ptLabValue = `n%ptTodayLab%
		ptMemo = `n%ptMemoRaw%
			
		sRow := (k - 1) * 3 + 1
		Func_WdTblDesign(objTable, sRow, fs11, fs12, fs13, fs21, fs22, fs31, fs32)
		objTable.Cell(sRow, 1).Range.Text := ptCValue
		objTable.Cell(sRow, 2).Range.Text := ptMValue
		objTable.Cell(sRow + 1, 1).Range.Text := ptYesterdayCBC
		objTable.Cell(sRow + 1, 2).Range.Text := ptTodayCBC
		objTable.Cell(sRow + 2, 1).Range.Text := ptVSData
		objTable.Cell(sRow + 2, 2).Range.Text := ptLabValue
		objTable.Cell(sRow, 3).Range.Text := ptMemo
		k++
		Sleep, 250
	}
	Progress, Off
	return
}

Func_WdTblDesign(tbl, sRow, fs11, fs12, fs13, fs21, fs22, fs31, fs32){
	tbl.Cell(sRow, 1).borders(3).LineStyle := 0
	tbl.Cell(sRow, 1).borders(4).LineStyle := 0
	tbl.Cell(sRow + 1, 1).borders(3).LineStyle := 0
	tbl.Cell(sRow + 1, 1).borders(4).LineStyle := 0
	tbl.Cell(sRow + 2, 1).borders(4).LineStyle := 0
	tbl.Cell(sRow, 2).borders(3).LineStyle := 0
	tbl.Cell(sRow, 2).borders(4).LineStyle := 0
	tbl.Cell(sRow + 1, 2).borders(3).LineStyle := 0
	tbl.Cell(sRow + 1, 2).borders(4).LineStyle := 0
	tbl.Cell(sRow + 2, 2).borders(4).LineStyle := 0
	
	tbl.Cell(sRow,3).Merge(tbl.cell(sRow + 2,3))
	
	tbl.Cell(sRow, 1).Range.Font.Size := fs11
	tbl.Cell(sRow, 1).Range.Font.bold := true
	tbl.Cell(sRow, 2).Range.Font.Size := fs12
	tbl.Cell(sRow, 2).Range.Font.bold := false
	tbl.Cell(sRow, 3).Range.Font.Size := fs13
	tbl.Cell(sRow, 3).Range.Font.bold := false
	tbl.Cell(sRow + 1, 1).Range.Font.Size := fs21
	tbl.Cell(sRow + 1, 1).Range.Font.bold := false
	tbl.Cell(sRow + 1, 2).Range.Font.Size := fs22
	tbl.Cell(sRow + 1, 2).VerticalAlignment := 0
	tbl.Cell(sRow + 1, 2).Range.Font.bold := true
	tbl.Cell(sRow + 2, 2).Range.Font.Size := fs32
	tbl.Cell(sRow + 2, 2).Range.Font.bold := false
	return
}

Func_WdCreateCalendar(WordApp, oWord, wordHeader){
	FormatTime, d_Today,,d
	FormatTime, m_Today,,M
	FormatTime, y_Today,,yyyy
	arr_WDname := ["일", "월", "화", "수", "목", "금", "토"]
	wd_Today := Func_GetWeekday(m_Today, d_Today, y_Today)

	cal_Table := Func_WdCreateTable_Global(WordApp, 7, 8, 30, 15)
	cal_Table.Columns(8).Width := 315
	cal_Table.Cell(1,8).Merge(cal_Table.cell(7,8))
	cal_Table.Cell(1,8).Range.Text := wordHeader
	cal_Table.Cell(1,8).VerticalAlignment := 0

	Func_WdCalendar_CellDesign(cal_Table.Cell(4, wd_Today), d_Today, 12, true, 1)
	cal_Table.Cell(4, wd_Today).Range.shading.backgroundpatterncolorindex := 16
	Loop, 7
	{
		Func_WdCalendar_CellDesign(cal_Table.Cell(1, A_Index), arr_WDname[A_Index], 10, true, 1)
		cal_Table.Cell(A_Index, 1).Range.Font.Italic := true
	}
	i := 0
	row := 4
	Loop, % wd_Today + 13
	{
		i := i - 1
		if (A_Index >= wd_Today){
			if (A_Index < wd_Today + 7){
				row := 3
				col := wd_Today - A_Index + 7
			}else{
				row := 2
				col := wd_Today - A_Index + 14
			}
		}else{
			col := wd_Today - A_Index
		}
		d_Next := Func_AddDays_ReturnDay(y_Today, m_Today, d_Today, i)
		m_Next := Func_AddDays_ReturnMonth(y_Today, m_Today, d_Today, i)
		if (m_Next < m_Today){
			Func_WdCalendar_CellDesign(cal_Table.Cell(row, col), d_Next, 9, false, 1)
			cal_Table.Cell(row, col).Range.shading.backgroundpatterncolorindex := 16
		}else{
			Func_WdCalendar_CellDesign(cal_Table.Cell(row, col), d_Next, 11, false, 1)
		}
	}
	i := 0
	row := 4
	Loop, % 28 - wd_Today
	{
		i := i + 1
		if (A_Index > 7- wd_Today){
			if (A_Index <= 14 - wd_Today){
				row := 5
				col := wd_Today + A_Index - 7
			}else if (A_Index <= 21 - wd_Today){
				row := 6
				col := wd_Today + A_Index - 14
			}else{
				row := 7
				col := wd_Today + A_Index - 21
			}
		}else{
			col := wd_Today + A_Index
		}
		d_Next := Func_AddDays_ReturnDay(y_Today, m_Today, d_Today, i)
		m_Next := Func_AddDays_ReturnMonth(y_Today, m_Today, d_Today, i)
		if (m_Next > m_Today){
			Func_WdCalendar_CellDesign(cal_Table.Cell(row, col), d_Next, 9, false, 1)
			cal_Table.Cell(row, col).Range.shading.backgroundpatterncolorindex := 16
		}else{
			Func_WdCalendar_CellDesign(cal_Table.Cell(row, col), d_Next, 11, false, 1)
		}
	}
	return
}

Func_WdCalendar_CellDesign(cell, value, fs, isBold, Align){
	cell.Range.Text := value
	cell.Range.Font.Size := fs
	cell.Range.Font.bold := isBold
	cell.Range.ParagraphFormat.Alignment := Align
	return
}

Func_WdCreateTable_Global(WordApp, numRows, numColumns, ColWidth, RowHeight){
	tbl := WordApp.ActiveDocument.tables.Add(WordApp.Selection.Range, numRows, numColumns)
	tbl.Borders.Enable := true
	Loop, % numColumns
		tbl.Columns(A_Index).Width := ColWidth
	Loop, % numRows
		tbl.Rows(A_Index).Height := RowHeight
	return tbl
}

Func_GetWeekday(month, day, year){
	d := day, m := month, y := year
    if (m < 3)
    {
        m += 12
        y -= 1
    }
    wd := mod(d + (2 * m) + floor(6 * (m + 1) / 10) + y + floor(y / 4) - floor(y / 100) + floor(y / 400) + 1, 7) + 1
	return wd
}

Func_AddDays_ReturnDay(Year, Month, Day, days2add){
	if Month <10
        Month := "0" . Month
    if Day < 10
        Day := "0" . Day
    timeStamp := Year Month Day
    EnvAdd, timeStamp, days2add, Days
    FormatTime, result, %timeStamp%, d
    return result
}

Func_AddDays_ReturnMonth(Year, Month, Day, days2add){
	if Month <10
        Month := "0" . Month
    if Day < 10
        Day := "0" . Day
    timeStamp := Year Month Day
    EnvAdd, timeStamp, days2add, Days
    FormatTime, result, %timeStamp%, M
    return result
}

Func_GetLabRawData(Xl, rowLab, labName, columnLabCount){
	if (rowLab = 0){
		labData := ""
	}
	else{
		labDataRaw := Xl.cells(rowLab, columnLabCount).value
		if (labDataRaw == ""){
			labData := ""
		}else{
			labData := labName . labDataRaw . " ]"
		}
	}
	return labData
}

Func_checkExcelReason(){
	Sleep, 500
	IfWinExist ahk_class TExCheckF
	{
		WinActivate ahk_class TExCheckF
		Loop, 29
		SendInput, {Down}
		Sleep, 250
		MouseClick, left, 205, 450, 1, 0
	}
	return
}

Func_isOPAdmin(memoCopy){
	opDate := Func_getOPDate(memoCopy)
	if (opDate){
		FormatTime, today, %A_Now%, yyyyMMdd
		EnvSub, today, %opDate%, days
		if (today < 14){
			return true
		}else{
			return false
		}
	}else{
		return false
	}
}

Func_getOPDate(memoCopy){
	if (p := RegExMatch(memoCopy, "i)op date ?:")){
	}else if(p := RegExMatch(memoCopy, "i)op ?:")){
	}
	if (p){
		nextCR := InStr(memoCopy, "`n", , p)
		if (nextCR) {
			contentOP := SubStr(memoCopy, p, nextCR - p - 1)
		}else{
			StringTrimLeft, contentOP, memoCopy, p - 1
		}
	}else{
		contentOP := ""
	}
	RegExMatch(contentOP, "(\d\d?)\/(\d\d?)\/(\d\d)", opDate)
	if (opDate){
		if (opDate3 > 50){
			opYear := "19" . opDate3
		}else{
			opYear := "20" . opDate3
		}
		if (opDate1 < 10){
			opMonth := "0" . opDate1
		}else{
			opMonth := opDate1
		}
		if (opDate2 < 10){
			opDay := "0" . opDate2
		}else{
			opDay := opDate2
		}
		result := opYear . opMonth . opDay
	}else{
		result := ""
	}
	return result
}

Excel_Get(WinTitle:="ahk_class XLMAIN", Excel7#:=1) {
    static h := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
    WinGetClass, WinClass, %WinTitle%
    if !(WinClass == "XLMAIN")
        return "Window class mismatch."
    ControlGet, hwnd, hwnd,, Excel7%Excel7#%, %WinTitle%
    if (ErrorLevel)
        return "Error accessing the control hWnd."
    VarSetCapacity(IID_IDispatch, 16)
    NumPut(0x46000000000000C0, NumPut(0x0000000000020400, IID_IDispatch, "Int64"), "Int64")
    if DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", -16, "Ptr", &IID_IDispatch, "Ptr*", pacc) != 0
        return "Error calling AccessibleObjectFromWindow."
    window := ComObject(9, pacc, 1)
    if ComObjType(window) != 9
        return "Error wrapping the window object."
    Loop
        try return window.Application
        catch e
            if SubStr(e.message, 1, 10) = "0x80010001"
                ControlSend, Excel7%Excel7#%, {Esc}, %WinTitle%
            else
                return "Error accessing the application object."
}

; Word_Get.ahk: burque505, modified from
; Excel_Get by jethrow (modified)
; Forum:    https://autohotkey.com/boards/viewtopic.php?f=6&t=31840
; Github:   https://github.com/ahkon/MS-Office-COM-Basics/blob/master/Examples/Excel/Excel_Get.ahk
; With subsequent mods by opiuetasfd - thanks!
Word_Get(WinTitle:="ahk_class OpusApp", _WwG#:=1) {
    static h := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
    WinGetClass, WinClass, %WinTitle%
    if !(WinClass == "OpusApp")
        return "Window class mismatch. (" WinClass ")"
    ControlGet, hwnd, hwnd,, _WwG%_WwG#%, %WinTitle%
    if (ErrorLevel)
        return "Error accessing the control hWnd. (" ErrorLevel ")"
    VarSetCapacity(IID_IDispatch, 16)
    NumPut(0x46000000000000C0, NumPut(0x0000000000020400, IID_IDispatch, "Int64"), "Int64")
    if (hr := DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", -16, "Ptr", &IID_IDispatch, "Ptr*", pacc)) != 0
        return "Error calling AccessibleObjectFromWindow. (" 
        . (hr = 0x80070057 ? "E_INVALIDARG" : hr = 0x80004002 ? "E_NOINTERFACE" : hr) ")"
    window := ComObject(9, pacc, 1)
    if ComObjType(window) != 9
        return "Error wrapping the window object."
    try return window.Application
    catch e
        return "Error accessing the application object. (" SubStr(e.message, 1, 10)  ")"
}
; References
;   https://autohotkey.com/board/topic/88337-ahk-failure-with-excel-get/?p=560328
;   https://autohotkey.com/board/topic/76162-excel-com-errors/?p=484371
;   https://autohotkey.com/boards/viewtopic.php?p=134048#p134048

GuiClose:
{
	ExitApp
}

^Esc::ExitApp