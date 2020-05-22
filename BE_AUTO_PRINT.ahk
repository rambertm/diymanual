#NoEnv
#SingleInstance Force
SetWorkingDir %A_ScriptDir%
if not A_IsAdmin
	Run *RunAs "%A_ScriptFullPath%"

Global isPrint := 1
	, holidayArr := ["20200430", "20200501", "20200505", "20200930", "20201001", "20201002", "20201009", "20201225"]

Func_isPrint()
if isPrint
{
	WordApp := Word_Get()
	oWord := WordApp.ActiveDocument
    oWord.PrintOut(,,,,,,,1)
    return
}
ExitApp

Func_isPrint(){
	Loop, % holidayArr.Length()
	{
		if (holidayArr[A_Index] = CurrentDate){
			isPrint := 0
			break
		}
	}
	return
}

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