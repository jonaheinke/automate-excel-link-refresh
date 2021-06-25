#SingleInstance, Force

; hold Ctrl+Pause | Ctrl+ScrollLock
CtrlBreak::
	Pause, On
	return

#IfWinActive ahk_class XLMAIN
!F5::
	ControlGetFocus, focus, ahk_class XLMAIN
	if (focus = "EXCEL72") {
		MsgBox, Aktiviere Bearbeitung.
		return
	}
	Loop, 11 {			
		send, {Alt down}v{Alt up}v
		WinWait, ahk_class bosa_sdm_XL9
		limit := A_Index - 1
		send, {Down %limit%}{Tab 3}{Enter}

		sleep, 1000
		send, !{F4}
		;WinClose, A
		sleep, 500

		;WinWaitActive, ahk_class XLMAIN,,,Admanager
		;MsgBox large found
		;WinClose ,;20_21.xlsx - Excel
		;sleep, 2000
		;WinClose, ahk_class XLMAIN,,,Admanager ;, Excel, ExcludeTitle:=Admanager
		;WinWaitClose
	}
	MsgBox Done!
	return