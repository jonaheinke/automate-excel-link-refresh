#SingleInstance Force

CtrlBreak:: ;hold Ctrl+Pause | Ctrl+ScrollLock
	Pause On
	Return

#IfWinActive ahk_class XLMAIN
!F5::
	title = Excel Link Refresh
	references_count := 11
	Loop {
		WinActivate ahk_class XLMAIN
		ControlGetFocus focus, ahk_class XLMAIN
		If (focus = "EXCEL72") {
			MsgBox 5, %title%, Warte auf Aktivierung der Bearbeitung für das Spreadsheet..., 3
			IfMsgBox Cancel
				Return
		}
		Else
			Break
	}
	Loop %references_count% {
		Send !vv
		;WinWait ahk_class bosa_sdm_XL9
		limit := A_Index - 1
		Send {Down %limit%}{Tab 3}{Enter}
		Sleep 1000
		If WinActive("ahk_class XLMAIN") ;doesn't work sadly
			Send !{F4}
		Sleep 400
	}
	MsgBox 0, %title%, Fertig! %references_count% Verknüpfungen wurden aktualisiert., 3
	Return



/*
Dokumentation zu MsgBox:
https://ahkde.github.io/docs/commands/MsgBox.htm
*/

/*
Potentiell bessere Alternativen zu "Send !{F4}":

WinWaitActive, ahk_class XLMAIN,,,Admanager
MsgBox large found
WinClose ,;20_21.xlsx - Excel
Sleep, 2000
WinClose, ahk_class XLMAIN,,,Admanager ;, Excel, ExcludeTitle:=Admanager
WinClose, A
WinWaitClose
*/