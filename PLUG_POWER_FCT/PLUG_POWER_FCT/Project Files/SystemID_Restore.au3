Global $hWnd = WinGetHandle("System ID Board Programmer") 
ToggleVisibility()
While 1
    Sleep(100)
WEnd
Func ToggleVisibility()
    If WinExists($hWnd) Then
        If WinGetState($hWnd) = 0 Then 
            WinSetState($hWnd, "", @SW_RESTORE)
        ElseIf WinGetState($hWnd) = 1 Then 
            WinSetState($hWnd, "", @SW_HIDE)
        EndIf
        WinActivate($hWnd) 
    EndIf
	Exit
EndFunc