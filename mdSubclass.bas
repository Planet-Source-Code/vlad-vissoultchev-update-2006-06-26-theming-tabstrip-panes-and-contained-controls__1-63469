Attribute VB_Name = "mdSubclass"
Option Explicit

'==============================================================================
' API
'==============================================================================

'--- for Get/SetWindowLong
Private Const GWL_WNDPROC               As Long = -4

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type ThunkBytes
    Thunk(5)                As Long
End Type

Public Type PushParamThunk
    pfn                     As Long
    Code                    As ThunkBytes
End Type

Public Type SubClassData
    WndProcNext             As Long
    WndProcThunkThis        As PushParamThunk
    #If DEBUGWINDOWPROC Then
        dbg_Hook            As WindowProcHook
    #End If
End Type

Public Type FireOnceTimerData
    TimerID                 As Long
    TimerProcThunkData      As PushParamThunk
    TimerProcThunkThis      As PushParamThunk
End Type

Public Type WindowsHookData
    hhkNext                 As Long
    hhkThunk                As PushParamThunk
    #If DEBUGHOOKPROC Then
        dbg_Hook            As HookProcHook
    #End If
End Type

Public Sub InitPushParamThunk(Thunk As PushParamThunk, ByVal ParamValue As Long, ByVal pfnDest As Long)
'push [esp]
'mov eax, 16h // Dummy value for parameter value
'mov [esp + 4], eax
'nop // Adjustment so the next long is nicely aligned
'nop
'nop
'mov eax, 1234h // Dummy value for function
'jmp eax
'nop
'nop
    
    With Thunk.Code
        .Thunk(0) = &HB82434FF
        .Thunk(1) = ParamValue
        .Thunk(2) = &H4244489
        .Thunk(3) = &HB8909090
        .Thunk(4) = pfnDest
        .Thunk(5) = &H9090E0FF
    End With
    Thunk.pfn = VarPtr(Thunk.Code)
End Sub

Public Sub SubClass(Data As SubClassData, ByVal hwnd As Long, ByVal ThisPtr As Long, ByVal pfnRedirect As Long)
    With Data
        If .WndProcNext Then
            SetWindowLong hwnd, GWL_WNDPROC, .WndProcNext
            .WndProcNext = 0
        End If
        InitPushParamThunk .WndProcThunkThis, ThisPtr, pfnRedirect
#If DEBUGWINDOWPROC Then
        On Error Resume Next
        Set .dbg_Hook = Nothing
        Set .dbg_Hook = CreateWindowProcHook
        If Err Then
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        With .dbg_Hook
            .SetMainProc Data.WndProcThunkThis.pfn
            Data.WndProcNext = SetWindowLong(hwnd, GWL_WNDPROC, .ProcAddress)
            .SetDebugProc Data.WndProcNext
        End With
#Else
        .WndProcNext = SetWindowLong(hwnd, GWL_WNDPROC, .WndProcThunkThis.pfn)
#End If
    End With
End Sub

Public Sub UnSubClass(Data As SubClassData, ByVal hwnd As Long)
    With Data
        If .WndProcNext Then
            SetWindowLong hwnd, GWL_WNDPROC, .WndProcNext
            .WndProcNext = 0
        End If
#If DEBUGWINDOWPROC Then
        Set .dbg_Hook = Nothing
#End If
    End With
End Sub

Public Function CallNextWndProc(Data As SubClassData, ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Data.WndProcNext <> 0 Then
        CallNextWndProc = CallWindowProc(Data.WndProcNext, hwnd, wMsg, wParam, lParam)
    End If
End Function

Public Sub InitFireOnceTimer(Data As FireOnceTimerData, ByVal ThisPtr As Long, ByVal pfnRedirect As Long)
    With Data
        InitPushParamThunk .TimerProcThunkData, VarPtr(Data), pfnRedirect
        InitPushParamThunk .TimerProcThunkThis, ThisPtr, .TimerProcThunkData.pfn
        .TimerID = SetTimer(0, 0, 0, .TimerProcThunkThis.pfn)
    End With
End Sub

Public Sub TerminateFireOnceTimer(Data As FireOnceTimerData)
    With Data
        If .TimerID Then
            KillTimer 0, .TimerID
            .TimerID = 0
        End If
    End With
End Sub

'hMod and ThreadID are likely never to be used in VB because
'it isn't equipped to do global hooks (except for journal hooks
'which call back on the same thread). However, these are provided
'for completeness. If ThisPtr is 0, then pfnRedirect is passed the
'next hook procedure as the extra first parameter.
Public Sub StartWindowsHook(Data As WindowsHookData, ByVal HookType As Long, ByVal ThisPtr As Long, ByVal pfnRedirect As Long, Optional ByVal hMod As Long = -1, Optional ByVal ThreadID As Long = -1)
    With Data
        If .hhkNext Then
            UnhookWindowsHookEx .hhkNext
            .hhkNext = 0
        End If
        InitPushParamThunk .hhkThunk, ThisPtr, pfnRedirect
        If ThreadID = -1 Then ThreadID = App.ThreadID
        If hMod = -1 Then hMod = 0
#If DEBUGHOOKPROC Then
        On Error Resume Next
        Set .dbg_Hook = Nothing
        Set .dbg_Hook = New HookProcHook
        If Err Then Exit Sub
        On Error GoTo 0
        With .dbg_Hook
            .SetMainProc Data.hhkThunk.pfn
            Data.hhkNext = SetWindowsHookEx(HookType, .ProcAddress, hMod, ThreadID)
            .SetDebugHandle Data.hhkNext
        End With
#Else
        .hhkNext = SetWindowsHookEx(HookType, .hhkThunk.pfn, hMod, ThreadID)
#End If
        'If a This pointer isn't provided, then pass the next
        'hook to the callback function. Reinitializing the thunk
        'will not change its pfn value.
        If ThisPtr = 0 Then InitPushParamThunk .hhkThunk, .hhkNext, pfnRedirect
    End With
End Sub

Public Sub StopWindowsHook(Data As WindowsHookData)
    With Data
        If .hhkNext Then
            UnhookWindowsHookEx .hhkNext
            .hhkNext = 0
        End If
#If DEBUGHOOKPROC Then
        Set .dbg_Hook = Nothing
#End If
    End With
End Sub

'==============================================================================
' Sample redirectors
'==============================================================================

'Public Function RedirectControlWndProc( _
'            ByVal This As MyControl, _
'            ByVal hWnd As Long, _
'            ByVal uMsg As Long, _
'            ByVal wParam As Long, _
'            ByVal lParam As Long) As Long
'    Select Case uMsg
'    Case WM_CANCELMODE
'        This.frCancelMode
'    End Select
'    ControlWndProc = CallWindowProc(This.frGetWndProcNext, hWnd, uMsg, wParam, ByVal lParam)
'End Function

'Public Sub RedirectTimerProc( _
'            Data As FireOnceTimerData, _
'            ByVal This As Form1, _
'            ByVal hwnd As Long, _
'            ByVal wMsg As Long, _
'            ByVal idEvent As Long, _
'            ByVal dwTime As Long)
'    TerminateFireOnceTimer Data
'    This.frTimer
'End Sub

'Public Function RedirectDTCalendHookProc( _
'            ByVal This As dtpDTPicker, _
'            ByVal nCode As Long, _
'            ByVal wParam As Long, _
'            ByVal lParam As Long) As Long
'    Const FUNC_NAME     As String = "RedirectDTCalendHookProc"
'
'    On Error GoTo EH
'    RedirectDTCalendHookProc = This.frHookProcCalend(nCode, wParam, lParam)
'    Exit Function
'EH:
'    PrintError FUNC_NAME
'    Resume Next
'End Function

