Attribute VB_Name = "mdTabStripPaneHelper"
Option Explicit
Private Const MODULE_NAME As String = "mdTabStripPaneHelper"

'=========================================================================
' API
'=========================================================================

'--- for Get/SetThemeAppProperties
Private Const STAP_ALLOW_CONTROLS       As Long = 2
'--- for InitCommonControlsEx
Private Const ICC_USEREX_CLASSES        As Long = &H200

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Private Declare Function IsAppThemed Lib "uxtheme" () As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Long
Private Declare Function GetThemeAppProperties Lib "uxtheme" () As Long
Private Declare Function DllGetVersion Lib "comctl32.dll" (pdvi As DLLVERSIONINFO) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'--- hooked
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByVal lpRect As Long, ByVal hBrush As Long) As Long
Private Declare Function ExtTextOutW Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, ByVal lpRect As Long, ByVal lpString As Long, ByVal nCount As Long, ByVal lpDx As Long) As Long

Private Type DLLVERSIONINFO
    cbSize              As Long
    dwMajor             As Long
    dwMinor             As Long
    dwBuildNumber       As Long
    dwPlatformID        As Long
End Type

Private Type tagInitCommonControlsEx
   lngSize              As Long
   lngICC               As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const MODULE_COMCTL32           As String = "COMCTL32.DLL"
Private Const MODULE_COMCT232           As String = "MSCOMCTL.OCX"
Private Const MODULE_MSVBVM60           As String = "MSVBVM60.DLL"
Private Const MODULE_VB6                As String = "VB6.EXE"
Private Const MODULE_USER32             As String = "USER32.DLL"
Private Const MODULE_GDI32              As String = "GDI32.DLL"
Private Const MODULE_USP10              As String = "USP10.DLL" '--- Uniscribe Unicode script processor
Private Const API_FILLRECT              As String = "FillRect"
Private Const API_EXTTEXTOUTW           As String = "ExtTextOutW"

Private m_oFwdTabStripPane          As cTabStripPane
Private m_lHookRefCount             As Long
Private m_pOrigFillRect             As Long
Private m_pOrigExtTextOutW          As Long

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Call OutputDebugString(MODULE_NAME & "." & sFunc & ": " & Err.Description & Timer & vbCrLf)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description & Timer
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get CurrentTabStripPane() As cTabStripPane
    Set CurrentTabStripPane = m_oFwdTabStripPane
End Property

Property Set CurrentTabStripPane(oValue As cTabStripPane)
    Set m_oFwdTabStripPane = oValue
End Property

'=========================================================================
' Redirectors
'=========================================================================

Public Function RedirectTabPaneTabWndProc( _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long, _
            ByVal uIdSubclass As Long, _
            ByVal This As cTabStripPane) As Long
    #If uIdSubclass Then '--- touch args
    #End If
    RedirectTabPaneTabWndProc = This.frWndProc(hWnd, wMsg, wParam, lParam)
End Function

Public Function RedirectTabPaneEditWndProc( _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long, _
            ByVal uIdSubclass As Long, _
            ByVal This As cTabStripPane) As Long
    #If uIdSubclass Then '--- touch args
    #End If
    RedirectTabPaneEditWndProc = This.frEditWndProc(hWnd, wMsg, wParam, lParam)
End Function

'=========================================================================
' Class factories
'=========================================================================

Public Function InitTabStripPane( _
            ByVal hWndTab As Long, _
            oContainer As Object, _
            oControls As Object, _
            Optional RetVal As cTabStripPane) As cTabStripPane
    Set RetVal = New cTabStripPane
    If RetVal.Init(hWndTab, oContainer, oControls) Then
        Set InitTabStripPane = RetVal
    End If
End Function

'=========================================================================
' Global Functions
'=========================================================================

Public Function InitCommonControlsVB() As Boolean
   Dim iccex            As tagInitCommonControlsEx
   
   On Error Resume Next
   Call LoadLibrary("shell32.dll")
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   Call InitCommonControlsEx(iccex)
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Public Function HookApiCalls() As Boolean
    Const FUNC_NAME     As String = "HookApiCalls"
        
    On Error GoTo EH
    If m_lHookRefCount = 0 Then
        m_pOrigFillRect = pvHookApiFunc(MODULE_USER32, API_FILLRECT, AddressOf pvCustomFillRect)
        m_pOrigExtTextOutW = pvHookApiFunc(MODULE_GDI32, API_EXTTEXTOUTW, AddressOf pvCustomExtTextOutW)
    End If
    m_lHookRefCount = m_lHookRefCount + 1
    '--- success
    HookApiCalls = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Sub UnhookApiCalls()
    Const FUNC_NAME     As String = "UnhookApiCalls"
        
    On Error GoTo EH
    If m_lHookRefCount > 0 Then
        m_lHookRefCount = m_lHookRefCount - 1
        If m_lHookRefCount = 0 Then
            pvHookApiFunc MODULE_USER32, API_FILLRECT, m_pOrigFillRect
            pvHookApiFunc MODULE_GDI32, API_EXTTEXTOUTW, m_pOrigExtTextOutW
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Function IsComCtl6Loaded() As Boolean
    Const FUNC_NAME     As String = "IsComCtl6Loaded"
    Dim uVer            As DLLVERSIONINFO
    
    On Error GoTo EH
    uVer.cbSize = Len(uVer)
    Call DllGetVersion(uVer)
    IsComCtl6Loaded = (uVer.dwMajor >= 6)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function IsThemed() As Boolean
    '--- uxtheme.dll is not present on earlier OS'es
    On Error Resume Next
    IsThemed = True
    If IsAppThemed() = 0 Then
        IsThemed = False
    ElseIf IsThemeActive() = 0 Then
        IsThemed = False
    ElseIf (GetThemeAppProperties() And STAP_ALLOW_CONTROLS) = 0 Then
        IsThemed = False
    End If
    On Error GoTo 0
End Function

Public Function EnumChildEdits(ByVal hWnd As Long) As Collection
    Set EnumChildEdits = New Collection
    Call EnumChildWindows(hWnd, AddressOf pvEnumChildEdits, VarPtr(EnumChildEdits))
End Function

Public Function VBGetClassName(ByVal hWnd As Long) As String
    VBGetClassName = String(1000, 0)
    Call GetClassName(hWnd, VBGetClassName, Len(VBGetClassName) - 1)
    VBGetClassName = Left$(VBGetClassName, InStr(VBGetClassName, Chr$(0)) - 1)
End Function

'= private ===============================================================

Private Function pvHookApiFunc(sModule As String, sFunc As String, ByVal lAddr As Long) As Long
    Const FUNC_NAME     As String = "pvHookApiFunc"
    Dim lPrev           As Long
    
    On Error GoTo EH
    '--- sanity check
    If lAddr = 0 Then Exit Function
    HookImportedFunctionByName GetModuleHandle(MODULE_COMCTL32), sModule, sFunc, lAddr, lPrev
    If lPrev <> 0 Then pvHookApiFunc = lPrev
    HookImportedFunctionByName GetModuleHandle(MODULE_MSVBVM60), sModule, sFunc, lAddr, lPrev
    If lPrev <> 0 Then pvHookApiFunc = lPrev
    HookImportedFunctionByName GetModuleHandle(MODULE_VB6), sModule, sFunc, lAddr, lPrev
    If lPrev <> 0 Then pvHookApiFunc = lPrev
    HookImportedFunctionByName GetModuleHandle(MODULE_COMCT232), sModule, sFunc, lAddr, lPrev
    If lPrev <> 0 Then pvHookApiFunc = lPrev
    HookImportedFunctionByName GetModuleHandle(MODULE_USP10), sModule, sFunc, lAddr, lPrev
    If lPrev <> 0 Then pvHookApiFunc = lPrev
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvCustomFillRect(ByVal hDC As Long, ByVal lpRect As Long, ByVal hBrush As Long) As Long
    Const FUNC_NAME     As String = "pvCustomFillRect"
    
    pvInitVbRuntime
    On Error GoTo EH
    If Not CurrentTabStripPane Is Nothing And lpRect <> 0 Then
        pvCustomFillRect = CurrentTabStripPane.frFillRectImpl(hDC, lpRect, hBrush)
    End If
    If pvCustomFillRect = 0 Then
        pvCustomFillRect = FillRect(hDC, lpRect, hBrush)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvCustomExtTextOutW(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, ByVal lpRect As Long, ByVal lpString As Long, ByVal nCount As Long, ByVal lpDx As Long) As Long
    Const FUNC_NAME     As String = "pvCustomExtTextOutW"
    
    pvInitVbRuntime
    On Error GoTo EH
    If Not CurrentTabStripPane Is Nothing Then
        pvCustomExtTextOutW = CurrentTabStripPane.frExtTextOutW(hDC, X, Y, wOptions, lpRect, lpString, nCount, lpDx)
    End If
    If pvCustomExtTextOutW = 0 Then
        pvCustomExtTextOutW = ExtTextOutW(hDC, X, Y, wOptions, lpRect, lpString, nCount, lpDx)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvEnumChildEdits(ByVal hWnd As Long, oCol As Collection) As Long
    Select Case VBGetClassName(hWnd)
    Case "Edit", "ThunderTextBox", "ThunderRT6TextBox"
        oCol.Add hWnd
    End Select
    pvEnumChildEdits = 1
End Function

Private Sub pvInitVbRuntime()
    Dim IID_IUnknown    As VBGUID
    Dim CLSID_Dummy     As VBGUID
    Dim pUnk            As IUnknown
    
    '--- create an object
    IID_IUnknown = GUIDFromString("{00000000-0000-0000-C000-000000000046}")
    CLSID_Dummy = CLSIDFromProgID("FolderWatcher.cDummy")
    Call CoCreateInstance(CLSID_Dummy, Nothing, CLSCTX_INPROC_SERVER, IID_IUnknown, pUnk)
End Sub
