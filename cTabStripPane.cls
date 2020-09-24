VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabStripPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "cTabStripPane"

'=========================================================================
' API
'=========================================================================

'--- window msgs
Private Const WM_SIZE                   As Long = &H5
Private Const WM_PAINT                  As Long = &HF
Private Const WM_KEYFIRST               As Long = &H100
Private Const WM_KEYLAST                As Long = &H108
Private Const WM_CTLCOLORBTN            As Long = &H135
Private Const WM_CTLCOLORSTATIC         As Long = &H138
Private Const WM_MOUSEFIRST             As Long = &H200
Private Const WM_LBUTTONUP              As Long = &H202
Private Const WM_RBUTTONUP              As Long = &H205
Private Const WM_MOUSELAST              As Long = &H209
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_PRINT                  As Long = &H317
Private Const WM_PRINTCLIENT            As Long = &H318
'--- for DrawThemeBackground
Private Const TABP_BODY                 As Long = 10
'--- for GetStockObject
Private Const NULL_BRUSH                As Long = 5
'--- for RedrawWindow
Private Const RDW_INVALIDATE            As Long = &H1
Private Const RDW_ALLCHILDREN           As Long = &H80
Private Const RDW_ERASENOW              As Long = &H200
'--- for logbrush style
Private Const BS_SOLID                  As Long = 0
'--- for ExtTextOut
Private Const ETO_OPAQUE                As Long = 2
'--- for Get/SetBkMode
Private Const TRANSPARENT               As Long = 1

Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Private Declare Function OpenThemeData Lib "UxTheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "UxTheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "UxTheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, ByVal pClipRect As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetBkMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ExtTextOutW Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, ByVal lpRect As Long, ByVal lpString As Long, ByVal nCount As Long, ByVal lpDx As Long) As Long

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type POINTAPI
    X                   As Long
    Y                   As Long
End Type

Private Type LOGBRUSH
    lbStyle             As Long
    lbColor             As Long
    lbHatch             As Long
End Type

'=========================================================================
' COnstants and member variables
'=========================================================================

Private m_hWndTab           As Long
Private m_hWndContainer     As Long
Private m_bFrame            As Long
Private m_cPanes            As Collection
Private m_cEdits            As Collection
Private m_hWndPaint         As Long

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

Property Get hWndTab() As Long
    hWndTab = m_hWndTab
End Property

Property Get hWndContainer() As Long
    hWndContainer = m_hWndContainer
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init( _
            ByVal hWndTab As Long, _
            oContainer As Object, _
            oControls As Object) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oCtl            As Object
    Dim vElem           As Variant
    
    On Error GoTo EH
    '--- init member vars
    m_hWndTab = hWndTab
    m_hWndContainer = oContainer.hWnd
    m_bFrame = TypeOf oContainer Is VB.Frame
    If IsComCtl6Loaded() Then
        '--- first collect contained panes
        Set m_cPanes = New Collection
        For Each oCtl In oControls
            If TypeOf oCtl Is VB.PictureBox Or TypeOf oCtl Is VB.Frame Then
                If oCtl.Container Is oContainer Then
                    m_cPanes.Add InitTabStripPane(hWndTab, oCtl, oControls), "#" & oCtl.hWnd
                End If
            End If
        Next
        If m_cPanes.Count = 0 Then
            Set m_cPanes = Nothing
        End If
        '--- next collect edit controls
        Set m_cEdits = EnumChildEdits(m_hWndContainer)
        For Each vElem In m_cEdits
            Call SetWindowSubclass(vElem, AddressOf RedirectTabPaneEditWndProc, 0, ObjPtr(Me))
        Next
        If m_cEdits.Count = 0 Then
            Set m_cEdits = Nothing
        End If
        '--- subclass current pane and install api hooks
        Call SetWindowSubclass(m_hWndContainer, AddressOf RedirectTabPaneTabWndProc, 0, ObjPtr(Me))
        HookApiCalls
    End If
    '--- success
    Init = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Sub Terminate()
    Const FUNC_NAME     As String = "Terminate"
    Dim oPane           As cTabStripPane
    Dim vElem           As Variant
    
    On Error GoTo EH
    '--- check if already terminated
    If m_hWndTab = 0 Then
        Exit Sub
    End If
    If IsComCtl6Loaded() Then
        '--- terminate contained panes
        If Not m_cPanes Is Nothing Then
            For Each oPane In m_cPanes
                oPane.Terminate
            Next
            Set m_cPanes = Nothing
        End If
        '--- unsubclass edit controls
        If Not m_cEdits Is Nothing Then
            For Each vElem In m_cEdits
                Call RemoveWindowSubclass(vElem, AddressOf RedirectTabPaneEditWndProc, 0)
            Next
            Set m_cEdits = Nothing
        End If
        '--- unsubclass current pane and remove api hooks
        Call RemoveWindowSubclass(m_hWndContainer, AddressOf RedirectTabPaneTabWndProc, 0)
        UnhookApiCalls
    End If
    m_hWndTab = 0
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'= friend ================================================================

Friend Function frFillRectImpl( _
            ByVal hDCCtl As Long, _
            ByVal lpRect As Long, _
            ByVal hBrush As Long) As Long
    Const FUNC_NAME     As String = "frFillRectImpl"
    Dim rc              As RECT
    Dim lb              As LOGBRUSH
    
    On Error Resume Next
    If hBrush <> 0 Then
        '--- check if brush is solid vbButtonFace color
        lb.lbStyle = -1
        Call GetObject(hBrush, Len(lb), lb)
        If lb.lbStyle = BS_SOLID Then
            If lb.lbColor = GetSysColor(vbButtonFace And &HFF) Then
                hBrush = 0
            End If
        End If
    End If
    If hBrush = 0 Then
        '--- paint using tab strip "body" style
        Call CopyMemory(rc, ByVal lpRect, Len(rc))
        pvPaintCtlBackground m_hWndTab, m_hWndPaint, rc, hDCCtl
        '--- success
        frFillRectImpl = 1
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Friend Function frExtTextOutW( _
            ByVal hDC As Long, _
            ByVal X As Long, _
            ByVal Y As Long, _
            ByVal wOptions As Long, _
            ByVal lpRect As Long, _
            ByVal lpString As Long, _
            ByVal nCount As Long, _
            ByVal lpDx As Long) As Long
    Const FUNC_NAME     As String = "frExtTextOutW"
    Dim lMode           As Long
    
    On Error GoTo EH
    If GetBkColor(hDC) = GetSysColor(vbButtonFace And &HFF) Then
        lMode = GetBkMode(hDC)
        Call SetBkMode(hDC, TRANSPARENT)
        If (wOptions And ETO_OPAQUE) <> 0 Then
            frFillRectImpl hDC, lpRect, 0
            wOptions = wOptions And (Not ETO_OPAQUE)
        End If
    End If
    frExtTextOutW = ExtTextOutW(hDC, X, Y, wOptions, lpRect, lpString, nCount, lpDx)
    If lMode <> 0 Then
        Call SetBkMode(hDC, lMode)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Friend Function frWndProc( _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
    Const FUNC_NAME     As String = "frWndProc"
    Dim rc              As RECT
    
    On Error GoTo EH
    Select Case wMsg
    Case WM_CTLCOLORBTN, WM_CTLCOLORSTATIC
        If IsThemed() Then
            '--- paint themed background and return null brush to contained control
            Select Case VBGetClassName(lParam)
            Case "ThunderCommandButton", "ThunderOptionButton", "ThunderCheckBox", _
                    "ThunderRT6CommandButton", "ThunderRT6OptionButton", "ThunderRT6CheckBox"
                pvPaintCtlBackground m_hWndTab, lParam, rc, wParam
                frWndProc = GetStockObject(NULL_BRUSH)
                Exit Function
            End Select
        End If
    Case WM_PAINT, WM_PRINT
        If IsThemed() Then
            '--- intercept FillRect on container background painting
            '---   and use tab pane themed background instead
            m_hWndPaint = hWnd
            Set CurrentTabStripPane = Me
            frWndProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
            Set CurrentTabStripPane = Nothing
            Exit Function
        End If
    Case WM_PRINTCLIENT, WM_MOUSELEAVE
        If IsThemed() Then
            '--- fix VB.Frame bugs w/ XP themes
            If m_bFrame Then
                frWndProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
                Exit Function
            End If
        End If
    Case WM_SIZE
        If IsThemed() Then
            '--- force repaint on contained controls if tab pane background is themed gradient
            Call RedrawWindow(hWnd, 0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_ERASENOW)
        End If
    End Select
    frWndProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Friend Function frEditWndProc( _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
    Const FUNC_NAME     As String = "frEditWndProc"
    
    On Error GoTo EH
    If wMsg = WM_PAINT _
            Or (wMsg >= WM_KEYFIRST And wMsg <= WM_KEYLAST) _
            Or (wMsg >= WM_MOUSEFIRST And wMsg <= WM_MOUSELAST And wMsg <> WM_LBUTTONUP And wMsg <> WM_RBUTTONUP) Then
        If IsThemed() Then
            m_hWndPaint = hWnd
            Set CurrentTabStripPane = Me
            frEditWndProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
            Set CurrentTabStripPane = Nothing
            Exit Function
        End If
    End If
    frEditWndProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'= private ===============================================================

Private Sub pvPaintCtlBackground( _
            ByVal hWndTab As Long, _
            ByVal hWndCtl As Long, _
            rcClip As RECT, _
            ByVal hDCCtl As Long)
    Const FUNC_NAME     As String = "pvPaintCtlBackground"
    Dim hTheme          As Long
    Dim rc              As RECT
    Dim hDC             As Long
    Dim pt              As POINTAPI
    
    On Error GoTo EH
    '--- size tab control "client" area
    Call GetClientRect(hWndTab, rc)
    '--- figure out dc
    hDC = hDCCtl
    If hDCCtl = 0 Then
        hDC = GetDC(hWndCtl)
    End If
    Call GetWindowOrgEx(hDC, pt)
    Call OffsetRect(rc, -pt.X, -pt.Y)
    '--- map "client" area to child control coordinates
    Call MapWindowPoints(hWndTab, hWndCtl, rc, 2)
    '--- use theme to draw tabctl pane
    hTheme = OpenThemeData(hWndTab, StrPtr("TAB"))
    Call DrawThemeBackground(hTheme, hDC, TABP_BODY, 0, rc, IIf(rcClip.Left <> rcClip.Right, VarPtr(rcClip), 0))
    Call CloseThemeData(hTheme)
    '--- cleanup
    If hDCCtl = 0 Then
        Call ReleaseDC(hWndCtl, hDC)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    Terminate
End Sub
