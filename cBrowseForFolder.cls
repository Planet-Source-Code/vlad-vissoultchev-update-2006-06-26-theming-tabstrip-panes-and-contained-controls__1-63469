VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrowseForFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "cBrowseForFolder"

'=========================================================================
' API
'=========================================================================

Private Const BIF_RETURNONLYFSDIRS      As Long = 1
Private Const BIF_STATUSTEXT            As Long = &H4
'Private Const BIF_EDITBOX               As Long = &H10
Private Const BIF_NEWDIALOGSTYLE        As Long = &H40
Private Const MAX_PATH                  As Long = 260
Private Const WM_USER                   As Long = &H400
Private Const BFFM_SETSTATUSTEXTA       As Long = (WM_USER + 100)
Private Const BFFM_SETSELECTIONA        As Long = (WM_USER + 102)
Private Const BFFM_INITIALIZED          As Long = 1
Private Const BFFM_SELCHANGED           As Long = 2
Private Const WM_COMMAND                As Long = &H111
Private Const IDABORT                   As Long = 3

Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type BROWSEINFO
    hWndOwner      As Long
    pidlRoot       As Long
    pszDisplayName As Long
    lpszTitle      As String
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As String
    iImage         As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_DEF_TITLE         As String = "Èçáåðåòå ïàïêà"

Private m_hWndOwner             As Long
Private m_sBaseFolder           As String
Private m_sTitle                As String
Private m_sFolder               As String
Private m_uThunk                As PushParamThunk
Private m_hWnd                  As Long
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Error management
'=========================================================================

Private Sub RaiseError(sFunc As String)
    Err.Raise Err.Number, MODULE_NAME & "." & sFunc & vbCrLf & Err.Source, Err.Description
End Sub


Private Sub PrintError(sFunc As String)
    Call OutputDebugString(MODULE_NAME & "." & sFunc & ": " & Err.Description & Timer & vbCrLf)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description & Timer
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get BaseFolder() As String
    BaseFolder = m_sBaseFolder
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init( _
            ByVal hWndOwner As Long, _
            sBaseFolder As String, _
            sTitle As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    
    On Error GoTo EH
    m_hWndOwner = hWndOwner
    m_sBaseFolder = sBaseFolder
    #If DebugMode Then
        DebugInstanceName m_sBaseFolder, m_sDebugID
    #End If
    m_sTitle = IIf(Len(sTitle) > 0, sTitle, STR_DEF_TITLE)
    InitPushParamThunk m_uThunk, ObjPtr(Me), AddressOf RedirectBrowseForFolderCallback
    '--- success
    Init = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function ShowSelect(sFolder As String) As Boolean
    Const FUNC_NAME     As String = "ShowSelect"
    Dim uBrowseInfo     As BROWSEINFO
    Dim lPIDL           As Long
    Dim sTemp           As String

    On Error GoTo EH
    With uBrowseInfo
        .hWndOwner = m_hWndOwner
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT Or BIF_NEWDIALOGSTYLE
        .lParam = MergePaths(m_sBaseFolder, sFolder)
        #If DebugMode Then
            DebugInstanceName .lParam, m_sDebugID
        #End If
        .lpszTitle = m_sTitle
        .lpfnCallback = m_uThunk.pfn
    End With
    lPIDL = SHBrowseForFolder(uBrowseInfo)
    m_hWnd = 0
    If lPIDL <> 0 Then
        sTemp = String(MAX_PATH + 1, 0)
        Call SHGetPathFromIDList(lPIDL, sTemp)
        If InStr(sTemp, Chr(0)) > 0 Then
            sFolder = PathDifference(m_sBaseFolder, Left(sTemp, InStr(sTemp, Chr(0)) - 1))
            #If DebugMode Then
                DebugInstanceName Left(sTemp, InStr(sTemp, Chr(0)) - 1), m_sDebugID
            #End If
        End If
        Call CoTaskMemFree(lPIDL)
        '--- success
        ShowSelect = True
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PathDifference( _
            sBase As String, _
            sFolder As String) As String
    Const FUNC_NAME     As String = "PathDifference"
    Dim vBase           As Variant
    Dim vFolder         As Variant
    Dim lIdx            As Long
    Dim lJ              As Long
    
    On Error GoTo EH
    If LCase(Left(sBase, 2)) <> LCase(Left(sFolder, 2)) Then
        PathDifference = sFolder
    Else
        vBase = Split(sBase, "\")
        vFolder = Split(sFolder, "\")
        For lIdx = 0 To UBound(vFolder)
            If lIdx <= UBound(vBase) Then
                If LCase(vBase(lIdx)) <> LCase(vFolder(lIdx)) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        Next
        If lIdx > UBound(vBase) Then
            PathDifference = "."
        Else
            For lJ = lIdx To UBound(vBase)
                PathDifference = PathDifference & IIf(Len(PathDifference) > 0, "\", "") & ".."
            Next
        End If
        For lJ = lIdx To UBound(vFolder)
            PathDifference = PathDifference & IIf(Len(PathDifference) > 0, "\", "") & vFolder(lJ)
        Next
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function MergePaths( _
            sBase As String, _
            sFolder As String) As String
    Const FUNC_NAME     As String = "MergePaths"
    
    On Error GoTo EH
    If Mid(sFolder, 2, 1) = ":" Or Left(sFolder, 2) = "\\" Then
        MergePaths = sFolder
    ElseIf Left(sFolder, 1) = "\" Then
        MergePaths = Left(sBase, 2) & sFolder
    Else
        MergePaths = sBase & IIf(Len(sBase) > 0, "\", "") & sFolder
    End If
    MergePaths = CanonicalPath(MergePaths)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function CanonicalPath(sPath As String) As String
    On Error Resume Next
    CanonicalPath = sPath
    CanonicalPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(sPath)
    On Error GoTo 0
End Function

'= friend ================================================================

Friend Function frCallback( _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal lParam As Long, _
            ByVal lpData As Long) As Long
    Const FUNC_NAME     As String = "frCallback"
    Dim sTemp           As String
    
    On Error GoTo EH
    Select Case wMsg
    Case BFFM_INITIALIZED
        m_hWnd = hWnd
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, 1, ByVal lpData)
        Screen.MousePointer = vbDefault
    Case BFFM_SELCHANGED
        sTemp = String(MAX_PATH + 1, 0)
        If SHGetPathFromIDList(lParam, sTemp) Then
            Call SendMessage(hWnd, BFFM_SETSTATUSTEXTA, 0, ByVal sTemp)
        End If
    End Select
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'=========================================================================
' Base class events
'=========================================================================

#If DebugMode Then
    Private Sub Class_Initialize()
        DebugInstanceInit MODULE_NAME, m_sDebugID, Me
    End Sub
#End If

#If DebugMode Then
    Private Sub Class_Terminate()
        DebugInstanceTerm MODULE_NAME, m_sDebugID
    End Sub
#End If

'=========================================================================
' IUcsSelector interface
'=========================================================================

Private Function IUcsSelector_GetNextItem(CurrentValue As Variant, CurrentText As String, Result As Variant) As Boolean
    If Len(m_sFolder) > 0 Then
        CurrentValue = m_sFolder
        CurrentText = m_sFolder
        '--- success
        IUcsSelector_GetNextItem = True
    End If
End Function

Private Function IUcsSelector_QuickFind(CurrentValue As Variant, CurrentText As String, ByVal MultiSelect As Boolean, Result As Variant) As Boolean
    m_sFolder = CurrentText
    IUcsSelector_QuickFind = True
End Function

Private Function IUcsSelector_Reset() As Boolean
    IUcsSelector_Reset = True
End Function

Private Function IUcsSelector_ShowSelect(CurrentValue As Variant, CurrentText As String, ByVal MultiSelect As Boolean, Result As Variant) As Boolean
    m_sFolder = CurrentText
    IUcsSelector_ShowSelect = ShowSelect(m_sFolder)
End Function

'=========================================================================
' IUcsTerminate interface
'=========================================================================

Private Sub IUcsTerminate_QueryTerminate(Cancel As Boolean)

End Sub

Private Sub IUcsTerminate_Terminate()
    If m_hWnd <> 0 Then
        Call PostMessage(m_hWnd, WM_COMMAND, IDABORT, 0)
        m_hWnd = 0
    End If
End Sub


