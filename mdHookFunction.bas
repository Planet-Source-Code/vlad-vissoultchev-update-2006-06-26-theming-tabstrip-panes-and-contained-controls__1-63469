Attribute VB_Name = "mdHookImportedFunctionByName"
Option Explicit

'--- will Debug.Print module imports
#Const SHOW_MODULE_IMPORTS = False

'=========================================================================
' API
'=========================================================================

Private Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES   As Long = 16
Private Const IMAGE_DIRECTORY_ENTRY_IMPORT       As Long = 1 ' Import Directory
Private Const IMAGE_ORDINAL_FLAG32               As Long = &H80000000
Private Const PAGE_READWRITE                     As Long = &H4
Private Const VER_PLATFORM_WIN32_NT              As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualQuery Lib "kernel32" (lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Type IMAGE_IMPORT_DESCRIPTOR
    OriginalFirstThunk      As Long
    TimeDateStamp           As Long
    ForwarderChain          As Long
    Name                    As Long
    FirstThunk              As Long
End Type

Private Type IMAGE_DOS_HEADER
    e_magic                 As Integer
    e_cblp                  As Integer
    e_cp                    As Integer
    e_crlc                  As Integer
    e_cparhdr               As Integer
    e_minalloc              As Integer
    e_maxalloc              As Integer
    e_ss                    As Integer
    e_sp                    As Integer
    e_csum                  As Integer
    e_ip                    As Integer
    e_cs                    As Integer
    e_lfarlc                As Integer
    e_ovno                  As Integer
    e_res(0 To 3)           As Integer
    e_oemid                 As Integer
    e_oeminfo               As Integer
    e_res2(0 To 9)          As Integer
    e_lfanew                As Long
End Type

Private Type IMAGE_FILE_HEADER
    Machine                 As Integer
    NumberOfSections        As Integer
    TimeDateStamp           As Long
    PointerToSymbolTable    As Long
    NumberOfSymbols         As Long
    SizeOfOptionalHeader    As Integer
    Characteristics         As Integer
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress          As Long
    Size                    As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic                   As Integer
    MajorLinkerVersion      As Byte
    MinorLinkerVersion      As Byte
    SizeOfCode              As Long
    SizeOfInitializedData   As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint     As Long
    BaseOfCode              As Long
    BaseOfData              As Long
    ImageBase               As Long
    SectionAlignment        As Long
    FileAlignment           As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion       As Integer
    MinorImageVersion       As Integer
    MajorSubsystemVersion   As Integer
    MinorSubsystemVersion   As Integer
    Win32VersionValue       As Long
    SizeOfImage             As Long
    SizeOfHeaders           As Long
    CheckSum                As Long
    Subsystem               As Integer
    DllCharacteristics      As Integer
    SizeOfStackReserve      As Long
    SizeOfStackCommit       As Long
    SizeOfHeapReserve       As Long
    SizeOfHeapCommit        As Long
    LoaderFlags             As Long
    NumberOfRvaAndSizes     As Long
    DataDirectory(0 To IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_NT_HEADERS
    Signature               As Long
    FileHeader              As IMAGE_FILE_HEADER
    OptionalHeader          As IMAGE_OPTIONAL_HEADER
End Type

Private Type IMAGE_THUNK_DATA32
    FunctionOrOrdinalOrAddress As Long
End Type

Private Type MEMORY_BASIC_INFORMATION
    BaseAddress             As Long
    AllocationBase          As Long
    AllocationProtect       As Long
    RegionSize              As Long
    State                   As Long
    Protect                 As Long
    lType                   As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize     As Long
    dwMajorVersion          As Long
    dwMinorVersion          As Long
    dwBuildNumber           As Long
    dwPlatformID            As Long
    szCSDVersion            As String * 128      '  Maintenance string for PSS usage
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)

End Sub

'=========================================================================
' Functions
'=========================================================================

Public Function HookImportedFunctionByName( _
            ByVal hModule As Long, _
            szImportMod As String, _
            szImportFunc As String, _
            ByVal pFuncAddress As Long, _
            pOrigAddress As Long) As Boolean
    Const FUNC_NAME     As String = "HookImportedFunctionByName"
    Dim udtDesc         As IMAGE_IMPORT_DESCRIPTOR
    Dim pOrigThunk      As Long
    Dim pRealThunk      As Long
    Dim udtOrigThunk    As IMAGE_THUNK_DATA32
    Dim udtRealThunk    As IMAGE_THUNK_DATA32
    Dim sBuffer         As String
    Dim udtMem          As MEMORY_BASIC_INFORMATION
    Dim lOldProtect     As Long
    Dim lNotUsed        As Long

    On Error GoTo EH
    '--- parameters check
    If hModule = 0 Or pFuncAddress = 0 Or szImportMod = "" Or szImportFunc = "" Then
        Exit Function
    End If
    '--- dll above 2G on 9x -> NOT working!!!!
    If hModule < 0 Then
        If Not pvIsNT() Then
            Exit Function
        End If
    End If
    '--- get Import Descriptor
    If Not pvGetNamedImportDescriptor(hModule, szImportMod, udtDesc) Then
        Exit Function
    End If
    '--- guard offset
    If udtDesc.FirstThunk = 0 Or udtDesc.OriginalFirstThunk = 0 Then
        Exit Function
    End If
    '--- loop Real and Original thunks
    pOrigThunk = hModule + udtDesc.OriginalFirstThunk
    pRealThunk = hModule + udtDesc.FirstThunk
    '--- dereference Original Thunk
    CopyMemory udtOrigThunk, ByVal pOrigThunk, LenB(udtOrigThunk)
    Do While udtOrigThunk.FunctionOrOrdinalOrAddress <> 0
        '--- check if imported by name
        If (udtOrigThunk.FunctionOrOrdinalOrAddress And IMAGE_ORDINAL_FLAG32) = 0 Then
#If SHOW_MODULE_IMPORTS Then
            sBuffer = String(1024, 0)
            lstrcpy sBuffer, hModule + udtOrigThunk.FunctionOrOrdinalOrAddress + 2
            Debug.Print Left(sBuffer, InStr(1, sBuffer, Chr(0)))
#End If
            '--- case-insensitive compare
            If lstrcmpi(szImportFunc, hModule + udtOrigThunk.FunctionOrOrdinalOrAddress + 2) = 0 Then
                '--- set read/write access to pRealThunk
                VirtualQuery ByVal pRealThunk, udtMem, LenB(udtMem)
                If VirtualProtect(ByVal udtMem.BaseAddress, udtMem.RegionSize, PAGE_READWRITE, lOldProtect) = 0 Then
                    '--- ooops!
                    Exit Function
                End If
                '--- save orig func address and change to our func address
                CopyMemory udtRealThunk, ByVal pRealThunk, LenB(udtRealThunk)
                pOrigAddress = udtRealThunk.FunctionOrOrdinalOrAddress
                udtRealThunk.FunctionOrOrdinalOrAddress = pFuncAddress
                CopyMemory ByVal pRealThunk, udtRealThunk, LenB(udtRealThunk)
                '--- restore protection
                VirtualProtect ByVal udtMem.BaseAddress, udtMem.RegionSize, lOldProtect, lNotUsed
                '--- success
                HookImportedFunctionByName = True
                Exit Function
            End If
        Else
            Debug.Print "."
        End If
        '--- check next thunks
        pOrigThunk = pOrigThunk + LenB(udtOrigThunk)
        pRealThunk = pRealThunk + LenB(udtRealThunk)
        CopyMemory udtOrigThunk, ByVal pOrigThunk, LenB(udtOrigThunk)
    Loop
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'= private ===============================================================

Private Function pvGetNamedImportDescriptor( _
            ByVal hModule As Long, _
            szImportMod As String, _
            udtDesc As IMAGE_IMPORT_DESCRIPTOR) As Boolean
    Dim udtDosHeader    As IMAGE_DOS_HEADER
    Dim udtNtHeaders    As IMAGE_NT_HEADERS
    Dim udtImportDesc   As IMAGE_IMPORT_DESCRIPTOR
    Dim pImportDesc     As Long
    
    On Error Resume Next
    '--- dereference DOS Header
    CopyMemory udtDosHeader, ByVal hModule, LenB(udtDosHeader)
    '--- dereference NT Header
    CopyMemory udtNtHeaders, ByVal hModule + udtDosHeader.e_lfanew, LenB(udtNtHeaders)
    '--- check if any imports
    If udtNtHeaders.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress = 0 Then
        Exit Function
    End If
    '--- loop and dereference Import Descriptions
    pImportDesc = hModule + udtNtHeaders.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress
    CopyMemory udtImportDesc, ByVal pImportDesc, LenB(udtImportDesc)
    Do While udtImportDesc.Name <> 0
        '--- case-insensitive compare
        If lstrcmpi(szImportMod, hModule + udtImportDesc.Name) = 0 Then
            udtDesc = udtImportDesc
            '--- success
            pvGetNamedImportDescriptor = True
            Exit Function
        End If
        '--- dereference next Import Descriptions in the array
        pImportDesc = pImportDesc + LenB(udtImportDesc)
        CopyMemory udtImportDesc, ByVal pImportDesc, LenB(udtImportDesc)
    Loop
End Function

Private Function pvIsNT() As Boolean
    Dim udtVer As OSVERSIONINFO
    
    On Error Resume Next
    udtVer.dwOSVersionInfoSize = Len(udtVer)
    If GetVersionEx(udtVer) Then
        If udtVer.dwPlatformID = VER_PLATFORM_WIN32_NT Then
            pvIsNT = True
        End If
    End If
End Function

