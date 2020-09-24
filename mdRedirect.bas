Attribute VB_Name = "mdRedirect"
Option Explicit

Public Function RedirectBrowseForFolderCallback( _
            ByVal This As cBrowseForFolder, _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal lParam As Long, _
            ByVal lpData As Long) As Long
    RedirectBrowseForFolderCallback = This.frCallback(hWnd, wMsg, lParam, lpData)
End Function



