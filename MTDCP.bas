Attribute VB_Name = "MTDCP"
Option Explicit

Public Const UI_MARGIN = 60

Public Const SHEET_CONFIG As String = "config"

Public Const DEBUG_FILE As String = "000000"

Public Function IsDebugApp() As Boolean
    If Dir(App.Path & "\" & DEBUG_FILE) = DEBUG_FILE Then
        IsDebugApp = True
    Else
        IsDebugApp = False
    End If
End Function
