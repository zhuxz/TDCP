VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTDCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isDebug As Boolean
Private m_xlsApp As Excel.Application

Private Sub Class_Initialize()
    m_isDebug = MTDCP.IsDebugApp()
End Sub

Private Function GetXLSApp() As Excel.Application
    If m_xlsApp Is Nothing Then Set m_xlsApp = MExcel.GetExcelApp()
    If m_isDebug Then
        m_xlsApp.Visible = True
    Else
        m_xlsApp.Visible = False
    End If
    Set GetXLSApp = m_xlsApp
End Function

Public Function BuildPathFile( _
    SourceData As Variant, _
    ByVal reportdate As Variant, _
    ByVal PathFileName As Variant, _
    ByVal ConfigFilePath As Variant, _
    ByVal BuilderFilePath As Variant, _
    ByRef ReturnValue As Variant, _
    Optional ByRef pSrcXML As Variant, _
    Optional ByRef pExtraXML As Variant, _
    Optional ByRef bPathfile As Boolean = False)
    
    On Error GoTo eh
    
    BuildPathFile = False
eh:
End Function

Public Function DataConfig2Pathfile(ByVal DataPath As String, ByVal ConfigPath As String, xlsWBPathfile As Excel.Workbook)
    'xlsWBData As Excel.Workbook, _
    xlsWBConfig As Excel.Workbook, xlsWBPathfile As Excel.Workbook
End Function


