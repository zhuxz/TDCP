VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_config As CConfig
Private m_data As Variant
Private m_isDebug As Boolean

Private Sub Class_Initialize()
    m_isDebug = MTDCP.IsDebugApp()
End Sub

Private Sub Class_Terminate()
    m_config.Clear
    Set m_config = Nothing
End Sub

Private Sub Build(ByVal configPath As String, ByVal dataPath As String)
    'm_config.PreviewData
End Sub
