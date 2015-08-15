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

Public Function InitConfig(ByVal filePath As String) As Boolean
On Error GoTo eh
    Dim xlsApp As Excel.Application
    MExcel.LoadExcelApp xlsApp
    
    'Dim s As CDataConfig
    's.vItems
eh:
End Function

Public Sub Launch(config, vInput, vOutPut)
    Dim i As Long
    For i = 0 To UBound(config)
        
    Next
End Sub

Private Function execExp(vInput, vFuncs, funcId As Long, errDesc As String)
    Dim iArgu As Long
    Dim vFunc As Variant: vFunc = vFuncs(funcId)
    
    Select Case vFunc(ExpArgu.Type_)
        Case EAType.Field
            execExp = vInput(vFunc(ExpArgu.Value))
        Case EAType.Const_
            execExp = vFunc(ExpArgu.Value)
        Case EAType.func
            If (vFunc(ExpArgu.Type_) > 0) Then
            
            Else
                'execExp = ExecFunction(vFunc(ExpArgu.FuncAddr), Empty, errDesc)
            End If
    End Select
End Function

Private Function execExpression(vExp)
    Dim ret
    If IsArray(vExp) Then
        
    Else
        'ret = Exp
    End If
End Function

Private Sub execNormalFields(configData, vInput, vOutPut)
    
End Sub
'Private Sub execConditionalField()
'Private Sub execDisplayFields()
'Private Sub execBalanceField()
'Private Sub execSortOutput()
Private Sub execFakedFields(configData, vInput)

End Sub
'Private Sub execTHCHeadDescription()
'Private Sub execOBSItems()
'Private Sub execAcceptableValues()
'Private Sub execValueCheckFields()
'Private Sub execMSRDataSheet()


