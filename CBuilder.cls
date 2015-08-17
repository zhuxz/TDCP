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

Private m_dataConfig As CDataConfig

Private m_vSections As Variant
Private m_nSection As Long

Private m_data As Variant
Private m_isDebug As Boolean

Private Sub Class_Initialize()
    m_isDebug = MTDCP.IsDebugApp()
End Sub

Private Function execExp(vFuncs, funcId As Long, vInput, errDesc As String)
    Dim iArgu As Long
    Dim vFunc As Variant: vFunc = vFuncs(funcId)
    
    Select Case vFunc(ExpArgu.type_)
        Case EAType.Field
            execExp = vInput(vFunc(ExpArgu.Value))
        Case EAType.Const_
            execExp = vFunc(ExpArgu.Value)
        Case EAType.func
            If (vFunc(ExpArgu.type_) > 0) Then
            
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

Private Sub execNormalFields(configData, vInput, vOutput)
    
End Sub
'Private Sub execConditionalField()
'Private Sub execDisplayFields()
'Private Sub execBalanceField()
'Private Sub execSortOutput()
Private Sub execFakedFields(fields, vInput, vOutput)
    Dim iField As Long
    For iField = LBound(fields) To UBound(fields)
        
    Next
End Sub
'Private Sub execTHCHeadDescription()
'Private Sub execOBSItems()
'Private Sub execAcceptableValues()
'Private Sub execValueCheckFields()
'Private Sub execMSRDataSheet()

Public Sub SetConfig(oDataConfig As CDataConfig)
    Set m_dataConfig = oDataConfig
    m_vSections = oDataConfig.vSections
    m_nSection = oDataConfig.Count
End Sub

Public Function InitConfig(ByVal filePath As String) As Boolean
On Error GoTo eh
    Dim xlsApp As Excel.Application
    MExcel.LoadExcelApp xlsApp
    'Dim s As CDataConfig
    's.vItems
eh:
End Function

Public Sub Launch(vInput, vOutput)
    Dim i As Long
    Dim vSec
    
    For i = 0 To m_nSection
        vSec = m_vSections(i)
        If vSec(TCSection.type_) = TCSType.faked Then
            execFakedFields vSec(TCSection.detail), vInput, vInput
        End If
    Next
End Sub


