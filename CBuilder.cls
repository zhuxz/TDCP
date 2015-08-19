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

Private m_vDataHeads As Variant
Private m_nDataHead As Long

Private m_vOutputHeads As Variant
Private m_nOutputHead As Long

Private m_isDebug As Boolean

Private Sub Class_Initialize()
    m_isDebug = MTDCP.IsDebugApp()
End Sub

Private Function execExp(vFuncs, funcId As Long, vInput, errDesc As String)
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
    
    m_vDataHeads = oDataConfig.DataHeads
    m_nDataHead = oDataConfig.DataHeadCount
    
    m_vOutputHeads = oDataConfig.OutputHeads
    m_nOutputHead = oDataConfig.OutputHeadCount
End Sub

Public Function InitConfig(ByVal filePath As String) As Boolean
On Error GoTo eh
    Dim xlsApp As Excel.Application
    MExcel.LoadExcelApp xlsApp
    'Dim s As CDataConfig
    's.vItems
eh:
End Function

Public Sub LaunchXLSRow(vInput, vOutput, Optional ByVal rowId As Long)
    Dim iCol As Long
    Dim dataVals(): ReDim dataVals(1 To m_nDataHead)
    For iCol = 1 To UBound(vInput)
        If IsError(vInput(iCol)) Then
            dataVals(iCol) = MExcel.GetExcelErrorValue(CStr(vInput(iCol)))
            
        Else
            dataVals(iCol) = vInput
        End If
    Next
    
    Dim vSec
    Dim iSec As Long
    
    'calc faked fields
    For iSec = 0 To m_nSection
        vSec = m_vSections(iSec)
        If vSec(TCSection.Type_) = TCSType.faked Then
            execFakedFields vSec(TCSection.detail), dataVals, dataVals
        End If
    Next
End Sub


