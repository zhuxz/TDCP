VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lastErr As String

Private m_srcData As Variant
Private m_srcDataProp As Variant

Private m_allHeads() As Variant
Private m_headCount As Variant

Private m_msrs() As Variant
Private m_msrCount As Long

Private m_configs As Variant
Private m_configCount As Long

Private m_curConfig As Long
Private m_curSection As Long

Private Sub Class_Initialize()
    m_headCount = 0
    m_msrCount = 0
    m_configCount = 0
End Sub

Private Function AppendCommonTHCHead(ByVal rowId As Long) As Long
    Dim head(1) As String
    Dim startCol As Long: startCol = 1
    head(0) = Trim$(m_srcData(rowId)(startCol))
    head(1) = Trim$(m_srcData(rowId)(startCol + 1))
    
    If m_headCount = 0 Then
        ReDim m_allHeads(m_headCount) As Variant
    Else
        ReDim Preserve m_allHeads(m_headCount) As Variant
    End If
    
    m_allHeads(m_headCount) = head
    AppendCommonTHCHead = m_headCount
    
    m_headCount = m_headCount + 1
End Function

Private Function AppendMSRConfig(ByVal rowId As Long) As Long
    Dim dataName As String
    Dim startCol As Long: startCol = 1
    dataName = Trim$(m_srcData(rowId)(startCol))
    
    If m_msrCount = 0 Then
        ReDim m_msrs(m_msrCount) As Variant
    Else
        ReDim Preserve m_msrs(m_msrCount) As Variant
    End If
    
    m_msrs(m_msrCount) = dataName
    AppendMSRConfig = m_msrCount
    
    m_msrCount = m_msrCount + 1
End Function

Private Function AppendConfig(ByVal rowId As Long) As Long
    Dim conf(TConfig.BOF_ + 1 To TConfig.EOF_ - 1) As Variant
    Dim startCol As Variant: startCol = 1
    
    conf(TConfig.srcName) = Trim$(m_srcData(rowId)(startCol + 1))
    conf(TConfig.srcStart) = MFunc.Var2Long(m_srcData(rowId)(startCol + 2), 1)
    conf(TConfig.destName) = Trim$(m_srcData(rowId)(startCol + 3))
    conf(TConfig.destStart) = MFunc.Var2Long(m_srcData(rowId)(startCol + 4), 1)
    conf(TConfig.isOptional) = IIf(UCase(m_srcData(rowId)(startCol + 5)) = UCase(KW_Optional), True, False)
    conf(TConfig.beginRow) = rowId
    
    If m_configCount = 0 Then
        ReDim m_configs(m_configCount) As Variant
    Else
        ReDim Preserve m_configs(m_configCount) As Variant
    End If
    
    m_configs(m_configCount) = conf
    AppendConfig = m_configCount
    
    m_configCount = m_configCount + 1
End Function

Public Function PreviewData(SourceData, Optional ByRef errInfo As String) As Variant
On Error GoTo eh
    m_srcData = SourceData
    If Not MFunc.CheckArray(SourceData, m_srcDataProp) Then
        errInfo = "config data is not an array"
        Exit Function
    End If
    
    Dim iRow As Long
    Dim rowStart As Variant: rowStart = m_srcDataProp(ArrayProp.lb)
    Dim str As String
    Dim section As TConfig
    Dim UKW As UCaseKeyWords: UKW = MConfig.GetUCaseKeyWords()
    Dim curConfig As Long
    
    For iRow = rowStart To m_srcDataProp(ArrayProp.ub)
    
        str = UCase(Trim(SourceData(iRow)(rowStart)))
        
        Select Case str
            Case UKW.FieldsMap
                section = TConfig.accepts
                
            Case UKW.RelatedDataSheet
                section = TConfig.ConfigMain
                curConfig = AppendConfig(iRow)
            
            Case UKW.MSRDataSheet
                section = TConfig.MSRDataSheet
                curConfig = AppendMSRConfig(iRow)
                
            Case Else
                If section = TConfig.allTHeads Then
                    AppendCommonTHCHead iRow
                ElseIf section = TConfig.ConfigMain Then
                    m_configs(curConfig).endRow = iRow
                End If
        End Select
    Next
    
    PreviewData = Me.Result
    
Exit Function
eh:
    errInfo = Err.Description
    m_lastErr = errInfo
    Err.Clear
End Function

Private Function GetConfigId(ByVal srcName As String) As Long
    GetConfigId = -1
    Dim iConf As Long
    For iConf = 0 To m_configCount
        If StrComp(srcName, m_configs(iConf)(TConfig.srcName), vbTextCompare) = 0 Then
            GetConfigId = iConf
            Exit For
        End If
    Next
End Function

Private Function AppendNormalFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    Dim data(1) As Variant
    data(0) = Trim$(CStr(m_srcData(iRow)(m_srcDataProp(ArrayProp.lb))))
    data(1) = Trim$(CStr(m_srcData(iRow)(m_srcDataProp(ArrayProp.lb) + 1)))
    
    Dim arrLen As Long
    If IsArray(ret) Then
        arrLen = UBound(ret) + 1
        ReDim Preserve ret(arrLen) As Variant
    Else
        arrLen = 0
        Dim arr() As Variant: ReDim arr(0) As Variant
        ret = arr
    End If
    
    ret(arrLen) = data
    
    AppendNormalFields = arrLen
End Function

Private Function AppendConditionalField(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long

End Function

Private Function AppendDisplayFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long

End Function

Private Function AppendBalanceField(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long

End Function

Private Function AppendSortOutput(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long

End Function

Private Function AppendFakedFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long

End Function

Private Function AppendTHCHeadDescription(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long

End Function

Private Function AppendOBSItems(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long

End Function

Private Function AppendAcceptableValues(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long

End Function

Private Function AppendValueCheckFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long

End Function

Private Function AppendMSRDataSheet(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long

End Function

Public Function ReadConfig(ByVal srcName As String, Optional ByRef errInfo As String) As Variant
On Error GoTo eh
    Dim iConf As Long: iConf = GetConfigId(srcName)
    If iConf < 0 Then
        errInfo = "cannot find config"
        Exit Function
    End If
    
    Dim oDataConf As New CDataConfig
    
    
    Dim iRow As Long
    Dim section As Long
    Dim str As String
    Dim UKW As UCaseKeyWords: UKW = MConfig.GetUCaseKeyWords()
    Dim colStart As Long: colStart = m_srcDataProp(ArrayProp.lb)
    Dim isTitle As Boolean
    
    For iRow = m_configs(iConf)(TConfig.beginRow) + 1 To m_configs(iConf)(TConfig.endRow)
    
        str = UCase(Trim(m_srcData(iRow)(colStart)))
        isTitle = True
        
        Select Case str
            'Case UKW.FieldsMap: section = tconfig.allTHeads
            'Case UKW.RelatedDataSheet: section = tconfig.ConfigMain
            Case UKW.NormalFields: section = TConfig.normals
            Case UKW.ConditionalField: section = TConfig.conditions
            Case UKW.DisplayFields: section = TConfig.display
            Case UKW.BalanceField: section = TConfig.balanceHead
            Case UKW.SortOutput: section = TConfig.Sort
            Case UKW.FakedFields: section = TConfig.faked
            Case UKW.THCHeadDescription: section = TConfig.tHeads
            Case UKW.OBSItems: section = TConfig.obs
            Case UKW.AcceptableValues: section = TConfig.accepts
            Case UKW.ValueCheckFields: section = TConfig.Validation
            'Case UKW.MSRDataSheet: section = tconfig.MSRDataSheet
            Case Else
                isTitle = False
        End Select
        
        Select Case section
            'Case TConfig.allTHeads: AppendFieldsMap iRow, isTitle, m_configs(iConf)(section)
            'Case TConfig.ConfigMain: AppendRelatedDataSheet iRow, isTitle, m_configs(iConf)(section)
            Case TConfig.normals: AppendNormalFields iRow, isTitle, m_configs(iConf)(section)
            Case TConfig.conditions: AppendConditionalField iRow, isTitle, m_configs(iConf)(section)
            Case TConfig.display: AppendDisplayFields iRow, isTitle, m_configs(iConf)(section)
            Case TConfig.balanceHead: AppendBalanceField iRow, isTitle, m_configs(iConf)(section)
            Case TConfig.Sort: AppendSortOutput iRow, isTitle, m_configs(iConf)(section)
            Case TConfig.faked: AppendFakedFields iRow, isTitle, m_configs(iConf)(section)
            Case TConfig.tHeads: AppendTHCHeadDescription iRow, isTitle, m_configs(iConf)(section)
            Case TConfig.obs: AppendOBSItems iRow, isTitle, m_configs(iConf)(section)
            Case TConfig.accepts: AppendAcceptableValues iRow, isTitle, m_configs(iConf)(section)
            Case TConfig.Validation: AppendValueCheckFields iRow, isTitle, m_configs(iConf)(section)
            'Case TConfig.MSRDataSheet: AppendMSRDataSheet iRow, isTitle, m_configs(iConf)(section)
        End Select
    Next
    
Exit Function
eh:
    errInfo = Err.Description
    m_lastErr = errInfo
    Err.Clear
End Function

Public Property Get THCHeads() As Variant
    If m_headCount > 0 Then THCHeads = m_allHeads
End Property

Public Property Get Result() As Variant
    If m_configCount > 0 Then Result = m_configs
End Property

Public Property Get MSRs() As Variant
    If m_msrCount > 0 Then MSRs = m_msrs
End Property

