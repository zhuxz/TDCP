VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_srcData As Variant
Private m_srcDataProp As ArrayProp_
Private m_srcBeginCol As Long
Private m_srcEndCol As Long

Private m_vNormalFields() As Variant
Private m_cntNormalFields As Long

Private m_vConditionalField() As Variant
Private m_cntConditionalField As Long

Private m_vDisplayFields() As Variant
Private m_cntDisplayFields As Long

Private m_vBalanceField() As Variant
Private m_cntBalanceField As Long

Private m_vSortOutput() As Variant
Private m_cntSortOutput As Long

Private m_vFakedFields() As Variant
Private m_cntFakedFields As Long

Private m_vTHCHeadDescription() As Variant
Private m_cntTHCHeadDescription As Long

Private m_vOBSItems() As Variant
Private m_cntOBSItems As Long

Private m_vAcceptableValues() As Variant
Private m_cntAcceptableValues As Long

Private m_vValueCheckFields() As Variant
Private m_cntValueCheckFields As Long

Private m_vMSRDataSheet() As Variant
Private m_cntMSRDataSheet As Long

Private Sub Class_Initialize()
    Me.Clear
End Sub

Private Function ValidateSourceData(SourceData) As Boolean
    Dim prop As Variant
    Dim ret As Boolean: ret = CheckArray(SourceData, prop)
    m_srcDataProp.lb = CLng(prop(ArrayProp.lb))
    m_srcDataProp.ub = CLng(prop(ArrayProp.ub))
    m_srcDataProp.isArr = CBool(prop(ArrayProp.isArr))
    m_srcDataProp.Size = CBool(prop(ArrayProp.Size))
    m_srcBeginCol = LBound(SourceData(m_srcDataProp.lb))
    m_srcEndCol = LBound(SourceData(m_srcDataProp.ub))
    ValidateSourceData = ret
End Function

Private Function AppendNormalFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    Dim data(1) As Variant
    data(0) = Trim$(CStr(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol)))
    data(1) = Trim$(CStr(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 1)))
    If data(0) = vbNullString Then Exit Function
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Private Function AppendConditionalField(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    Static body() As Variant
    Static headCount As Long
    Static bodyRowCount As Long
    Dim str As String
    Dim data() As Variant
    
    If isTitle Then
        Dim str As String: str = Trim$(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 1))
        ReDim data(TCondition.BOF_ + 1 To TCondition.EOF_ - 1) As Variant
        If str = vbNullString Then
            data(TCondition.hasValid) = False
        Else
            data(TCondition.hasValid) = True
            data(TCondition.ValidField) = str
        End If
        
        If m_cntConditionalField = 0 Then
            ReDim m_vConditionalField(m_cntConditionalField) As Variant
        Else
            ReDim Preserve m_vConditionalField(m_cntConditionalField) As Variant
        End If
        
        m_vConditionalField(m_cntConditionalField) = data
        m_cntConditionalField = m_cntConditionalField + 1
        
        bodyRowCount = 0
        headCount = 0
    Else
        Dim iCol As Long
        Dim heads As Variant: heads = m_vConditionalField(m_cntConditionalField - 1)(TCondition.heads)
        
        If IsEmpty(heads) Then
            headCount = 0
            For iCol = m_srcBeginCol To m_srcEndCol
                str = Trim$(GetSafeArrayValue(m_srcData(iRow), iCol))
                If str = vbNullString Then
                    Exit For
                Else
                    If headCount = 0 Then
                        ReDim data(headCount) As String
                    Else
                        ReDim Preserve data(headCount) As String
                    End If
                    data(iCol) = str
                    headCount = headCount + 1
                End If
            Next
            
            If headCount > 0 Then
                m_vConditionalField(m_cntConditionalField)(TCondition.heads) = data
            End If
        Else
            ReDim data(headCount - 1) As String
            For iCol = 0 To headCount - 1
                data(iCol) = GetSafeArrayValue(m_srcData(iRow), iCol + m_srcBeginCol)
            Next
            
            If bodyRowCount = 0 Then
                ReDim body(bodyRowCount) As Variant
            Else
                ReDim Preserve body(bodyRowCount) As Variant
            End If
            body(bodyRowCount) = data
        End If
    End If
    
    AppendConditionalField = m_cntConditionalField
End Function

Private Function AppendDisplayFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Private Function AppendBalanceField(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Private Function AppendSortOutput(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Private Function AppendFakedFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Private Function AppendTHCHeadDescription(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Private Function AppendOBSItems(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Private Function AppendAcceptableValues(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Private Function AppendValueCheckFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Private Function AppendMSRDataSheet(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
    If isTitle Then Exit Function
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Public Function AcceptData(SourceData) As Variant
    If Not ValidateSourceData(SourceData) Then Exit Function
    m_srcData = SourceData
    
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
End Function

Public Sub Clear()
    m_cntNormalFields = 0
    m_cntConditionalField = 0
    m_cntDisplayFields = 0
    m_cntBalanceField = 0
    m_cntSortOutput = 0
    m_cntFakedFields = 0
    m_cntTHCHeadDescription = 0
    m_cntOBSItems = 0
    m_cntAcceptableValues = 0
    m_cntValueCheckFields = 0
End Sub
