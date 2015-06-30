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

Private m_vBalanceField As Variant

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

Private Function IsBlankValues(SourceValues, iBegin As Long, iEnd As Long) As Boolean
    IsBlankValues = True
    Dim i As Long
    For i = iBegin To iEnd
        If SourceValues(i) <> vbNullString Then
            IsBlankValues = False
            Exit For
        End If
    Next
End Function

Private Function AppendNormalFields(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    If isTitle Then Exit Function
    
    Dim data(1) As Variant
    data(0) = Trim$(CStr(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol)))
    If data(0) = vbNullString Then Exit Function
    data(1) = Trim$(CStr(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 1)))
    
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendNormalFields = m_cntNormalFields
End Function

Private Function AppendConditionalField(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    Static body() As Variant
    Static headCount As Long
    Static bodyRowCount As Long
    
    Dim str As String
    Dim data() As Variant
    
    If isTitle Then
        str = Trim$(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 1))
        ReDim data(TCondition.BOF_ + 1 To TCondition.EOF_ - 1) As Variant
        If str = vbNullString Then
            data(TCondition.hasValidField) = False
        Else
            data(TCondition.hasValidField) = True
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
                        ReDim data(headCount) As Variant
                    Else
                        ReDim Preserve data(headCount) As Variant
                    End If
                    data(iCol) = str
                    headCount = headCount + 1
                End If
            Next
            
            If headCount > 0 Then
                m_vConditionalField(m_cntConditionalField - 1)(TCondition.heads) = data
            End If
        Else
            ReDim data(headCount - 1) As Variant
            
            For iCol = 0 To headCount - 1
                data(iCol) = Trim$(GetSafeArrayValue(m_srcData(iRow), iCol + m_srcBeginCol))
            Next
            
            If Not IsBlankValues(data, 0, headCount - 1) Then
                If bodyRowCount = 0 Then
                    ReDim body(bodyRowCount) As Variant
                Else
                    ReDim Preserve body(bodyRowCount) As Variant
                End If
                body(bodyRowCount) = data
                bodyRowCount = bodyRowCount + 1
                m_vConditionalField(m_cntConditionalField - 1)(TCondition.body) = body
            End If
        End If
    End If
    
    AppendConditionalField = m_cntConditionalField
End Function

Private Function AppendDisplayFields(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    If isTitle Then Exit Function
    
    Dim data(2) As Variant
    data(0) = Trim$(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol))
    If data(0) = vbNullString Then Exit Function
    
    data(1) = Trim$(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 1))
    data(2) = Trim$(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 2))
    
    If m_cntDisplayFields = 0 Then
        ReDim m_vDisplayFields(m_cntDisplayFields) As Variant
    Else
        ReDim Preserve m_vDisplayFields(m_cntDisplayFields) As Variant
    End If
    
    m_vDisplayFields(m_cntDisplayFields) = data
    m_cntDisplayFields = m_cntDisplayFields + 1
    AppendDisplayFields = m_cntDisplayFields
End Function

Private Function AppendBalanceField(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    If isTitle Then Exit Function
    Dim str As String: str = Trim$(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 1))
    If str <> vbNullString Then m_vBalanceField = str
End Function

Private Function AppendSortOutput(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    If isTitle Then Exit Function
    Static headCount As Long: headCount = 0
    
    Dim data() As Variant
    Dim iCol As Long
    Dim str As String
    
    If m_cntSortOutput = 0 Then
        For iCol = m_srcBeginCol To m_srcEndCol
            str = Trim$(GetSafeArrayValue(m_srcData(iRow), iCol))
            If str = vbNullString Then Exit For
            
            If headCount = 0 Then
                ReDim data(headCount) As Variant
            Else
                ReDim Preserve data(headCount) As Variant
            End If
            
            data(headCount) = str
            headCount = headCount + 1
        Next
        
        If headCount > 0 Then
            ReDim m_vSortOutput(m_cntSortOutput) As Variant
            m_vSortOutput(m_cntSortOutput) = data
            m_cntSortOutput = m_cntSortOutput + 1
        End If
    Else
        Dim nBlankValue As Long: nBlankValue = 0
        ReDim data(headCount - 1) As Variant
        
        For iCol = 0 To headCount - 1
            str = Trim$(GetSafeArrayValue(m_srcData(iRow), iCol + m_srcBeginCol))
            If str = vbNullString Then nBlankValue = nBlankValue + 1
            data(iCol) = str
        Next
        
        If nBlankValue < headCount Then
            ReDim Preserve m_vSortOutput(m_cntSortOutput) As Variant
            m_vSortOutput(m_cntSortOutput) = data
            m_cntSortOutput = m_cntSortOutput + 1
        End If
    End If
    
    AppendSortOutput = m_cntSortOutput
End Function

Private Function AppendFakedFields(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    If isTitle Then Exit Function
    
    Dim data(1) As Variant
    data(0) = Trim$(CStr(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol)))
    If data(0) = vbNullString Then Exit Function
    data(1) = Trim$(CStr(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 1)))
    
    If m_cntFakedFields = 0 Then
        ReDim m_vFakedFields(m_cntFakedFields) As Variant
    Else
        ReDim Preserve m_vFakedFields(m_cntFakedFields) As Variant
    End If
    
    m_vFakedFields(m_cntFakedFields) = data
    m_cntFakedFields = m_cntFakedFields + 1
    AppendFakedFields = m_cntFakedFields
End Function

Private Function AppendTHCHeadDescription(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    If isTitle Then Exit Function
    
    Dim data(1) As Variant
    data(0) = Trim$(CStr(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol)))
    If data(0) = vbNullString Then Exit Function
    data(1) = Trim$(CStr(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 1)))
    
    If m_cntTHCHeadDescription = 0 Then
        ReDim m_vTHCHeadDescription(m_cntTHCHeadDescription) As Variant
    Else
        ReDim Preserve m_vTHCHeadDescription(m_cntTHCHeadDescription) As Variant
    End If
    
    m_vTHCHeadDescription(m_cntTHCHeadDescription) = data
    m_cntTHCHeadDescription = m_cntTHCHeadDescription + 1
    AppendTHCHeadDescription = m_cntTHCHeadDescription
End Function

Private Function AppendOBSItems(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    Static body() As Variant
    Static headCount As Long
    Static bodyRowCount As Long
    
    Dim str As String
    Dim data() As Variant
    Dim iCol As Long
    Dim heads As Variant: heads = m_vOBSItems(m_cntOBSItems - 1)(TCondition.heads)
    
    If isTitle Then
        ReDim data(TCondition.BOF_ + 1 To TCondition.EOF_ - 1) As Variant
        If m_cntOBSItems = 0 Then
            ReDim m_vOBSItems(m_cntOBSItems) As Variant
        Else
            ReDim Preserve m_vOBSItems(m_cntOBSItems) As Variant
        End If
        m_vOBSItems(m_cntOBSItems) = data
        m_cntOBSItems = m_cntOBSItems + 1
        
        bodyRowCount = 0
        headCount = 0
    Else
        If IsEmpty(heads) Then
            For iCol = m_srcBeginCol To m_srcEndCol
                str = Trim$(GetSafeArrayValue(m_srcData(iRow), iCol))
                If str = vbNullString Then
                    Exit For
                Else
                    If headCount = 0 Then
                        ReDim data(headCount) As Variant
                    Else
                        ReDim Preserve data(headCount) As Variant
                    End If
                    data(iCol) = str
                    headCount = headCount + 1
                End If
            Next
            
            If headCount > 0 Then
                m_vOBSItems(m_cntOBSItems - 1)(TCondition.heads) = data
            End If
        Else
            Dim nBlankValue As Long
            ReDim data(headCount - 1) As Variant
            
            For iCol = 0 To headCount - 1
                data(iCol) = Trim$(GetSafeArrayValue(m_srcData(iRow), iCol + m_srcBeginCol))
                If data(iCol) = vbNullString Then nBlankValue = nBlankValue + 1
            Next
            
            If nBlankValue < headCount Then
                If bodyRowCount = 0 Then
                    ReDim body(bodyRowCount) As Variant
                Else
                    ReDim Preserve body(bodyRowCount) As Variant
                End If
                body(bodyRowCount) = data
                bodyRowCount = bodyRowCount + 1
                m_vOBSItems(m_cntOBSItems - 1)(TCondition.body) = body
            End If
        End If
    End If
    
    AppendOBSItems = m_cntOBSItems
End Function

Private Function AppendAcceptableValues(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    If isTitle Then Exit Function
    Dim data(1) As Variant
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendAcceptableValues = m_cntNormalFields
End Function

Private Function AppendValueCheckFields(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    If isTitle Then Exit Function
    Dim data(1) As Variant
    If m_cntNormalFields = 0 Then
        ReDim m_vNormalFields(m_cntNormalFields) As Variant
    Else
        ReDim Preserve m_vNormalFields(m_cntNormalFields) As Variant
    End If
    
    m_vNormalFields(m_cntNormalFields) = data
    m_cntNormalFields = m_cntNormalFields + 1
    AppendValueCheckFields = m_cntNormalFields
End Function

Private Function AppendMSRDataSheet(ByVal iRow As Long, ByVal isTitle As Boolean) As Long
    If isTitle Then
        If m_cntMSRDataSheet = 0 Then
            ReDim m_vMSRDataSheet(m_cntMSRDataSheet) As Variant
        Else
            ReDim Preserve m_vMSRDataSheet(m_cntMSRDataSheet) As Variant
        End If
        
        m_vMSRDataSheet(m_cntMSRDataSheet) = Trim$(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol))
        m_cntMSRDataSheet = m_cntMSRDataSheet + 1
        AppendMSRDataSheet = m_cntMSRDataSheet
    End If
End Function

Public Function AcceptData(SourceData) As Variant
    If Not ValidateSourceData(SourceData) Then Exit Function
    m_srcData = SourceData
    
    Dim iRow As Long
    Dim section As Long
    Dim str As String
    Dim UKW As UCaseKeyWords: UKW = MConfig.GetUCaseKeyWords()
    Dim colStart As Long: colStart = m_srcBeginCol
    Dim isTitle As Boolean
    
    For iRow = m_srcDataProp.lb To m_srcDataProp.ub
    
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
            'Case TConfig.allTHeads: AppendFieldsMap iRow, isTitle
            'Case TConfig.ConfigMain: AppendRelatedDataSheet iRow, isTitle
            Case TConfig.normals: AppendNormalFields iRow, isTitle
            Case TConfig.conditions: AppendConditionalField iRow, isTitle
            Case TConfig.display: AppendDisplayFields iRow, isTitle
            Case TConfig.balanceHead: AppendBalanceField iRow, isTitle
            Case TConfig.Sort: AppendSortOutput iRow, isTitle
            Case TConfig.faked: AppendFakedFields iRow, isTitle
            Case TConfig.tHeads: AppendTHCHeadDescription iRow, isTitle
            Case TConfig.obs: AppendOBSItems iRow, isTitle
            Case TConfig.accepts: AppendAcceptableValues iRow, isTitle
            Case TConfig.Validation: AppendValueCheckFields iRow, isTitle
            'Case TConfig.MSRDataSheet: AppendMSRDataSheet iRow, isTitle
        End Select
    Next
End Function

Public Sub Clear()
    m_cntNormalFields = 0
    m_cntConditionalField = 0
    m_cntDisplayFields = 0
    m_cntSortOutput = 0
    m_cntFakedFields = 0
    m_cntTHCHeadDescription = 0
    m_cntOBSItems = 0
    m_cntAcceptableValues = 0
    m_cntValueCheckFields = 0
End Sub
