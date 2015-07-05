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

Private m_vSections() As Variant
Private m_cntSection As Long

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

Private Function AppendSection(Optional ByVal whichSection As Long = -1) As Long
    Dim data(TCSection.BOF_ + 1 To TCSection.EOF_ - 1) As Variant
    data(TCSection.type_) = whichSection
    If m_cntSection = 0 Then
        ReDim m_vSections(m_cntSection) As Variant
    Else
        ReDim Preserve m_vSections(m_cntSection) As Variant
    End If
    m_vSections(m_cntSection) = data
    m_cntSection = m_cntSection + 1
    AppendSection = m_cntSection
End Function

Private Function AppendNormalFields(dataRow As Variant, isTitle As Boolean, sectionId As Long, Optional beginCol As Long = 1) As Long
    Static rows() As Variant
    Static nRow As Long
    
    If isTitle Then
        nRow = 0
        sectionId = AppendSection(TCSType.normals)
        AppendNormalFields = sectionId
    Else
        Dim data(1) As Variant
        data(0) = Trim$(GetSafeArrayValue(dataRow, beginCol))
        If data(0) = vbNullString Then Exit Function
        data(1) = Trim$(GetSafeArrayValue(dataRow, beginCol + 1))
        
        If nRow = 0 Then
            ReDim rows(nRow) As Variant
        Else
            ReDim Preserve rows(nRow) As Variant
        End If
        rows(nRow) = data
        nRow = nRow + 1
        
        m_vSections(sectionId - 1)(TCSection.detail) = rows
    End If
End Function

Private Function AppendConditionalField(dataRow As Variant, isTitle As Boolean, sectionId As Long, Optional beginCol As Long = 1, Optional endCol As Long = 1) As Long
    Static rows() As Variant
    Static nRow As Long
    Static nHead As Long

    Dim str As String
    Dim data() As Variant

    If isTitle Then
        nRow = 0: nHead = 0
        
        sectionId = AppendSection(TCSType.conditions)
        str = Trim$(GetSafeArrayValue(dataRow, beginCol + 1))
        ReDim data(TCondition.BOF_ + 1 To TCondition.EOF_ - 1) As Variant
        If str = vbNullString Then
            data(TCondition.hasValidField) = False
        Else
            data(TCondition.hasValidField) = True
            data(TCondition.ValidField) = str
        End If
        m_vConditionalField(sectionId - 1)(TCSection.detail) = data

        AppendConditionalField = sectionId
    Else
        Dim iCol As Long

        If nHead = 0 Then
            For iCol = beginCol To endCol
                str = Trim$(GetSafeArrayValue(dataRow, iCol))
                If str = vbNullString Then
                    Exit For
                Else
                    If nHead = 0 Then
                        ReDim data(nHead) As Variant
                    Else
                        ReDim Preserve data(nHead) As Variant
                    End If
                    data(nHead) = str
                    nHead = nHead + 1
                End If
            Next

            If nHead > 0 Then
                m_vConditionalField(sectionId - 1)(TCSection.detail)(TCondition.heads) = data
            End If
        Else
            Dim nBlankValue As Long: nBlankValue = 0
            ReDim data(nHead - 1) As Variant

            For iCol = 0 To nHead - 1
                str = Trim$(GetSafeArrayValue(dataRow, iCol + beginCol))
                data(iCol) = str
                If str = vbNullString Then nBlankValue = nBlankValue + 1
            Next

            If nBlankValue < nHead Then
                If nRow = 0 Then
                    ReDim rows(nRow) As Variant
                Else
                    ReDim Preserve rows(nRow) As Variant
                End If
                rows(nRow) = data
                nRow = nRow + 1
                m_vConditionalField(sectionId - 1)(TCSection.detail)(TCondition.body) = rows
            End If
        End If
    End If
End Function

Private Function AppendDisplayFields(dataRow As Variant, isTitle As Boolean, sectionId As Long, Optional beginCol As Long = 1) As Long
    Static rows() As Variant
    Static nRow As Long
    
    If isTitle Then
        nRow = 0
        sectionId = AppendSection(TCSType.display)
        AppendDisplayFields = sectionId
    Else
        Dim data(2) As Variant
        data(0) = Trim$(GetSafeArrayValue(dataRow, beginCol))
        If data(0) = vbNullString Then Exit Function
        data(1) = Trim$(GetSafeArrayValue(dataRow, beginCol + 1))
        data(2) = Trim$(GetSafeArrayValue(dataRow, beginCol + 2))
        
        If nRow = 0 Then
            ReDim rows(nRow) As Variant
        Else
            ReDim Preserve rows(nRow) As Variant
        End If
        rows(nRow) = data
        nRow = nRow + 1
        
        m_vSections(sectionId - 1)(TCSection.detail) = rows
    End If
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
    Static body() As Variant
    Static bodyRowCount As Long
    
    Dim data() As Variant
    
    If isTitle Then
        ReDim data(TAccept.BOF_ + 1 To TAccept.EOF_ - 1) As Variant
        data(TAccept.srcField) = Trim$(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 1))
        If m_cntAcceptableValues = 0 Then
            ReDim m_vAcceptableValues(m_cntAcceptableValues) As Variant
        Else
            ReDim Preserve m_vAcceptableValues(m_cntAcceptableValues) As Variant
        End If
        
        m_vAcceptableValues(m_cntAcceptableValues) = data
        m_cntAcceptableValues = m_cntAcceptableValues + 1
        
        bodyRowCount = 0
    Else
        ReDim data(1) As Variant
        data(0) = Trim$(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol))
        data(1) = Trim$(GetSafeArrayValue(m_srcData(iRow), m_srcBeginCol + 1))
        If (data(0) <> vbNullString And data(1) <> vbNullString) Then
            If bodyRowCount = 0 Then
                ReDim body(bodyRowCount) As Variant
            Else
                ReDim Preserve body(bodyRowCount) As Variant
            End If
            body(bodyRowCount) = data
            m_vAcceptableValues(m_cntAcceptableValues - 1)(TAccept.body) = body
        End If
    End If
    
    AppendAcceptableValues = m_cntAcceptableValues
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
    Dim sectionType As Long
    Dim str As String
    Dim UKW As UCaseKeyWords: UKW = MConfig.GetUCaseKeyWords()
    Dim colStart As Long: colStart = m_srcBeginCol
    Dim isTitle As Boolean
    Dim curSection As Long
    Dim row As Variant
    
    For iRow = m_srcDataProp.lb To m_srcDataProp.ub
        row = m_srcData(iRow)
        str = UCase(Trim(row(colStart)))
        isTitle = True
        
        Select Case str
            'Case UKW.FieldsMap: sectionType = tconfig.allTHeads
            'Case UKW.RelatedDataSheet: sectionType = tconfig.ConfigMain
            Case UKW.NormalFields: sectionType = TCSType.normals
            Case UKW.ConditionalField: sectionType = TCSType.conditions
            Case UKW.DisplayFields: sectionType = TCSType.display
            Case UKW.BalanceField: sectionType = TCSType.balanceHead
            Case UKW.SortOutput: sectionType = TCSType.Sort
            Case UKW.FakedFields: sectionType = TCSType.faked
            Case UKW.THCHeadDescription: sectionType = TCSType.tHeads
            Case UKW.OBSItems: sectionType = TCSType.obs
            Case UKW.AcceptableValues: sectionType = TCSType.accepts
            Case UKW.ValueCheckFields: sectionType = TCSType.Validation
            'Case UKW.MSRDataSheet: sectionType = tconfig.MSRDataSheet
            Case Else
                isTitle = False
        End Select
        
        Select Case sectionType
            'Case TCSType.allTHeads: AppendFieldsMap row, isTitle
            'Case TCSType.ConfigMain: AppendRelatedDataSheet row, isTitle
            Case TCSType.normals: AppendNormalFields row, isTitle, curSection, colStart
            Case TCSType.conditions: AppendConditionalField row, isTitle, curSection, colStart, m_srcEndCol
            Case TCSType.display: AppendDisplayFields row, isTitle
            Case TCSType.balanceHead: AppendBalanceField row, isTitle
            Case TCSType.Sort: AppendSortOutput row, isTitle
            Case TCSType.faked: AppendFakedFields row, isTitle
            Case TCSType.tHeads: AppendTHCHeadDescription row, isTitle
            Case TCSType.obs: AppendOBSItems row, isTitle
            Case TCSType.accepts: AppendAcceptableValues row, isTitle
            Case TCSType.Validation: AppendValueCheckFields row, isTitle
            'Case TCSType.MSRDataSheet: AppendMSRDataSheet row, isTitle
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
