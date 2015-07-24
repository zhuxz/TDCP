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

Private m_srcDataProp As ArrayProp_

Private m_vSections() As Variant
Private m_cntSection As Long
Private m_vBalanceField As String

Private m_srcBalanceField As String
Public m_srcName As String
Public m_srcStart As Long
Public m_destName As String
Public m_destStart As Long
Public m_isOptional As Boolean
Public m_srcBeginRow As Long
Public m_srcEndRow As Long
Public m_srcBeginCol As Long
Public m_srcEndCol As Long

Private Sub Class_Initialize()
    Me.Clear
End Sub

Public Sub Clear()
    m_cntSection = 0
    m_vBalanceField = vbNullString
    m_srcBeginCol = 0
    m_srcEndCol = 0
End Sub

Private Function ValidateSourceData(SourceData) As Boolean
On Error GoTo eh
    Dim prop As Variant
    Dim ret As Boolean: ret = CheckArray(SourceData, prop)
    If (ret) Then
        m_srcDataProp.lb = CLng(prop(ArrayProp.lb))
        m_srcDataProp.ub = CLng(prop(ArrayProp.ub))
        m_srcDataProp.isArr = CBool(prop(ArrayProp.isArr))
        m_srcDataProp.Size = CBool(prop(ArrayProp.Size))
        m_srcBeginCol = LBound(SourceData(m_srcDataProp.lb))
        m_srcEndCol = UBound(SourceData(m_srcDataProp.ub))
    Else
        ValidateSourceData = False
    End If
    Exit Function
eh:
    ValidateSourceData = False
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

Private Function AppendSection(Optional ByVal sectionType As Long = -1) As Long
    Dim data(TCSection.BOF_ + 1 To TCSection.EOF_ - 1) As Variant
    data(TCSection.type_) = sectionType
    If m_cntSection = 0 Then
        ReDim m_vSections(m_cntSection) As Variant
    Else
        ReDim Preserve m_vSections(m_cntSection) As Variant
    End If
    m_vSections(m_cntSection) = data
    m_cntSection = m_cntSection + 1
    AppendSection = m_cntSection
End Function

Public Function AppendNormalFields(dataRow As Variant, isTitle As Boolean, Optional beginCol As Long = 1) As Long
    Static sectionId As Long
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
        
        RedimVarArr rows, nRow
        rows(nRow) = data
        m_vSections(sectionId - 1)(TCSection.detail) = rows
        
        nRow = nRow + 1
    End If
End Function

Public Function AppendConditionalField(dataRow As Variant, isTitle As Boolean, Optional beginCol As Long = 1, Optional endCol As Long = 1) As Long
    Static sectionId As Long
    Static rows() As Variant
    Static nRow As Long
    Static nHead As Long

    Dim str As String
    Dim data() As Variant
    Dim iCol As Long
    
    If isTitle Then
        nRow = 0: nHead = 0
        
        sectionId = AppendSection(TCSType.conditions)
        data = New_TCondition()
        str = Trim$(GetSafeArrayValue(dataRow, beginCol + 1))
        If str = vbNullString Then
            data(TCondition.hasValidField) = False
        Else
            data(TCondition.hasValidField) = True
            data(TCondition.ValidField) = str
        End If
        m_vSections(sectionId - 1)(TCSection.detail) = data

        AppendConditionalField = sectionId
    Else
        If nHead = 0 Then
            For iCol = beginCol To endCol
                str = Trim$(GetSafeArrayValue(dataRow, iCol))
                If str = vbNullString Then
                    Exit For
                Else
                    RedimVarArr data, nHead
                    data(nHead) = str
                    nHead = nHead + 1
                End If
            Next

            If nHead > 0 Then
                m_vSections(sectionId - 1)(TCSection.detail)(TCondition.heads) = data
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
                RedimVarArr rows, nRow
                rows(nRow) = data
                nRow = nRow + 1
                m_vSections(sectionId - 1)(TCSection.detail)(TCondition.body) = rows
            End If
        End If
    End If
End Function

Public Function AppendDisplayFields(dataRow, isTitle As Boolean, Optional beginCol As Long = 1) As Long
    Static sectionId As Long
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
        
        RedimVarArr rows, nRow
        rows(nRow) = data
        nRow = nRow + 1
        
        m_vSections(sectionId - 1)(TCSection.detail) = rows
    End If
End Function

Public Function AppendBalanceField(dataRow, ByVal isTitle As Boolean, Optional beginCol As Long = 1) As Long
    Static sectionId As Long
    If isTitle Then
        sectionId = AppendSection(TCSType.balanceHead)
        AppendBalanceField = sectionId
    Else
        m_vSections(sectionId - 1)(TCSection.detail) = Trim$(GetSafeArrayValue(dataRow, beginCol))
    End If
End Function

Public Function AppendSortOutput(dataRow, ByVal isTitle As Boolean, Optional beginCol As Long = -1, Optional endCol As Long = -1) As Long
    Static sectionId As Long
    Static rows() As Variant
    Static nRow As Long
    Static nHead As Long
    
    Dim data() As Variant
    Dim iCol As Long
    Dim str As String
    
    If isTitle Then
        nRow = 0: nHead = 0
        sectionId = AppendSection(TCSType.Sort)
        m_vSections(sectionId - 1)(TCSection.detail) = New_TSort()
        AppendSortOutput = sectionId
    Else
        If nHead = 0 Then
            For iCol = beginCol To endCol
                str = Trim$(GetSafeArrayValue(dataRow, iCol))
                If str = vbNullString Then Exit For
                RedimVarArr data, nHead
                data(nHead) = str
                nHead = nHead + 1
            Next
            
            If nHead > 0 Then
                m_vSections(sectionId - 1)(TCSection.detail)(TSort.heads) = data
            End If
        Else
            Dim nBlankValue As Long: nBlankValue = 0
            ReDim data(nHead - 1) As Variant
            
            For iCol = 0 To nHead - 1
                str = Trim$(GetSafeArrayValue(dataRow, iCol + m_srcBeginCol))
                If str = vbNullString Then nBlankValue = nBlankValue + 1
                data(iCol) = str
            Next
            
            If nBlankValue < nHead Then
                RedimVarArr rows, nRow
                rows(nRow) = data
                nRow = nRow + 1
                m_vSections(sectionId - 1)(TCSection.detail)(TCondition.body) = rows
            End If
        End If
    End If
End Function

Public Function AppendFakedFields(dataRow, ByVal isTitle As Boolean, Optional beginCol As Long = -1) As Long
    Static sectionId As Long
    Static rows() As Variant
    Static nRow As Long
    
    If isTitle Then
        nRow = 0
        sectionId = AppendSection(TCSType.faked)
        AppendFakedFields = sectionId
    Else
        Dim data(1) As Variant
        data(0) = Trim$(GetSafeArrayValue(dataRow, beginCol))
        If data(0) = vbNullString Then Exit Function
        data(1) = Trim$(GetSafeArrayValue(dataRow, beginCol + 1))
        
        RedimVarArr rows, nRow
        rows(nRow) = data
        m_vSections(sectionId - 1)(TCSection.detail) = rows
        
        nRow = nRow + 1
    End If
End Function

Public Function AppendTHCHeadDescription(dataRow, ByVal isTitle As Boolean, Optional beginCol As Long = -1) As Long
    Static sectionId As Long
    Static rows() As Variant
    Static nRow As Long
    
    If isTitle Then
        nRow = 0
        sectionId = AppendSection(TCSType.tHeads)
        AppendTHCHeadDescription = sectionId
    Else
        Dim data(1) As Variant
        data(0) = Trim$(GetSafeArrayValue(dataRow, beginCol))
        If data(0) = vbNullString Then Exit Function
        data(1) = Trim$(GetSafeArrayValue(dataRow, beginCol + 1))
        
        RedimVarArr rows, nRow
        rows(nRow) = data
        m_vSections(sectionId - 1)(TCSection.detail) = rows
        
        nRow = nRow + 1
    End If
End Function

Public Function AppendOBSItems(dataRow, ByVal isTitle As Boolean, Optional beginCol As Long = 1, Optional endCol As Long = 1) As Long
    Static sectionId As Long
    Static rows() As Variant
    Static nRow As Long
    Static nHead As Long

    Dim str As String
    Dim data() As Variant
    Dim iCol As Long
    
    If isTitle Then
        nRow = 0: nHead = 0
        
        sectionId = AppendSection(TCSType.obs)
        m_vSections(sectionId - 1)(TCSection.detail) = New_TCondition()

        AppendOBSItems = sectionId
    Else
        If nHead = 0 Then
            For iCol = beginCol To endCol
                str = Trim$(GetSafeArrayValue(dataRow, iCol))
                If str = vbNullString Then
                    Exit For
                Else
                    RedimVarArr data, nHead
                    data(nHead) = str
                    nHead = nHead + 1
                End If
            Next

            If nHead > 0 Then
                m_vSections(sectionId - 1)(TCSection.detail)(TCondition.heads) = data
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
                RedimVarArr rows, nRow
                rows(nRow) = data
                nRow = nRow + 1
                m_vSections(sectionId - 1)(TCSection.detail)(TCondition.body) = rows
            End If
        End If
    End If
End Function

Public Function AppendAcceptableValues(dataRow, ByVal isTitle As Boolean, Optional beginCol As Long = 1, Optional endCol As Long = 1) As Long
    Static sectionId As Long
    Static rows() As Variant
    Static nRow As Long

    Dim str As String
    Dim data() As Variant
    
    If isTitle Then
        nRow = 0
        
        sectionId = AppendSection(TCSType.accepts)
        data = New_TCondition()
        str = Trim$(GetSafeArrayValue(dataRow, beginCol + 1))
        If str = vbNullString Then
            data(TAccept.srcFieldId) = -1
        Else
            data(TAccept.srcField) = str
        End If
        m_vSections(sectionId - 1)(TCSection.detail) = data

        AppendAcceptableValues = sectionId
    Else
        ReDim data(1) As Variant
        data(0) = Trim$(GetSafeArrayValue(dataRow, beginCol))
        If data(0) = vbNullString Then Exit Function
        data(1) = Trim$(GetSafeArrayValue(dataRow, beginCol + 1))
        
        RedimVarArr rows, nRow
        rows(nRow) = data
        m_vSections(sectionId - 1)(TCSection.detail)(TAccept.body) = rows
        
        nRow = nRow + 1
    End If
End Function

Public Function AppendValueCheckFields(dataRow, ByVal isTitle As Boolean, Optional beginCol As Long = 1, Optional endCol As Long = 1) As Long
    Static sectionId As Long
    Static rows() As Variant
    Static nRow As Long
    
    If isTitle Then
        nRow = 0
        sectionId = AppendSection(TCSType.Validation)
        AppendValueCheckFields = sectionId
    Else
        Dim data As Variant: data = New_TValueCheck()
        Dim str As String
        
        data(TValueCheck.srcField) = Trim$(GetSafeArrayValue(dataRow, beginCol))
        If data(TValueCheck.srcField) = vbNullString Then Exit Function
        data(TValueCheck.Desc) = Trim$(GetSafeArrayValue(dataRow, beginCol + 1))
        str = Trim$(GetSafeArrayValue(dataRow, beginCol + 2))
        If str = ">" Then
            data(TValueCheck.rule) = VCRule.MoreThan
            data(TValueCheck.data) = Trim$(GetSafeArrayValue(dataRow, beginCol + 3))
        ElseIf str = "<" Then
            data(TValueCheck.rule) = VCRule.LessThan
            data(TValueCheck.data) = Trim$(GetSafeArrayValue(dataRow, beginCol + 3))
        ElseIf str = "=" Then
            data(TValueCheck.rule) = VCRule.equal
            data(TValueCheck.data) = Trim$(GetSafeArrayValue(dataRow, beginCol + 3))
        ElseIf StrComp(str, "Between", vbTextCompare) = 0 Then
            data(TValueCheck.rule) = VCRule.Between
            
            Dim iCol As Long
            Dim oUpper As New CArray, oLower As New CArray
            Dim strUpper As String, strLower As String
            
            For iCol = beginCol + 4 To endCol Step 2
                strLower = Trim$(GetSafeArrayValue(dataRow, iCol))
                strUpper = Trim$(GetSafeArrayValue(dataRow, iCol + 1))
                If strLower = vbNullString Then Exit For
                If strUpper = vbNullString Then Exit For
                oLower.AppendVarItem IIf(IsNumeric(strLower), Val(strLower), strLower)
                oUpper.AppendVarItem IIf(IsNumeric(strUpper), Val(strUpper), strUpper)
            Next
            
            If (oLower.Count > 0) Then
                data(TValueCheck.lowers) = oLower.VarList
                data(TValueCheck.uppers) = oUpper.VarList
            End If
            
            Set oLower = Nothing
            Set oUpper = Nothing
        ElseIf StrComp(str, "NULL", vbTextCompare) Then
            data(TValueCheck.rule) = VCRule.NULL_
        ElseIf StrComp(str, "NOT NULL", vbTextCompare) Then
            data(TValueCheck.rule) = VCRule.NULL_
        End If
        
        RedimVarArr rows, nRow
        rows(nRow) = data
        m_vSections(sectionId - 1)(TCSection.detail) = rows
        
        nRow = nRow + 1
    End If
End Function

Public Function AcceptData(SourceData) As Variant
    If Not ValidateSourceData(SourceData) Then Exit Function
    
    Dim iRow As Long
    Dim sectionType As Long
    Dim str As String
    Dim UKW As UCaseKeyWords: UKW = MConfig.GetUCaseKeyWords()
    Dim colStart As Long: colStart = m_srcBeginCol
    Dim isTitle As Boolean
    Dim curSection As Long
    Dim row As Variant
    
    For iRow = m_srcDataProp.lb To m_srcDataProp.ub
        row = SourceData(iRow)
        str = UCase(Trim(row(colStart)))
        isTitle = True
        
        Select Case str
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
            Case Else
                isTitle = False
        End Select
        
        Select Case sectionType
            Case TCSType.normals: AppendNormalFields row, isTitle
            Case TCSType.conditions: AppendConditionalField row, isTitle, colStart, m_srcEndCol
            Case TCSType.display: AppendDisplayFields row, isTitle
            Case TCSType.balanceHead: AppendBalanceField row, isTitle
            Case TCSType.Sort: AppendSortOutput row, isTitle
            Case TCSType.faked: AppendFakedFields row, isTitle
            Case TCSType.tHeads: AppendTHCHeadDescription row, isTitle
            Case TCSType.obs: AppendOBSItems row, isTitle, colStart, m_srcEndCol
            Case TCSType.accepts: AppendAcceptableValues row, isTitle
            Case TCSType.Validation: AppendValueCheckFields row, isTitle
        End Select
    Next
End Function
