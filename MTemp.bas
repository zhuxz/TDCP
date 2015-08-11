Attribute VB_Name = "MTemp"
'Private Function AppendNormalFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'    If isTitle Then Exit Function
'
'    Dim data(1) As Variant
'    data(0) = Trim$(CStr(m_srcData(iRow)(m_srcDataProp(ArrayProp.lb))))
'    data(1) = Trim$(CStr(m_srcData(iRow)(m_srcDataProp(ArrayProp.lb) + 1)))
'
'    Dim arrLen As Long
'    If IsArray(ret) Then
'        arrLen = UBound(ret) + 1
'        ReDim Preserve ret(arrLen) As Variant
'    Else
'        arrLen = 0
'        Dim arr() As Variant: ReDim arr(0) As Variant
'        ret = arr
'    End If
'
'    ret(arrLen) = data
'
'    AppendNormalFields = arrLen
'End Function
'
'Private Function AppendConditionalField(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'
'End Function
'
'Private Function AppendDisplayFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'
'End Function
'
'Private Function AppendBalanceField(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'
'End Function
'
'Private Function AppendSortOutput(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'
'End Function
'
'Private Function AppendFakedFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'
'End Function
'
'Private Function AppendTHCHeadDescription(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'
'End Function
'
'Private Function AppendOBSItems(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'
'End Function
'
'Private Function AppendAcceptableValues(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'
'End Function
'
'Private Function AppendValueCheckFields(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'
'End Function
'
'Private Function AppendMSRDataSheet(ByVal iRow As Long, ByVal isTitle As Long, ret) As Long
'
'End Function

Private Function getSection(sectionId As Long) As Variant
    Dim ret As Variant
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

Private Function GetXLSApp() As Excel.Application
    If m_xlsApp Is Nothing Then
        Set m_xlsApp = New Excel.Application
        m_xlsApp.Visible = m_isDebug
    End If
    Set GetXLSApp = m_xlsApp
End Function
