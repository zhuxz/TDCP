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

Private m_srcDataProp As ArrayProp

Private m_vSections() As Variant
Private m_nSection As Long

Private m_vOutputs()
Private m_nOutput As Long

Private m_vDataHeads()
Private m_nDataHead As Long

Private m_vMissFields()
Private m_nMissField As Long

Private m_vParseErrs()
Private m_nParseErr As Long

Private m_vExecErrs()
Private m_nExecErr As Long

'Private m_vBalanceField As String
'Private m_srcBalanceField As String

Public m_srcName As String
Public m_srcStart As Long
Public m_destName As String
Public m_destStart As Long
Public m_isOptional As Boolean
Public m_srcBeginRow As Long
Public m_srcEndRow As Long
Public m_srcBeginCol As Long
Public m_srcEndCol As Long

Public m_reportDate As Date

Public m_exp As CExpression

Private Sub Class_Initialize()
    Me.Clear
End Sub

Private Sub Class_Terminate()
    Set m_exp = Nothing
End Sub

Public Sub Clear()
    m_nSection = 0
    'm_vBalanceField = vbNullString
    m_srcBeginCol = 0
    m_srcEndCol = 0
    Set m_exp = New CExpression
End Sub

Private Function ValidateSourceData(SourceData) As Boolean
On Error GoTo eh
    Dim ret As Boolean: ret = CheckArray(SourceData, m_srcDataProp)
    If (ret) Then
        m_srcBeginCol = LBound(SourceData(m_srcDataProp.lb))
        m_srcEndCol = UBound(SourceData(m_srcDataProp.ub))
        ValidateSourceData = True
    Else
        ValidateSourceData = False
    End If
    Exit Function
eh:
    ValidateSourceData = False
End Function

Private Function cacheParseErr(expression)
    ReDim Preserve m_vParseErrs(m_nParseErr)
    m_vParseErrs(m_nParseErr) = expression
    m_nParseErr = m_nParseErr + 1
End Function

Public Function GetParseErrXML() As String
    If m_nParseErr < 1 Then Exit Function
    Dim retVal() As String
    VarArrAppend retVal, "<ParseErr>"
    Dim i As Long
    For i = 0 To m_nParseErr - 1
        VarArrAppend retVal, "<i>" & CXml(m_vParseErrs(i)) & "</i>"
    Next
    VarArrAppend retVal, "</ParseErr>"
    GetParseErrXML = Join(retVal, "")
End Function

Private Function cacheMissField(fieldName)
    ReDim Preserve m_vMissFields(m_nMissField)
    m_vMissFields(m_nMissField) = fieldName
    m_nMissField = m_nMissField + 1
End Function

Public Function GetMissFieldXML() As String
    If m_nMissField < 1 Then Exit Function
    Dim retVal()
    VarArrAppend retVal, "<MissField>"
    Dim i As Long
    For i = 0 To m_nMissField - 1
        VarArrAppend retVal, "<i>" & CXml(m_vMissFields(i)) & "</i>"
    Next
    VarArrAppend retVal, "</MissField>"
    GetMissFieldXML = Join(retVal, "")
End Function

Private Function cacheExecErr(rowId, expression, errDesc)
    ReDim Preserve m_vExecErrs(m_nExecErr)
    Dim e: e = New_ExecErr()
    e(ExecErr.rowId) = rowId
    e(ExecErr.expression) = expression
    e(ExecErr.errDesc) = errDesc
    m_vExecErrs(m_nExecErr) = e
    m_nExecErr = m_nExecErr + 1
End Function

Public Function GetExecErrXML() As String
    If m_nExecErr < 1 Then Exit Function
    Dim retVal()
    VarArrAppend retVal, "<ExecErr>"
    Dim i As Long
    For i = 0 To m_nExecErr - 1
        VarArrAppend retVal, "<i r=""" & m_vExecErrs(i)(ExecErr.rowId) & """"
        VarArrAppend retVal, " exp=""" & CXml(m_vExecErrs(i)(ExecErr.expression)) & """"
        VarArrAppend retVal, " desc=""" & CXml(m_vExecErrs(i)(ExecErr.errDesc)) & """"
        VarArrAppend retVal, " />"
    Next
    VarArrAppend retVal, "</ExecErr>"
    GetExecErrXML = Join(retVal, "")
End Function

Private Function AppendSection(Optional ByVal sectionType As Long = -1) As Long
    Dim data(TCSection.BOF_ + 1 To TCSection.EOF_ - 1) As Variant
    data(TCSection.type_) = sectionType
    If m_nSection = 0 Then
        ReDim m_vSections(m_nSection) As Variant
    Else
        ReDim Preserve m_vSections(m_nSection) As Variant
    End If
    m_vSections(m_nSection) = data
    m_nSection = m_nSection + 1
    AppendSection = m_nSection
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

Public Function AppendSortOutput(dataRow, ByVal isTitle As Boolean, Optional beginCol As Long = 1, Optional endCol As Long = -1) As Long
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

Public Function AppendFakedFields(dataRow, ByVal isTitle As Boolean, Optional beginCol As Long = 1) As Long
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

Public Function AppendTHCHeadDescription(dataRow, ByVal isTitle As Boolean, Optional beginCol As Long = 1) As Long
    Static sectionId As Long
    Static rows() As Variant
    Static nRow As Long
    
    If isTitle Then
        nRow = 0
        sectionId = AppendSection(TCSType.tHeads)
        AppendTHCHeadDescription = sectionId
    Else
        Dim data(1) As Variant
        data(0) = LCase(Trim$(GetSafeArrayValue(dataRow, beginCol)))
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

Private Function indexOfDataHead(headName) As Long
    indexOfDataHead = -1
    Dim i As Long
    For i = 0 To m_nDataHead - 1
        If StrComp(m_vDataHeads(i)(DataField.name_), headName, vbTextCompare) = 0 Then
            indexOfDataHead = i + 1: Exit For
        End If
    Next
End Function

Private Function cacheDataHead(ByVal headName As String) As Long
    Dim id As Long: id = indexOfDataHead(headName)
    If (id < 0) Then
        ReDim Preserve m_vDataHeads(m_nDataHead)
        Dim vItem As Variant: vItem = New_OutputField()
        vItem(OutputField.id) = m_nDataHead
        vItem(OutputField.name_) = headName
        m_vDataHeads(m_nDataHead) = vItem
        cacheDataHead = m_nDataHead
        m_nDataHead = m_nDataHead + 1
    Else
        cacheDataHead = id
    End If
End Function

Private Function indexOfOutput(fieldName) As Long
    indexOfOutput = -1
    Dim i As Long
    For i = 0 To m_nOutput - 1
        If StrComp(m_vOutputs(i)(OutputField.name_), fieldName, vbTextCompare) = 0 Then
            indexOfOutput = i: Exit For
        End If
    Next
End Function

Private Function cacheOutputHead(fieldName) As Long
    Dim id As Long: id = indexOfOutput(fieldName)
    If (id < 0) Then
        ReDim Preserve m_vOutputs(m_nOutput)
        Dim vItem As Variant: vItem = New_OutputField()
        vItem(OutputField.id) = m_nOutput
        vItem(OutputField.name_) = fieldName
        m_vOutputs(m_nOutput) = vItem
        cacheOutputHead = m_nOutput
        m_nOutput = m_nOutput + 1
    Else
        cacheOutputHead = id
    End If
End Function

Private Function parseExpression(expression) As Variant
    Dim ret As Variant: ret = New_TExp()
    
    If m_exp.Parse(expression) Then
        Dim funcs As Variant
        Dim i As Long
        Dim idx As Long
        
        funcs = m_exp.FuncList
        
        For i = LBound(funcs) To UBound(funcs)
            Select Case funcs(i)(ExpArgu.type_)
                'Case EAType.Const_
                Case EAType.Field
                    idx = indexOfDataHead(funcs(i)(ExpArgu.Value))
                    If idx < 0 Then
                        cacheMissField funcs(i)(ExpArgu.body)
                        funcs(i)(ExpArgu.type_) = EAType.Const_
                        funcs(i)(ExpArgu.Value) = funcs(i)(ExpArgu.Value)
                    Else
                        funcs(i)(ExpArgu.Value) = idx
                    End If
                Case EAType.func
                    If funcs(i)(ExpArgu.funcId) = ExpFunc.f_getreportdate Then
                        funcs(i)(ExpArgu.type_) = EAType.Const_
                        funcs(i)(ExpArgu.Value) = m_reportDate
                    End If
                    'ExpFunc
                'Case EAType.Unknow
                Case EAType.var
                    idx = indexOfDataHead(funcs(i)(ExpArgu.body))
                    If idx < 0 Then
                        funcs(i)(ExpArgu.type_) = EAType.Const_
                        funcs(i)(ExpArgu.Value) = funcs(i)(ExpArgu.body)
                    Else
                        funcs(i)(ExpArgu.type_) = EAType.Field
                        funcs(i)(ExpArgu.Value) = idx
                    End If
            End Select
        Next
        
        ret(TExp.Text) = expression
        ret(TExp.FuncList) = funcs
        ret(TExp.MainFunc) = m_exp.MainFunc
        ret(TExp.FuncCount) = m_exp.FuncCount
    Else
        cacheParseErr "[" & expression & "] " & m_exp.errDesc
    End If
    
    parseExpression = ret
End Function

Private Sub parseFakedFields(vData)
    Dim prop As ArrayProp
    If Not CheckArray(vData, prop) Then Exit Sub
    Dim i As Long
    Dim vParseExp
    For i = prop.lb To prop.ub
        vParseExp = parseExpression(vData(i)(1))
        vParseExp(TExp.outId) = cacheDataHead(vData(i)(0))
        vData(i) = vParseExp
    Next
End Sub

Private Sub parseNormalFields(vData)
    Dim prop As ArrayProp
    If Not CheckArray(vData, prop) Then Exit Sub
    Dim i As Long
    Dim vParseExp
    For i = prop.lb To prop.ub
        vParseExp = parseExpression(vData(i)(1))
        vParseExp(TExp.outId) = cacheOutputHead(vData(i)(0))
        vData(i) = vParseExp
    Next
End Sub

Private Sub parseConditionalField(vData)
    Dim prop As ArrayProp
    If Not CheckArray(vData, prop) Then Exit Sub
    
    Dim headProp As ArrayProp
    If Not CheckArray(vData(TCondition.heads), headProp) Then Exit Sub
    
    Dim bodyProp As ArrayProp
    If Not CheckArray(vData(TCondition.body), bodyProp) Then Exit Sub
    
    Dim idx As Long
    If vData(TCondition.hasValidField) Then
        idx = indexOfDataHead(vData(TCondition.ValidField))
        If idx < 0 Then
            cacheMissField vData(TCondition.ValidField)
            vData(TCondition.hasValidField) = False
        Else
            vData(TCondition.ValidField) = idx
        End If
    End If
    
    Dim iHead As Long
    Dim iRow As Long
    Dim arrSwitches()
    Dim arrOuts()
    Dim arrCases(): ReDim arrCases(bodyProp.lb To bodyProp.ub)
    Dim arrOutExps(): ReDim arrOutExps(bodyProp.lb To bodyProp.ub)
    Dim isDataHead As Boolean
    
    For iHead = headProp.lb To headProp.ub
        idx = indexOfDataHead(vData(TCondition.heads)(iHead))
        If idx < 0 Then
            isDataHead = False
            MFunc.VarArrAppend arrOuts, cacheOutputHead(vData(TCondition.heads)(iHead))
        Else
            isDataHead = True
            MFunc.VarArrAppend arrSwitches, idx
        End If
        
        For iRow = bodyProp.lb To bodyProp.ub
            If isDataHead Then
                MFunc.VarArrAppend arrCases(iRow), vData(TCondition.body)(iRow)(iHead)
            Else
                MFunc.VarArrAppend arrOutExps(iRow), parseExpression(vData(TCondition.body)(iRow)(iHead))
            End If
        Next
    Next
    
    vData(TCondition.switchs) = arrSwitches
    vData(TCondition.cases) = arrCases
    vData(TCondition.outs) = arrOuts
    vData(TCondition.outExps) = arrOutExps
End Sub

Private Function ApplyNormalFields() As Boolean
    Dim vSection As Variant
    Dim i As Long, iField As Long
    For i = 0 To m_nSection - 1
        vSection = m_vSections(i)
        If (vSection(TCSection.type_) = TCSType.normals) Then
            parseNormalFields m_vSections(i)(TCSection.detail)
        End If
    Next
End Function

Private Function ApplyConditionalField() As Boolean
    Dim vSection As Variant
    Dim i As Long, iField As Long
    For i = 0 To m_nSection - 1
        vSection = m_vSections(i)
        If (vSection(TCSection.type_) = TCSType.conditions) Then
            parseConditionalField m_vSections(i)(TCSection.detail)
        End If
    Next
End Function
'Private Function ApplyDisplayFields(dataHeads) As Boolean
'Private Function ApplyBalanceField(dataHeads) As Boolean
'Private Function ApplySortOutput(dataHeads) As Boolean
Private Function ApplyFakedFields() As Boolean
    Dim vSection As Variant
    Dim i As Long, iField As Long
    For i = 0 To m_nSection - 1
        vSection = m_vSections(i)
        If (vSection(TCSection.type_) = TCSType.faked) Then
            parseFakedFields m_vSections(i)(TCSection.detail)
        End If
    Next
End Function

Private Function ApplyTHCHeadDescription() As Boolean
    Dim iOutHead As Long
    Dim iHead As Long
    Dim iSec As Long
    Dim vHeads
    
    For iSec = 0 To m_nSection - 1
        If m_vSections(iSec)(TCSection.type_) = TCSType.tHeads Then
        
            vHeads = m_vSections(iSec)(TCSection.detail)
            
            For iOutHead = 0 To m_nOutput - 1
                For iHead = 0 To UBound(vHeads)
                    If LCase(m_vOutputs(iOutHead)(OutputField.name_)) = vHeads(iHead)(0) Then
                        m_vOutputs(iOutHead)(OutputField.Desc) = vHeads(iHead)(1)
                    End If
                Next
            Next
            
        End If
    Next
End Function

'Private Function ApplyOBSItems(dataHeads) As Boolean
'Private Function ApplyAcceptableValues(dataHeads) As Boolean
'Private Function ApplyValueCheckFields(dataHeads) As Boolean
'Private Function ApplyMSRDataSheet(dataHeads) As Boolean
Public Sub Apply(dataHeads)
    Dim prop As ArrayProp
    If Not CheckArray(dataHeads, prop) Then Exit Sub
    
    Dim iField As Long
    Dim vHead As Variant
    
    For iField = prop.lb To prop.ub
        If m_nDataHead = 0 Then
            ReDim m_vDataHeads(m_nDataHead)
        Else
            ReDim Preserve m_vDataHeads(m_nDataHead)
        End If
        vHead = New_OutputField()
        vHead(OutputField.id) = m_nDataHead
        vHead(OutputField.name_) = dataHeads(iField)
        
        m_vDataHeads(m_nDataHead) = vHead
        m_nDataHead = m_nDataHead + 1
    Next
    
    ApplyFakedFields
    ApplyNormalFields
    ApplyConditionalField
    ApplyTHCHeadDescription
End Sub

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

Private Function execFunc(vFuncs, funcId, vInputs, outVal, errDesc As String) As Boolean
    Dim iArgu As Long
    Dim vFunc As Variant: vFunc = vFuncs(funcId)
    Dim arguments()
    Dim i As Long
    
    Select Case vFunc(ExpArgu.type_)
        Case EAType.Field
            outVal = vInputs(vFunc(ExpArgu.Value))
        Case EAType.Const_
            outVal = vFunc(ExpArgu.Value)
        Case EAType.func
            If vFunc(ExpArgu.ArguCount) > 0 Then
                ReDim arguments(vFunc(ExpArgu.ArguCount) - 1)
                For i = 0 To vFunc(ExpArgu.ArguCount) - 1
                    If Not execFunc(vFuncs, vFunc(ExpArgu.arguments)(i), vInputs, arguments(i), errDesc) Then
                        Exit Function
                    End If
                Next
            End If
            
            outVal = CallFunc(vFunc(ExpArgu.FuncAddr), arguments, errDesc)
    End Select
    
    execFunc = True
End Function

Private Function execExp(vParseExp, vInputs, Optional ByVal rowId As Long)
    Dim outVal
    Dim errDesc As String
    If execFunc(vParseExp(TExp.FuncList), vParseExp(TExp.MainFunc), vInputs, outVal, errDesc) Then
        execExp = outVal
    Else
        execExp = "#Error"
        cacheExecErr rowId, vParseExp(TExp.Text), errDesc
    End If
End Function

Private Sub execFakedFields(fields, vInputs, vOutputs, Optional ByVal rowId As Long)
    Dim iField As Long
    Dim vParseExp
    For iField = LBound(fields) To UBound(fields)
        vParseExp = fields(iField)
        vOutputs(vParseExp(TExp.outId)) = execExp(vParseExp, vInputs, rowId)
    Next
End Sub

Private Sub execNormalFields(fields, vInputs, vOutputs, Optional ByVal rowId As Long)
    Dim iField As Long
    Dim vParseExp
    For iField = LBound(fields) To UBound(fields)
        vParseExp = fields(iField)
        vOutputs(vParseExp(TExp.outId)) = execExp(vParseExp, vInputs, rowId)
    Next
End Sub

Private Sub execConditionalField(conditionData, vInputs, vOutputs, Optional ByVal rowId As Long)
    Dim iField As Long
    Dim vParseExp
    'TCondition
    'For iField = LBound(fields) To UBound(fields)
     '   vParseExp = fields(iField)
    '    vOutputs(vParseExp(TExp.outId)) = execExp(vParseExp, vInputs, rowId)
    'Next
End Sub

Public Sub LaunchXLSRow(vInputs, vOutputs, Optional ByVal rowId As Long)
    Dim vSec
    Dim iSec As Long
    
    'calc faked fields
    For iSec = 0 To m_nSection - 1
        vSec = m_vSections(iSec)
        If vSec(TCSection.type_) = TCSType.faked Then
            execFakedFields vSec(TCSection.detail), vInputs, vInputs, rowId
        End If
    Next
    
    For iSec = 0 To m_nSection - 1
        vSec = m_vSections(iSec)
        
        'execute normal fields
        If vSec(TCSection.type_) = TCSType.normals Then
            execNormalFields vSec(TCSection.detail), vInputs, vOutputs, rowId
        End If
        
        'execute conditional fields
        If vSec(TCSection.type_) = TCSType.conditions Then
            execConditionalField vSec(TCSection.detail), vInputs, vOutputs, rowId
        End If
    Next
End Sub

Public Property Get vSections() As Variant
    vSections = m_vSections
End Property

Public Property Get Count() As Long
    Count = m_nSection
End Property

Public Property Get outputHeads() As Variant
    outputHeads = m_vOutputs
End Property

Public Property Get OutputHeadCount() As Long
    OutputHeadCount = m_nOutput
End Property

Public Property Get dataHeads() As Variant
    dataHeads = m_vDataHeads
End Property

Public Property Get DataHeadCount() As Long
    DataHeadCount = m_nDataHead
End Property
