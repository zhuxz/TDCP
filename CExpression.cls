VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum EConst
    begin_
    FieldMark_L '[
    FieldMark_R ']
    FieldMark_Func '_F
    ConstMark_L '"
    ConstMark_R '"
    ConstMark_Func '_C
    FuncMark_L '(
    FuncMark_R ')
    ArgumentMask '___agmt___
    ErrMissFieldMark_L 'miss [
    ErrMissFieldMark_R 'miss ]
    ErrMissConstMark_L 'miss "
    ErrMissConstMark_R 'miss "
    ErrMissFuncMark_L 'miss (
    ErrMissFuncMark_R 'miss )
    end_
End Enum

Private m_result As Long

Private m_arguments()
Private m_argumentCount As Long

Private m_errDesc As String

Private Sub Class_Initialize()
    m_argumentCount = 0
End Sub

Private Sub Class_Terminate()
    ''''
End Sub

Public Property Get Result() As Long
    Result = m_result
End Property

Public Property Get errDesc() As Variant
    errDesc = m_errDesc
End Property

Private Function Con(ByVal ConId As EConst)
    Static consts As Variant
    If Not IsArray(consts) Then
        Dim arr(EConst.begin_ + 1 To EConst.end_ - 1)
        arr(EConst.FieldMark_L) = "["
        arr(EConst.FieldMark_R) = "]"
        arr(EConst.FieldMark_Func) = "_F"
        arr(EConst.ConstMark_L) = """"
        arr(EConst.ConstMark_R) = """"
        arr(EConst.ConstMark_Func) = "_C"
        arr(EConst.FuncMark_L) = "("
        arr(EConst.FuncMark_R) = ")"
        arr(EConst.ArgumentMask) = "?" '"{_Argument_}"
        arr(EConst.ErrMissFieldMark_L) = "miss '['"
        arr(EConst.ErrMissFieldMark_R) = "miss ']'"
        arr(EConst.ErrMissConstMark_L) = "miss '""'"
        arr(EConst.ErrMissConstMark_R) = "miss '""'"
        arr(EConst.ErrMissFuncMark_L) = "miss '('"
        arr(EConst.ErrMissFuncMark_R) = "miss ')'"
        consts = arr
    End If
    
    Con = consts(ConId)
End Function

Private Function appendArgument(Optional ByVal body_ As String = "", _
    Optional ByVal type_ As EAType = EAType.Const_, _
    Optional ByVal funcAddr_ As Long = -1, _
    Optional ByVal funcName_ As String = "", _
    Optional ByVal arguCount_ As Long = -1, _
    Optional ByVal value_ As Variant = Empty) As Variant
    
    Dim vArgu: vArgu = NewExpArgument
    
    vArgu(ExpArgu.id) = m_argumentCount
    vArgu(ExpArgu.body) = body_
    vArgu(ExpArgu.MaskBody) = Con(ArgumentMask) & m_argumentCount
    vArgu(ExpArgu.type_) = type_
    vArgu(ExpArgu.FuncAddr) = funcAddr_
    vArgu(ExpArgu.funcName) = funcName_
    vArgu(ExpArgu.Value) = value_
    vArgu(ExpArgu.ArguCount) = arguCount_
    
    If m_argumentCount = 0 Then
        ReDim m_arguments(0)
    Else
        ReDim Preserve m_arguments(m_argumentCount)
    End If
    m_arguments(m_argumentCount) = vArgu
    m_argumentCount = m_argumentCount + 1
    
    appendArgument = vArgu
End Function

Private Function appendArgumentEx(ByRef vArgument As Variant)
    vArgument(ExpArgu.id) = m_argumentCount
    vArgument(ExpArgu.MaskBody) = Con(ArgumentMask) & m_argumentCount
    
    If m_argumentCount = 0 Then
        ReDim m_arguments(0)
    Else
        ReDim Preserve m_arguments(m_argumentCount)
    End If
    m_arguments(m_argumentCount) = vArgument
    m_argumentCount = m_argumentCount + 1
    
    appendArgumentEx = vArgument
End Function

'Private Function updateArgument(ByVal Id As Long, _
'    Optional ByVal body_ As Variant = Empty, _
'    Optional ByVal Type_ As Variant = Empty, _
'    Optional ByVal funcAddr_ As Variant = Empty, _
'    Optional ByVal funcName_ As Variant = Empty, _
'    Optional ByVal arguments_ As Variant = Empty, _
'    Optional ByVal arguCount_ As Variant = Empty, _
'    Optional ByVal value_ As Variant = Empty) As Variant
'    If Id < m_argumentCount Or Id >= m_argumentCount Then Exit Function
'    If Not IsEmpty(body_) Then m_arguments(Id)(ExpArgu.Body) = body_
'    If Not IsEmpty(Type_) Then m_arguments(Id)(ExpArgu.Type_) = Type_
'    If Not IsEmpty(funcAddr_) Then m_arguments(Id)(ExpArgu.FuncAddr) = funcAddr_
'    If Not IsEmpty(funcName_) Then m_arguments(Id)(ExpArgu.funcName) = funcName_
'    If Not IsEmpty(arguCount_) Then m_arguments(Id)(ExpArgu.ArguCount) = arguCount_
'    If Not IsEmpty(Arguments) Then m_arguments(Id)(ExpArgu.Arguments) = Arguments
'    If Not IsEmpty(value_) Then m_arguments(Id)(ExpArgu.Value) = value_
'    updateArgument = m_arguments(Id)
'End Function
'
'Private Function updateArgumentEx(ByVal Id As Long, ByVal Property As ExpArgu, ByVal NewValue As Variant) As Variant
'    If Id < m_argumentCount Or Id >= m_argumentCount Then Exit Function
'    m_arguments(Id)(Property) = NewValue
'    updateArgumentEx = m_arguments(Id)
'End Function

Private Function getArgumentId(ByVal Expression As String) As Long
    getArgumentId = -1
    Dim expTmp As String: expTmp = Expression
    Dim nLen As Long: nLen = Len(Con(ArgumentMask))
    If Left$(expTmp, nLen) = Con(ArgumentMask) Then
        Dim arguId As String: arguId = Mid$(expTmp, nLen + 1)
        If IsNumeric(arguId) Then getArgumentId = CLng(arguId)
    End If
End Function

Private Function parseArgumentExp(ByVal Expression As String)
    Dim chrPos As Long: chrPos = InStr(1, Expression, Con(ArgumentMask))
    
    If chrPos > 0 Then
        Dim arguId As Long: arguId = CLng(Mid$(Expression, chrPos + Len(Con(ArgumentMask))))
        Dim arguFunc As Variant: arguFunc = m_arguments(arguId)
        Dim ret As Variant
        
        If arguFunc(ExpArgu.type_) = EAType.func Then
            Dim funcName As String
            Dim arguments As Variant
            Dim arguExp As String
            Dim i As Long
            
            funcName = Trim$(Left$(Expression, chrPos - 1))
            arguExp = Trim$(arguFunc(ExpArgu.body))
            If funcName = "" Then
                If (Left$(arguExp, 1) = Con(FuncMark_L) And Right$(arguExp, 1) = Con(FuncMark_R)) Then
                    ret = doParse(Mid$(arguExp, 2, Len(arguExp) - 2))
                Else
                    ret = doParse(arguExp)
                End If
            Else
                Dim f As Variant: f = MExpFuncs.GetExpFuncByName(funcName)
                If IsArray(f) Then
                    arguFunc(ExpArgu.funcName) = f(ExpArgu.funcName)
                    arguFunc(ExpArgu.FuncAddr) = f(ExpArgu.FuncAddr)
                    arguFunc(ExpArgu.ArguCount) = f(ExpArgu.ArguCount)
                    
                    arguExp = Trim$(Mid$(arguExp, 2, Len(arguExp) - 2))
                    
                    If Len(arguExp) > 0 Then
                        arguments = Split(arguExp, ",")
                        For i = 0 To UBound(arguments)
                            arguments(i) = doParse(Trim$(arguments(i)))(ExpArgu.id)
                        Next
                        arguFunc(ExpArgu.arguments) = arguments
                    Else
                        If arguFunc(ExpArgu.ArguCount) > 0 Then
                            m_errDesc = "miss argument(s) for function '" & funcName & "'"
                        End If
                    End If
                    
                    m_arguments(arguId) = arguFunc
                    ret = arguFunc
                Else
                    arguFunc(ExpArgu.type_) = EAType.Const_
                    arguFunc(ExpArgu.Value) = arguFunc(ExpArgu.body)
                    
                    m_arguments(arguId) = arguFunc
                    ret = arguFunc
                    
                    m_errDesc = "unknow function '" & funcName & "'"
                End If
            End If
        Else
            ret = arguFunc
        End If
    Else
        ret = appendArgument(body_:=Expression, value_:=Trim$(Expression))
    End If
    
    parseArgumentExp = ret
End Function

Private Function parseMinusExp(ByVal Expression As String) ', ByVal Operator As String, ByVal OperatorFunc As ExpFunc
    Dim argu As Variant
    argu = MExpFuncs.GetExpFunc(ExpFunc.f_minus)
    argu(ExpArgu.body) = Expression
    
    Dim arr As Variant: arr = Split(Expression, EFN_MINUS)
    Dim n As Long: n = UBound(arr)
    Dim arguments() As Long: ReDim arguments(n) As Long
    Dim i As Long
    
    For i = 0 To n
        arguments(i) = doParse(Trim$(arr(i)))(ExpArgu.id)
    Next
    
    argu(ExpArgu.arguments) = arguments

    parseMinusExp = appendArgumentEx(argu)
End Function

Private Function parseOperatorExp(ByVal Expression As String, ByVal Operator As String, ByVal OperatorFunc As ExpFunc)
    Dim argu As Variant
    argu = MExpFuncs.GetExpFunc(OperatorFunc)
    argu(ExpArgu.body) = Expression
    
    Dim arr As Variant: arr = Split(Expression, Operator)
    Dim n As Long: n = UBound(arr)
    Dim arguments() As Long: ReDim arguments(n) As Long
    Dim i As Long
    
    For i = 0 To n
        arguments(i) = doParse(Trim$(arr(i)))(ExpArgu.id)
    Next
    
    argu(ExpArgu.arguments) = arguments

    parseOperatorExp = appendArgumentEx(argu)
End Function

Private Function parseArgumentList(ByVal Expression As String) As Variant
    Dim arrArguments As Variant: arrArguments = Split(Expression, ",")
    Dim i As Long
    Dim arrArgumentIds() As Long: ReDim arrArgumentIds(UBound(arrArguments)) As Long
    Dim expTemp As String
    For i = 0 To UBound(arrArguments)
        expTemp = Trim$(arrArguments(i))
        doParse2 expTemp
        arrArgumentIds(i) = getArgumentId(expTemp)
    Next
    parseArgumentList = arrArgumentIds
End Function

Private Function parseFunciton(ByRef Expression As String) As Boolean
    Dim expTmp As String: expTmp = Expression
    Dim lPos As Long
    Dim rPos As Long
    Dim pos As Long
    
    Dim char As String
    Dim vFunc As Variant
    Dim funcExp As String
    Dim funcName As String
    Dim arguments As String
    
    Do
        If FindPairStr(expTmp, 1, Con(FuncMark_L), Con(FuncMark_R), lPos, rPos) Then
            pos = FindOpratorRev(expTmp, lPos)
            Do
                pos = pos + 1
                char = Mid(expTmp, pos, 1)
            Loop While char = " "
            
            funcExp = Trim$(Mid$(expTmp, pos, rPos - pos + 1))
            funcName = Trim$(Mid$(expTmp, pos, lPos - pos))
            arguments = Trim$(Mid$(expTmp, lPos + 1, rPos - lPos - 1))
            
            If (funcName = "") Then
                vFunc = MExpFuncs.GetExpFuncByName(EFN_VOID)
            Else
                vFunc = MExpFuncs.GetExpFuncByName(funcName)
            End If
            
            If IsArray(vFunc) Then
                doParse2 arguments
                parseArgumentList arguments
                vFunc(ExpArgu.arguments) = arguments
                vFunc(ExpArgu.body) = funcExp
                appendArgumentEx vFunc
                expTmp = Left$(expTmp, pos - 1) & vFunc(ExpArgu.MaskBody) & Mid$(expTmp, rPos + 1)
            Else
                m_errDesc = "unknow function '" & funcName & "'"
                Exit Do
            End If
        Else
            If lPos = 0 And rPos = 0 Then
                Exit Do
            ElseIf lPos > 0 And rPos = 0 Then
                m_errDesc = Con(ErrMissFuncMark_R)
                Exit Do
            ElseIf lPos = 0 And rPos > 0 Then
                m_errDesc = Con(ErrMissFuncMark_L)
                Exit Do
            Else
                m_errDesc = Con(ErrMissFuncMark_L)
                Exit Do
            End If
        End If
    Loop
    
    Expression = expTmp
    
    If (Len(m_errDesc) > 0) Then
        parseFunciton = False
    Else
        parseFunciton = True
    End If
End Function

Private Function parseOperator(ByRef Expression As String, ByVal OperatorPos As Long) As Boolean
    Dim arr() As String: ReDim arr(OperatorExp.BOF_ + 1 To OperatorExp.EOF_ - 1) As String
    Dim posLeft1 As Long, posLeft2 As Long
    Dim posRight1 As Long, posRight2 As Long
    
    arr(OperatorExp.Operator) = Mid$(Expression, OperatorPos, 1)
    
    arr(OperatorExp.argu1) = TrimSpace(Left$(Expression, OperatorPos - 1), posLeft1, posRight1)
    If posLeft1 > 0 Then arr(OperatorExp.argu1_leftSpace) = Left$(Expression, posLeft1 - 1)
    If OperatorPos > posRight1 Then arr(OperatorExp.argu1_rightSpace) = Mid$(Expression, posRight1 + 1, OperatorPos - posRight1 - 1)
    
    arr(OperatorExp.argu2) = TrimSpace(Mid$(Expression, OperatorPos + 1), posLeft2, posRight2)
    If posLeft2 > 0 Then arr(OperatorExp.argu2_leftSpace) = Mid$(Expression, OperatorPos + 1, posLeft2 - 1)
    arr(OperatorExp.argu2_rightSpace) = Mid$(Expression, OperatorPos + posRight2 + 1)
    
    Dim vFunc As Variant
    Dim argumentIds(1) As Long
    
    If (Len(arr(OperatorExp.argu1)) = 0 And Len(arr(OperatorExp.argu2)) = 0) Then
        m_errDesc = "miss parameter near '" & arr(OperatorExp.Operator) & "'"
    ElseIf (Len(arr(OperatorExp.argu1)) > 0 And Len(arr(OperatorExp.argu2)) = 0) Then
        m_errDesc = "miss parameter near '" & arr(OperatorExp.Operator) & "'"
    ElseIf (Len(arr(OperatorExp.argu1)) = 0 And Len(arr(OperatorExp.argu2)) > 0) Then
        If arr(OperatorExp.Operator) = EFN_PLUS Or arr(OperatorExp.Operator) = EFN_MINUS Then
            vFunc = appendArgument("")
            arr(OperatorExp.argu1) = vFunc(ExpArgu.MaskBody)
            argumentIds(0) = getArgumentId(arr(OperatorExp.argu1))
            If doParse2(arr(OperatorExp.argu2)) Then argumentIds(1) = getArgumentId(arr(OperatorExp.argu2))
        Else
            m_errDesc = "miss parameter near '" & Operator & "'"
        End If
    Else
        If doParse2(arr(OperatorExp.argu1)) Then argumentIds(0) = getArgumentId(arr(OperatorExp.argu1))
        If doParse2(arr(OperatorExp.argu2)) Then argumentIds(1) = getArgumentId(arr(OperatorExp.argu2))
    End If
    
    If Len(m_errDesc) > 0 Then
        parseOperator = False
    Else
        vFunc = MExpFuncs.GetExpFuncByName(arr(OperatorExp.Operator))
        vFunc(ExpArgu.body) = Expression
        vFunc(ExpArgu.arguments) = argumentIds
        appendArgumentEx vFunc
        Expression = vFunc(ExpArgu.MaskBody)
        parseOperator = True
    End If
End Function

Private Function parseLogicOperator(ByVal Expression As String) As String
    Dim expTmp As String: expTmp = Expression
    Dim pos As Long
    Dim char As String
    Dim hasError As Boolean: hasError = False
    Dim vFunc As Variant
    Dim arguExp1 As String, arguExp2 As String
    Dim posLeft1 As Long, posLeft2 As Long
    Dim posRight1 As Long, posRight2 As Long
    Dim isPos As Boolean
    Dim argumentIds(1) As Long
    Do
        isPos = False
        pos = pos + 1
        char = Mid$(Expression, pos, 1)
        If char = EFN_LARGER Then
            isPos = True
        ElseIf char = EFN_SMALLER Then
            isPos = True
        ElseIf char = EFN_EQUAL Then
            isPos = True
        End If
        
        If isPos Then
            arguExp1 = TrimSpace(Left$(Expression, pos - 1), posLeft1, posRight1)
            arguExp2 = TrimSpace(Mid$(Expression, pos + 1), posLeft2, posRight2)
            
            If (Len(arguExp1) = 0 Or Len(arguExp1) = 0) Then
                m_errDesc = "miss parameter near '" & char & "'"
                Exit Do
            Else
                arguExp1 = doParse(arguExp1)
                
            End If
            
            vFunc = MExpFuncs.GetExpFuncByName(char)
            
'            arguExp1 = TrimSpace(Left$(Expression, pos), posLeft1, posRight1)
'            arguExp = doParse(arguExp)
'            argumentIds(0) = getArgumentId(arguExp)
'            expTmp = Left(Expression, posLeft) & arguExp
'
'            arguExp = TrimSpace(Mid$(Expression, pos), posLeft, posRight)
'            arguExp = doParse(parseLogicOperator(arguExp))
'            argumentIds(1) = getArgumentId(arguExp)
        End If
    Loop
    
    parseLogicOperator = expTmp
End Function

Private Function doParse(ByVal Expression As String)
    Dim ret
    
    If InStr(1, Expression, EFN_JOIN_STR) > 0 Then
        ret = parseOperatorExp(Expression, EFN_JOIN_STR, ExpFunc.f_join_str)
    ElseIf InStr(1, Expression, EFN_PLUS) > 0 Then
        ret = parseOperatorExp(Expression, EFN_PLUS, ExpFunc.f_plus)
    ElseIf InStr(1, Expression, EFN_MINUS) > 0 Then
        ret = parseMinusExp(Expression)
    ElseIf InStr(1, Expression, EFN_MULTIPLY) > 0 Then
        ret = parseOperatorExp(Expression, EFN_MULTIPLY, ExpFunc.f_multiply)
    ElseIf InStr(1, Expression, EFN_DIVIDE) > 0 Then
        ret = parseOperatorExp(Expression, EFN_DIVIDE, ExpFunc.f_divide)
    ElseIf InStr(1, Expression, EFN_SMALLER) > 0 Then
        ret = parseOperatorExp(Expression, EFN_SMALLER, ExpFunc.f_smaller)
    ElseIf InStr(1, Expression, EFN_LARGER) > 0 Then
        ret = parseOperatorExp(Expression, EFN_LARGER, ExpFunc.f_larger)
    ElseIf InStr(1, Expression, EFN_EQUAL) > 0 Then
        ret = parseOperatorExp(Expression, EFN_EQUAL, ExpFunc.f_equal)
    Else
        ret = parseArgumentExp(Expression)
    End If
    
    doParse = ret
End Function

Private Function doParse2(ByRef Expression As String)
    'function
    If Not parseFunciton(Expression) Then Exit Function
    
    '&
    Dim pos As Long
    
    pos = InStrRev(Expression, EFN_JOIN_STR)
    If pos > 0 Then
        If Not parseOperator(Expression, pos) Then Exit Function
    End If
    
    '=, > , <
    Dim nLen As Long: nLen = Len(Expression)
    Dim i As Long
    Dim char As String
    Dim bParse As Boolean: bParse = True
    
    For i = nLen To 1 Step -1
        char = Mid$(Expression, i, 1)
        If char = EFN_EQUAL Then
            bParse = parseOperator(Expression, i)
            Exit For
        ElseIf char = EFN_LARGER Then
            bParse = parseOperator(Expression, i)
            Exit For
        ElseIf char = EFN_SMALLER Then
            bParse = parseOperator(Expression, i)
            Exit For
        End If
    Next
    If (Not bParse) Then Exit Function
    
    '+, -
    nLen = Len(Expression)
    bParse = True
    For i = nLen To 1 Step -1
        char = Mid$(Expression, i, 1)
        If char = EFN_PLUS Then
            bParse = parseOperator(Expression, i)
            Exit For
        ElseIf char = EFN_MINUS Then
            bParse = parseOperator(Expression, i)
            Exit For
        End If
    Next
    If (Not bParse) Then Exit Function
    
    '*, /
    nLen = Len(Expression)
    bParse = True
    For i = nLen To 1 Step -1
        char = Mid$(Expression, i, 1)
        If char = EFN_MULTIPLY Then
            bParse = parseOperator(Expression, i)
            Exit For
        ElseIf char = EFN_DIVIDE Then
            bParse = parseOperator(Expression, i)
            Exit For
        End If
    Next
    
    If (Not bParse) Then Exit Function
    
    If getArgumentId(Expression) < 0 Then
        Dim vFunc As Variant
        vFunc = appendArgument(Expression, EAType.Unknow)
        Expression = vFunc(ExpArgu.MaskBody)
    End If
    
    doParse2 = True
End Function

Private Function maskConsts(ByVal Expression As String) As String
    Dim exp_len As Long: exp_len = Len(Expression)
    Dim u_exp As String: u_exp = UCase(Expression)
    Dim vArg As Variant
    Dim iChr As Long: iChr = 1
    Dim chrPos As Long
    Dim chr2Pos As Long
    Dim nConstFunc As Long: nConstFunc = Len(Con(ConstMark_Func))
    Dim nFieldFunc As Long: nFieldFunc = Len(Con(FieldMark_Func))
    Dim strChr As String
    Dim ret() As String: ReDim ret(1 To exp_len) As String
    Dim partExp As String
    Dim constVal As String
    
    Do
        If Len(m_errDesc) > 0 Then Exit Do
        If iChr > exp_len Then Exit Do
        
        If Mid(u_exp, iChr, 1) = Con(ConstMark_L) Then
            '' const value mark by ""
            chrPos = InStr(iChr + 1, u_exp, Con(ConstMark_R))
            If chrPos > iChr Then
                partExp = Mid$(Expression, iChr, chrPos - iChr + 1)
                constVal = Trim$(Mid$(Expression, iChr + 1, chrPos - iChr - 1))
                vArg = appendArgument(body_:=partExp, value_:=constVal)
                ret(iChr) = vArg(ExpArgu.MaskBody)
                iChr = chrPos + 1
            Else
                m_errDesc = Con(ErrMissConstMark_R)
            End If
        ElseIf StrComp(Mid$(u_exp, iChr, nConstFunc), Con(ConstMark_Func), vbTextCompare) = 0 Then
            '' const value mark by _C("")
            chrPos = iChr + nConstFunc
            If MFunc.NextTrimChar(u_exp, chrPos, exp_len) = Con(FuncMark_L) Then
                chr2Pos = InStr(chrPos, u_exp, Con(FuncMark_R))
                If chr2Pos > chrPos Then
                    partExp = Mid$(Expression, iChr, chr2Pos - iChr + 1)
                    constVal = Trim$(Mid$(Expression, chrPos + 1, chr2Pos - chrPos - 1))
                    vArg = appendArgument(body_:=partExp, value_:=constVal)
                    ret(iChr) = vArg(ExpArgu.MaskBody)
                    iChr = chr2Pos + 1
                Else
                    m_errDesc = Con(ErrMissFuncMark_R)
                End If
            Else
                ret(iChr) = Mid(Expression, iChr, 1)
                iChr = iChr + 1
            End If
        ElseIf Mid(u_exp, iChr, 1) = Con(FieldMark_L) Then
            '' const field mark by []
            chrPos = InStr(iChr + 1, u_exp, Con(FieldMark_R))
            If chrPos > iChr Then
                partExp = Mid$(Expression, iChr, chrPos - iChr + 1)
                constVal = Trim$(Mid$(Expression, iChr + 1, chrPos - iChr - 1))
                vArg = appendArgument(body_:=partExp, type_:=EAType.Field, value_:=constVal)
                ret(iChr) = vArg(ExpArgu.MaskBody)
                iChr = chrPos + 1
            Else
                m_errDesc = Con(ErrMissFuncMark_R)
            End If
        ElseIf StrComp(Mid$(u_exp, iChr, nFieldFunc), Con(FieldMark_Func), vbTextCompare) = 0 Then
            '' const field mark by _F("")
            chrPos = iChr + nFieldFunc
            If MFunc.NextTrimChar(u_exp, chrPos, exp_len) = Con(FuncMark_L) Then
                chr2Pos = InStr(chrPos, u_exp, Con(FuncMark_R))
                If chr2Pos > chrPos Then
                    partExp = Mid$(Expression, iChr, chr2Pos - iChr + 1)
                    constVal = Trim$(Mid$(Expression, chrPos + 1, chr2Pos - chrPos - 1))
                    vArg = appendArgument(body_:=partExp, type_:=EAType.Field, value_:=constVal)
                    ret(iChr) = vArg(ExpArgu.MaskBody)
                    iChr = chr2Pos + 1
                Else
                    m_errDesc = Con(ErrMissFuncMark_R)
                End If
            Else
                ret(iChr) = Mid(Expression, iChr, 1)
                iChr = iChr + 1
            End If
        Else
            ret(iChr) = Mid(Expression, iChr, 1)
            iChr = iChr + 1
        End If
    Loop
    
    maskConsts = Join(ret, "")
End Function

Private Function maskFunctions(ByVal Expression As String) As String
    Dim exp_len As Long
    Dim vArg As Variant
    Dim iChr As Long: iChr = 1
    Dim strChr As String
    Dim iFunc As Long: iFunc = 0
    Dim nFunc As Long: nFunc = 0
    Dim posPair() As Long
    Dim partExp As String
    
    Do
        exp_len = Len(Expression)
        If iChr > exp_len Then Exit Do
        
        strChr = Mid$(Expression, iChr, 1)
        
        If strChr = Con(FuncMark_L) Then
            iFunc = iFunc + 1
            ReDim Preserve posPair(1 To iFunc) As Long
            posPair(iFunc) = iChr
            iChr = iChr + 1
        ElseIf strChr = Con(FuncMark_R) Then
            If iFunc < 1 Then
                m_errDesc = Con(ErrMissFuncMark_L)
                Exit Do
            Else
                partExp = Mid$(Expression, posPair(iFunc), iChr - posPair(iFunc) + 1)
                vArg = appendArgument(body_:=partExp, type_:=EAType.func)
                Expression = Left$(Expression, posPair(iFunc) - 1) & vArg(ExpArgu.MaskBody) & Mid$(Expression, iChr + 1)
                iChr = posPair(iFunc) + Len(vArg(ExpArgu.MaskBody)) - 1
                iFunc = iFunc - 1
            End If
        Else
            iChr = iChr + 1
        End If
    Loop
    
    If iFunc > 0 Then m_errDesc = Con(ErrMissFuncMark_L)
    
    maskFunctions = Expression
End Function

Public Function Parse(ByVal Expression As String)
On Error GoTo eh

    m_argumentCount = 0
    m_errDesc = ""
    
    Dim stdExp As String
    stdExp = maskConsts(Trim$(Expression))
    'doParse stdExp
    'stdExp = maskFunctions(stdExp) m_arguments
    If (doParse2(stdExp)) Then
        m_result = getArgumentId(stdExp)
        If m_result < 0 Then
            Dim vFunc As Variant
            vFunc = appendArgument(stdExp, EAType.Unknow)
            m_result = vFunc(ExpArgu.id)
        End If
    Else
        m_result = -1
    End If
    'm_result = getArgumentId(doParse(stdExp)) '(ExpArgu.id)
Exit Function
eh:
    If Len(m_errDesc) = 0 Then m_errDesc = Err.Description
    Err.Clear
End Function

Private Function createArray(Optional ByVal DataType As Long = 1, Optional ByVal Start As Long = 0) As CArray
    Dim ret As New CArray
    ret.type_ = DataType
    ret.StartPos = Start
    Set createArray = ret
End Function

Private Function argumentToXML(ByVal arguId As Long) As String
    Dim oArr As CArray
    Dim Argument As Variant: Argument = m_arguments(arguId)
    Dim tag As String
    
    If IsArray(Argument) Then
        Set oArr = createArray(4)
        
        With oArr
            Select Case Argument(ExpArgu.type_)
                Case EAType.Const_
                    tag = "Const"
                Case EAType.Field
                    tag = "Field"
                Case EAType.func
                    tag = "Function"
                Case EAType.Unknow
                    tag = "Unknow"
                Case EAType.var
                    tag = "Var"
            End Select
            
            .AppendStrItem "<" & tag
            .AppendStrItem " id=""" & Argument(ExpArgu.id) & """"
            .AppendStrItem " name=""" & CXml(Argument(ExpArgu.funcName)) & """"
            .AppendStrItem " addr=""" & Argument(ExpArgu.FuncAddr) & """"
            .AppendStrItem " value=""" & CXml(Argument(ExpArgu.Value)) & """"
            .AppendStrItem " body=""" & CXml(Argument(ExpArgu.body)) & """"
            .AppendStrItem ">"
            
            If IsArray(Argument(ExpArgu.arguments)) Then
                Dim i As Long
                For i = 0 To UBound(Argument(ExpArgu.arguments))
                    .AppendStrItem argumentToXML(Argument(ExpArgu.arguments)(i))
                Next
            End If
            
            .AppendStrItem "</" & tag & ">"
            
            argumentToXML = Join(.StrList, "")
        End With
    Else
        argumentToXML = Argument
    End If
End Function

Private Function revertArguments(ByVal arguId As Long)
    Dim argus As Variant
    argus = m_arguments(arguId)(ExpArgu.arguments)
End Function

Public Function ToXML() As String
    ToXML = argumentToXML(m_result)
End Function
