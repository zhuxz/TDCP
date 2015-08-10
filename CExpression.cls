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

Private Function getArgumentId(ByVal Expression As String) As Long
    getArgumentId = -1
    Dim expTmp As String: expTmp = Expression
    Dim nLen As Long: nLen = Len(Con(ArgumentMask))
    If Left$(expTmp, nLen) = Con(ArgumentMask) Then
        Dim arguId As String: arguId = Mid$(expTmp, nLen + 1)
        If IsNumeric(arguId) Then getArgumentId = CLng(arguId)
    End If
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
                parseFunciton arguments
                vFunc(ExpArgu.arguments) = parseArgumentList(arguments)
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

Private Function doParse2(ByRef Expression As String) As Boolean
    If getArgumentId(Expression) >= 0 Then
        doParse2 = True
        Exit Function
    End If
    
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

Public Function Parse(ByVal Expression As String)
On Error GoTo eh
    m_argumentCount = 0
    m_errDesc = ""
    
    Dim stdExp As String
    stdExp = maskConsts(Trim$(Expression))
    If (doParse2(stdExp)) Then
        m_result = getArgumentId(stdExp)
    Else
        m_result = -1
    End If
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
            .AppendStrItem " name=""" & Argument(ExpArgu.funcName) & """"
            .AppendStrItem " addr=""" & Argument(ExpArgu.FuncAddr) & """"
            .AppendStrItem " value=""" & Argument(ExpArgu.Value) & """"
            .AppendStrItem " body=""" & Argument(ExpArgu.body) & """"
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
