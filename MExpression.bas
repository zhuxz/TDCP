Attribute VB_Name = "MExpression"
Option Explicit

Public Enum EAType
    Const_ 'const
    Field '
    func 'function
    var 'variable
    Unknow '
End Enum

Public Enum ExpArgu
    begin_
    id
    body
    MaskBody
    type_
    FuncAddr
    funcName
    funcId
    arguments
    ArguCount
    Value
    end_
End Enum

Public Enum ConfigCell
    begin_
    expType
    parseRet
    execId
    end_
End Enum

Public Enum OperatorExp
    BOF_
    argu1_leftSpace
    argu1
    argu1_rightSpace
    Operator
    argu2_leftSpace
    argu2
    argu2_rightSpace
    EOF_
End Enum

Public Function NewExpArgument()
    Dim arr(ExpArgu.begin_ + 1 To ExpArgu.end_ - 1)
    NewExpArgument = arr
End Function

Public Function NewExpFunction(ByVal Addr As Long, ByVal name_ As String, Optional ByVal ArgumentCount As Long = -1)
    Dim arr(ExpArgu.begin_ + 1 To ExpArgu.end_ - 1)
    arr(ExpArgu.type_) = EAType.func
    arr(ExpArgu.FuncAddr) = Addr
    arr(ExpArgu.funcName) = name_
    arr(ExpArgu.ArguCount) = ArgumentCount
    NewExpFunction = arr
End Function

Public Function NewOperatorExp()
    Dim arr() As Variant
    ReDim arr(OperatorExp.BOF_ + 1 To OperatorExp.EOF_ - 1) As Variant
    NewOperatorExp = arr
End Function

Public Function FindPairStr(src As String, Start As Long, pairLeft As String, pairRight As String, posLeft As Long, posRight As Long) As Boolean
    Dim nLeft As Long
    Dim char As String
    Dim iChr As Long
    Dim nLen As Long: nLen = Len(src)
    
    posLeft = InStr(Start, src, pairLeft)
    posRight = InStr(Start, src, pairRight)
    
    If posLeft > 0 And posRight > 0 And posRight > posLeft Then
        nLeft = 1
        iChr = posLeft + 1
        posRight = 0
        Do
            char = Mid$(src, iChr, 1)
            If char = pairLeft Then
                nLeft = nLeft + 1
            ElseIf char = pairRight Then
                nLeft = nLeft - 1
                If nLeft = 0 Then
                    posRight = iChr
                    Exit Do
                End If
            End If
            iChr = iChr + 1
        Loop While iChr <= nLen
        
        If posRight > posLeft Then
            FindPairStr = True
        Else
            FindPairStr = False
        End If
    Else
        FindPairStr = False
    End If
End Function

Public Function TrimSpace(src As String, posLeft As Long, posRight As Long) As String
    Dim char As String
    Dim pos As Long: pos = 1
    Dim nLen As Long: nLen = Len(src)
    
    If Trim$(src) = "" Then
        posLeft = nLen
        posRight = nLen
        TrimSpace = ""
        Exit Function
    End If
    
    posLeft = 1
    posRight = nLen
    
    For pos = 1 To nLen
        If Mid$(src, pos, 1) <> " " Then
            posLeft = pos: Exit For
        End If
    Next
    
    For pos = nLen To 1 Step -1
        If Mid$(src, pos, 1) <> " " Then
            posRight = pos: Exit For
        End If
    Next
    
    TrimSpace = Mid$(src, posLeft, posRight - posLeft + 1)
End Function

Public Function FindOpratorRev(src As String, Start As Long) As Long
    Dim ret As Long
    Dim i As Long
    Dim nLen As Long: nLen = Len(src)
    
    For i = Start To 1 Step -1
        Select Case Mid$(src, i, 1)
            Case ",": ret = i
            Case EFN_PLUS: ret = i
            Case EFN_MINUS: ret = i
            Case EFN_MULTIPLY: ret = i
            Case EFN_DIVIDE: ret = i
            Case EFN_JOIN_STR: ret = i
            Case EFN_EQUAL: ret = i
            Case EFN_SMALLER: ret = i
            Case EFN_LARGER: ret = i
        End Select
        If ret > 0 Then Exit For
    Next
    
    FindOpratorRev = ret
End Function

