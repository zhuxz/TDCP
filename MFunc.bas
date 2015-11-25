Attribute VB_Name = "MFunc"
Option Explicit

Public Const TYPE_ERROR As String = "Error"
Public Const TYPE_STRING As String = "String"
Public Const TYPE_EMPTY As String = "Empty"

Public Type ArrayProp
    lb As Long
    ub As Long
    Size As Long
    isArr As Boolean
End Type

Public Function PosReplace(ByRef SourceStr As String, ByVal PosStart As Long, ByVal PosEnd As String, ByVal ReplaceStr As String) As String
    SourceStr = Left(SourceStr, PosStart - 1) & ReplaceStr & Mid(SourceStr, PosEnd + 1)
End Function

Public Function NextTrimChar(ByVal SourceStr As String, ByRef PosStart As Long, Optional ByVal MaxPos As Long = -1) As String
    Dim str As String
    If MaxPos = -1 Then MaxPos = Len(SourceStr)
    Do
        str = Trim$(Mid$(SourceStr, PosStart, 1))
        If str = "" Then
            If PosStart >= MaxPos Then Exit Do
        Else
            NextTrimChar = str
            Exit Do
        End If
        PosStart = PosStart + 1
    Loop
End Function

Public Function CXml(ByVal sValue)
    If sValue <> "" Then
        CXml = Replace(Replace(Replace(Replace(Replace(sValue, "&", "&amp;"), "'", "&apos;"), """", "&quot;"), "<", "&lt;"), ">", "&gt;")
    Else
        CXml = ""
    End If
End Function

Public Function CheckArray(srcVar, ArrProp As ArrayProp) As Boolean
On Error GoTo eh
    Dim ret As ArrayProp
    ret.lb = LBound(srcVar)
    ret.ub = UBound(srcVar)
    ret.Size = ret.ub - ret.lb + 1
    ret.isArr = True
    ArrProp = ret
    CheckArray = ret.isArr
Exit Function
eh:
    Err.Clear
End Function

Public Function Var2Long(ByVal var As Variant, Optional ByVal default As Long = 0) As Long
On Error GoTo eh
    Var2Long = CLng(var)
    Exit Function
eh:
    Err.Clear
    Var2Long = default
End Function

Public Function GetSafeArrayValue(SourceArray, indexId) As Variant
On Error GoTo eh
    GetSafeArrayValue = SourceArray(indexId)
eh:
    Err.Clear
End Function

Public Sub RedimVarArr(SourceArr, nLen)
    If (IsArray(SourceArr)) Then
        If (nLen = 0) Then
            ReDim SourceArr(0)
        Else
        ReDim Preserve SourceArr(nLen)
        End If
    Else
        Dim arr(): ReDim arr(0)
        SourceArr = arr
    End If
End Sub

Public Function NewVarArray(Optional begin_ As Long = -1, Optional end_ As Long = -1)
    Dim arr()
    If begin_ < 0 Then
        If end_ < 0 Then
            'NewVarArray = arr
        Else
            ReDim arr(0 To end_)
        End If
    Else
        If end_ < 0 Then
            ReDim arr(begin_ To begin_)
        Else
            ReDim arr(begin_ To end_)
        End If
    End If
    NewVarArray = arr
End Function

Public Function UBoundEx(SourceArr) As Long
    On Error GoTo eh
    UBoundEx = UBound(SourceArr)
    Exit Function
eh:
    Err.Clear
    UBoundEx = -1
End Function

Public Sub VarArrAppend(SourceArr, item)
    Dim nLen As Long
    If (IsArray(SourceArr)) Then
        nLen = UBoundEx(SourceArr)
        If nLen < 0 Then
            nLen = 0
        Else
            nLen = nLen + 1
        End If
        ReDim Preserve SourceArr(nLen)
    Else
        Dim arr(): ReDim arr(0)
        SourceArr = arr
    End If
    SourceArr(nLen) = item
End Sub

Public Function ETLCompare(v1, v2) As Boolean
    If VarType(v1) = VarType(v2) Then
        ETLCompare = (v1 = v2)
    Else
        ETLCompare = (CStr(v1) = CStr(v2))
    End If
End Function

Public Function parseBool(srcVal, parseVal) As Boolean
On Error GoTo eh
    parseVal = CBool(srcVal)
    parseBool = True
    Exit Function
eh:
    Err.Clear
End Function

Public Function parseDate(srcVal, parseVal) As Boolean
On Error GoTo eh
    parseVal = CDate(srcVal)
    parseDate = True
    Exit Function
eh:
    Err.Clear
End Function

Public Function parseDouble(srcVal, parseVal) As Boolean
On Error GoTo eh
    parseVal = CDate(srcVal)
    parseDouble = True
    Exit Function
eh:
    Err.Clear
End Function

Public Function ETLMatch(srcVal, matchVal) As Boolean
    If (matchVal) = "*" Then
        ETLMatch = True
        Exit Function
    End If
    
    Dim srcValType As VbVarType: srcValType = VarType(srcVal)
    Dim matchValType As VbVarType: matchValType = VarType(matchVal)
    Dim srcValTmp
    Dim ret As Boolean
    
    If srcValType = matchValType Then
        If matchValType = VbVarType.vbString Then
            ret = (Trim$(srcVal) = Trim$(matchVal))
        Else
            ret = (srcVal = matchVal)
        End If
    Else
        Select Case VarType(matchVal)
            Case VbVarType.vbEmpty: '0 未初始化（默认）
                ETLMatch = (Trim$(CStr(srcVal)) = "")
            Case VbVarType.vbNull: '1 不包含任何有效数据
                ETLMatch = (Trim$(CStr(srcVal)) = "")
            Case VbVarType.vbInteger: '2 整型子类型
                If parseDouble(srcVal, srcValTmp) Then
                    ret = (srcValTmp = matchVal)
                Else
                    ret = (Trim$(CStr(srcVal)) = CStr(matchVal))
                End If
            Case VbVarType.vbLong: '3 长整型子类型
                If parseDouble(srcVal, srcValTmp) Then
                    ret = (srcValTmp = matchVal)
                Else
                    ret = (Trim$(CStr(srcVal)) = CStr(matchVal))
                End If
            Case VbVarType.vbSingle: '4 单精度子类型
                If parseDouble(srcVal, srcValTmp) Then
                    ret = (srcValTmp = matchVal)
                Else
                    ret = (Trim$(CStr(srcVal)) = CStr(matchVal))
                End If
            Case VbVarType.vbDouble: '5 双精度子类型
                If parseDouble(srcVal, srcValTmp) Then
                    ret = (srcValTmp = matchVal)
                Else
                    ret = (Trim$(CStr(srcVal)) = CStr(matchVal))
                End If
            Case VbVarType.vbCurrency: '6 货币子类型
                If parseDouble(srcVal, srcValTmp) Then
                    ret = (srcValTmp = matchVal)
                Else
                    ret = (Trim$(CStr(srcVal)) = CStr(matchVal))
                End If
            Case VbVarType.vbDate: '7 日期或时间值
                If parseDate(srcVal, srcValTmp) Then
                    ret = (srcValTmp = matchVal)
                Else
                    ret = (Trim$(CStr(srcVal)) = CStr(matchVal))
                End If
            Case VbVarType.vbString: '8 字符串值
                ret = (Trim$(CStr(srcVal)) = Trim$(matchVal))
            Case VbVarType.vbObject: '9 字符串子类型
                ret = False
            Case VbVarType.vbError: '10 错误子类型
                ret = False
            Case VbVarType.vbBoolean: '11 Boolean 子类型
                If parseBool(srcVal, srcValTmp) Then
                    ret = (srcValTmp = matchVal)
                Else
                    ret = (Trim$(CStr(srcVal)) = CStr(matchVal))
                End If
            Case VbVarType.vbVariant: '12 Variant （仅用于变量数组）
                ret = False
            Case VbVarType.vbDataObject: '13 数据访问对象
                ret = False
            Case VbVarType.vbDecimal: '14 十进制子类型
                If parseDouble(srcVal, srcValTmp) Then
                    ret = (srcValTmp = matchVal)
                Else
                    ret = (Trim$(CStr(srcVal)) = CStr(matchVal))
                End If
            Case VbVarType.vbByte: '17 字节子类型
                ret = False
            Case VbVarType.vbArray: '8192 数组
                ret = False
        End Select
    End If
    
    ETLMatch = ret
End Function

Public Function IndexOfStr(SourceArray, item, Optional CompMethod As VBA.VbCompareMethod = vbTextCompare) As Long
    Dim ret As Long: ret = -1
    If IsArray(SourceArray) Then
        Dim i As Long
        For i = LBound(SourceArray) To UBound(SourceArray)
            If StrComp(SourceArray(i), item, CompMethod) = 0 Then
                ret = i: Exit For
            End If
        Next
    End If
    IndexOfStr = ret
End Function

Public Function Max(v1, v2)
    If (v1 > v2) Then
        Max = v1
    Else
        Max = v2
    End If
End Function

Public Function Min(v1, v2)
    If (v1 < v2) Then
        Min = v1
    Else
        Min = v2
    End If
End Function
