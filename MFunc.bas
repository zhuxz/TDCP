Attribute VB_Name = "MFunc"
Option Explicit

Public Const TYPE_ERROR As String = "Error"
Public Const TYPE_STRING As String = "String"
Public Const TYPE_EMPTY As String = "Empty"

'Public Enum ArrayProp
'    BOF_
'    lb 'lbound
'    ub 'ubound
'    Size 'length
'    isArr
'    EOF_
'End Enum

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
            ReDim SourceArr(nLen)
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

Public Sub VarArrAppend(SourceArr, item)
    Dim nLen As Long
    If (IsArray(SourceArr)) Then
        nLen = UBound(SourceArr)
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
