Attribute VB_Name = "MFunc"
Option Explicit

Public Const TYPE_ERROR As String = "Error"
Public Const TYPE_STRING As String = "String"
Public Const TYPE_EMPTY As String = "Empty"

Public Enum ArrayProp
    BOF_
    lb 'lbound
    ub 'ubound
    Size 'length
    isArr
    EOF_
End Enum

Public Type ArrayProp_
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

Public Function CheckArray(srcVar, Optional ArrProp) As Boolean
On Error GoTo eh
    Dim ret(ArrayProp.BOF_ + 1 To ArrayProp.EOF_ - 1) As Variant
    ret(ArrayProp.lb) = LBound(srcVar)
    ret(ArrayProp.ub) = UBound(srcVar)
    ret(ArrayProp.Size) = ArrayProp.ub - ArrayProp.lb + 1
    ret(ArrayProp.isArr) = True
    ArrProp = ret
    CheckArray = True
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
