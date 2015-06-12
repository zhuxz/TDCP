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
    With ArrProp
        ret(.lb) = LBound(srcVar)
        ret(.ub) = UBound(srcVar)
        ret(.Size) = .ub - .lb + 1
        ret(.isArr) = True
    End With
    CheckArray = True
Exit Function
eh:
    Err.Clear
End Function
