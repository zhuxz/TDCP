Attribute VB_Name = "MFunc"
Option Explicit

Public Type ArrayProp
    lb As Long 'lbound
    ub As Long 'ubound
    size As Long 'length
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

Public Function CheckArray(srcVar) As ArrayProp
    On Error GoTo eh
    Dim arrProp As ArrayProp
    With arrProp
        .lb = LBound(srcVar)
        .ub = UBound(srcVar)
        .size = .ub - .lb + 1
    End With
    CheckArray = arrProp
    Exit Function
eh:
    arrProp.ub = -1
    CheckArray = arrProp
    Err.Clear
End Function
