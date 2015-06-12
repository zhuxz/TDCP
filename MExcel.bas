Attribute VB_Name = "MExcel"
Option Explicit

Public Const XLS_ERROR2042_STR As String = "Error 2042"
Public Const XLS_ERROR2042 As String = "#N/A"
Public Const XLS_ERROR2007_STR As String = "Error 2007"
Public Const XLS_ERROR2007 As String = "#DIV/0!"

Public Const XLS_MAX_COLUMN As Long = 200

Public Function IsSheetBlankRow(xlsRow As Range, Optional ByRef SheetRowValues As Variant) As Boolean
    Dim arr As Variant
    Dim iCol As Long
    Dim str As String
    Dim xlsVal As Variant
    
    With xlsRow
        arr = .Value2
        For iCol = 1 To UBound(arr, 2)
            xlsVal = arr(1, iCol)
            If VarType(arr(1, iCol)) = vbError Then
                str = CStr(xlsVal)
                If str = XLS_ERROR2042_STR Then
                    xlsVal = XLS_ERROR2042
                ElseIf xlsVal = XLS_ERROR2007_STR Then
                    xlsVal = XLS_ERROR2007
                End If
            End If
        Next
    End With
End Function

Public Function GetExcelErrorValue(ByVal ErrorStr As String) As String
    Dim ret As String
    Select Case ErrorStr
        Case "Error 2000": ret = "#NULL!"
        Case "Error 2007": ret = "#DIV/0!"
        Case "Error 2015": ret = "#VALUE!"
        Case "Error 2023": ret = "#REF!"
        Case "Error 2029": ret = "#NAME?"
        Case "Error 2036": ret = "#NUM!"
        Case "Error 2042": ret = "#N/A"
        Case "Error 2043": ret = "#GETTING_DATA"
        Case Else: ret = ErrorStr
    End Select
    GetExcelErrorValue = ret
End Function

Public Function Int2XlsColumn(ByVal IntVal As Long) As String
    Dim re As Long
    Dim ret As String
    Do
        If IntVal < 1 Then Exit Do
        
        re = IntVal Mod 26
        IntVal = (IntVal - re) / 26
        
        If re = 0 Then
            IntVal = IntVal - 1
            ret = "Z" & ret
        Else
            ret = Chr(64 + re) & ret
        End If
    Loop
    
    Int2XlsColumn = ret
End Function

Public Function GetSheetValues(xlsSheet, Optional ByVal MaxBlankColumn As Long = 50, Optional ByVal MaxBlankRow As Long = 100)
    Dim sht As Worksheet
    
    Dim maxCol As Long, maxRow As Long
    Dim nBlankRow As Long, nBlankColumn As Long
    Dim iRow As Long, iCol As Long
    Dim xlsRowVals As Variant
    Dim oArr As New CArray: oArr.Type_ = 1: oArr.StartPos = 1
    Dim arr() As Variant
    
    With sht
        maxRow = .UsedRange.Row + .UsedRange.Rows.Count - 1
        maxCol = .UsedRange.Column + .UsedRange.Columns.Count - 1
        If maxCol > XLS_MAX_COLUMN Then maxCol = XLS_MAX_COLUMN
        
        ReDim arr(1 To maxCol) As Variant
        
        For iRow = 1 To maxRow
            xlsRowVals = .Range(iRow & "A:" & iRow & Int2XlsColumn(maxCol)).Value2
            
            For iCol = 1 To UBound(xlsRowVals, 2)
                If IsError(xlsRowVals(1, iCol)) Then
                    arr(iCol) = GetExcelErrorValue(CStr(xlsRowVals(1, iCol)))
                Else
                    arr(iCol) = arr(1, iCol)
                End If
                
                oArr.AppendVarItem (arr)
            Next
        Next
    End With
End Function
