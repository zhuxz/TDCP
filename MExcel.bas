Attribute VB_Name = "MExcel"
Option Explicit

Public Const XLS_ERROR2042_STR As String = "Error 2042"
Public Const XLS_ERROR2042 As String = "#N/A"
Public Const XLS_ERROR2007_STR As String = "Error 2007"
Public Const XLS_ERROR2007 As String = "#DIV/0!"

Public Const XLS_MAX_COLUMN As Long = 200
Public Const XLS_MAX_BlankRow As Long = 200

Public m_xlsApp As Excel.Application

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

Public Function Int2ABC(ByVal IntVal As Long) As String
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
    
    Int2ABC = ret
End Function

Public Function GetExcelApp() As Excel.Application
    On Error GoTo eh:
    Set GetExcelApp = GetObject(, "Excel.Application")
    Exit Function
eh:
    Set GetExcelApp = CreateObject("Excel.Application")
End Function

Public Function GetXLSApp() As Excel.Application
    If m_xlsApp Is Nothing Then
        Set m_xlsApp = New Excel.Application
        m_xlsApp.Visible = IsDebugApp()
    End If
    Set GetXLSApp = m_xlsApp
End Function

Public Function GetExcelSheet(ExcelBook As Excel.Workbook, ByVal ExcelSheetName As String) As Excel.Worksheet
    On Error GoTo eh
    Set GetExcelSheet = ExcelBook.Sheets(ExcelSheetName)
eh:
    Err.Clear
End Function

Public Function IsSheetBlankRow(SheetRow As Range, Optional ByRef SheetRowValues As Variant) As Boolean
    Dim arr As Variant
    Dim arrLen As Long
    Dim ret() As Variant
    Dim iCol As Long
    Dim nBlank As Long
    
    With SheetRow
        arr = .Value2
        arrLen = UBound(arr, 2)
        ReDim ret(1 To arrLen) As Variant
        
        For iCol = 1 To arrLen
            If IsError(arr(1, iCol)) Then
                ret(iCol) = GetExcelErrorValue(CStr(arr(1, iCol)))
            Else
                ret(iCol) = arr(1, iCol)
            End If
            
            If IsEmpty(arr(1, iCol)) Then
                nBlank = nBlank + 1
            ElseIf Trim$(CStr(arr(1, iCol))) = "" Then
                nBlank = nBlank + 1
            End If
        Next
    End With
    
    If Not IsMissing(SheetRowValues) Then
        SheetRowValues = ret
    End If
    
    If nBlank = arrLen Then
        IsSheetBlankRow = True
    End If
End Function

Public Function GetSafeSheetValues(xlsSheet, Optional ByVal MaxBlankRow As Long = -1, Optional ByVal MaxColumnCount As Long = -1)
    Dim maxCol As Long, maxRow As Long
    Dim nBlankRow As Long
    Dim iRow As Long, iCol As Long
    Dim srcRowVals As Variant
    Dim oArr As New CArray: oArr.type_ = 1: oArr.StartPos = 1
    
    With xlsSheet
        maxRow = .UsedRange.row + .UsedRange.rows.Count - 1
        maxCol = .UsedRange.Column + .UsedRange.Columns.Count - 1
        If MaxColumnCount <> -1 Then
            If maxCol > MaxColumnCount Then maxCol = MaxColumnCount
        End If
        
        If MaxBlankRow = -1 Then
            Dim arrRowVals() As Variant
            ReDim arrRowVals(1 To maxCol) As Variant
            
            For iRow = 1 To maxRow
                srcRowVals = .Range("A" & iRow & ":" & Int2ABC(maxCol) & iRow).values2
                
                For iCol = 1 To maxCol
                    If IsError(srcRowVals(1, iCol)) Then
                        arrRowVals(iCol) = GetExcelErrorValue(CStr(srcRowVals(1, iCol)))
                    Else
                        arrRowVals(iCol) = srcRowVals(1, iCol)
                    End If
                Next
                oArr.AppendVarItem arrRowVals
            Next
        Else
            Dim varRowVals As Variant
            
            For iRow = 1 To maxRow
                If IsSheetBlankRow(.Range("A" & iRow & ":" & Int2ABC(maxCol) & iRow), varRowVals) Then
                    nBlankRow = nBlankRow + 1
                    If nBlankRow > MaxBlankRow Then Exit For
                Else
                    nBlankRow = 0
                End If
                oArr.AppendVarItem varRowVals
            Next
        End If
    End With
    
    If oArr.Count > 0 Then GetSafeSheetValues = oArr.List
    
    Set oArr = Nothing
End Function
