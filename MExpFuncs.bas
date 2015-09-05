Attribute VB_Name = "MExpFuncs"
Option Explicit

Public Const EXP_ERROR As String = "error"
Public Const EXP_WORNING As String = "worning"

Public Const EFN_GETREPORTDATE As String = "GetReportDate"

Public Const EFN_JOIN_STR As String = "&"
Public Const EFN_PLUS As String = "+"
Public Const EFN_MINUS As String = "-"
Public Const EFN_MULTIPLY As String = "*"
Public Const EFN_DIVIDE As String = "/"

Public Const EFN_SMALLER As String = "<"
Public Const EFN_LARGER As String = ">"
Public Const EFN_EQUAL As String = "="
'Public Const EFN_UNEQUAL As String = "!"

Public Const EFN_VOID As String = "Void"
Public Const EFN_IF As String = "If"
Public Const EFN_EDATE As String = "EDate"
Public Const EFN_IS_ERROR As String = "IsError"
Public Const EFN_IS_NUMBER As String = "IsNumber"
Public Const EFN_INT As String = "Int"
Public Const EFN_VALUE As String = "Value"
Public Const EFN_LEFT As String = "Left"
Public Const EFN_MID As String = "Mid"
Public Const EFN_DATE As String = "Date"
Public Const EFN_TEXT As String = "Text"
Public Const EFN_YEAR As String = "Year"
Public Const EFN_MONTH As String = "Month"
Public Const EFN_LEN As String = "Len"

Public Enum ExpFunc
    BOF_
    f_getreportdate

    f_join_str
    
    f_plus
    f_minus
    
    f_multiply
    f_divide
    
    f_smaller
    f_larger
    f_equal
    
    f_void
    f_if
    f_edate
    f_is_number
    f_int
    f_value
    f_left
    f_mid
    f_date
    f_text
    f_year
    f_month
    f_len
    EOF_
End Enum

Public Declare Function CallFunction Lib "user32" Alias "CallWindowProcA" ( _
    ByVal FunctionAddr As Long, _
    ByVal ArgumentsPtr As Long, _
    ByVal ReturnPtr As Long, _
    ByVal ErrDescPtr As Long, _
    ByVal Unused As Long) As Integer

Public Function NewExpFuncList()
    Dim arr() As Variant
    ReDim arr(ExpFunc.BOF_ + 1 To ExpFunc.EOF_ - 1) As Variant
    NewExpFuncList = arr
End Function

Public Function GetExpFuncList() As Variant
    Static expFuncList As Variant
    If Not IsArray(expFuncList) Then
        expFuncList = NewExpFuncList
        expFuncList(ExpFunc.f_getreportdate) = NewExpFunction(AddressOf ef_getreportdate, EFN_GETREPORTDATE, 0)
        expFuncList(ExpFunc.f_join_str) = NewExpFunction(AddressOf ef_join_str, EFN_JOIN_STR, 2)
        expFuncList(ExpFunc.f_plus) = NewExpFunction(AddressOf ef_plus, EFN_PLUS, 2)
        expFuncList(ExpFunc.f_minus) = NewExpFunction(AddressOf ef_minus, EFN_MINUS, 2)
        expFuncList(ExpFunc.f_multiply) = NewExpFunction(AddressOf ef_multiply, EFN_MULTIPLY, 2)
        expFuncList(ExpFunc.f_divide) = NewExpFunction(AddressOf ef_divide, EFN_DIVIDE, 2)
        expFuncList(ExpFunc.f_smaller) = NewExpFunction(AddressOf ef_smaller, EFN_SMALLER, 2)
        expFuncList(ExpFunc.f_larger) = NewExpFunction(AddressOf ef_larger, EFN_LARGER, 2)
        expFuncList(ExpFunc.f_equal) = NewExpFunction(AddressOf ef_equal, EFN_EQUAL, 2)
        expFuncList(ExpFunc.f_void) = NewExpFunction(AddressOf ef_void, EFN_VOID, 1)
        expFuncList(ExpFunc.f_if) = NewExpFunction(AddressOf ef_if, EFN_IF, 3)
        expFuncList(ExpFunc.f_edate) = NewExpFunction(AddressOf ef_edate, EFN_EDATE, 2)
        expFuncList(ExpFunc.f_is_number) = NewExpFunction(AddressOf ef_is_number, EFN_IS_NUMBER, 1)
        expFuncList(ExpFunc.f_int) = NewExpFunction(AddressOf ef_int, EFN_INT, 1)
        expFuncList(ExpFunc.f_value) = NewExpFunction(AddressOf ef_value, EFN_VALUE, 1)
        expFuncList(ExpFunc.f_left) = NewExpFunction(AddressOf ef_left, EFN_LEFT, 2)
        expFuncList(ExpFunc.f_mid) = NewExpFunction(AddressOf ef_mid, EFN_MID, 2)
        expFuncList(ExpFunc.f_date) = NewExpFunction(AddressOf ef_date, EFN_DATE, 3)
        expFuncList(ExpFunc.f_text) = NewExpFunction(AddressOf ef_text, EFN_TEXT, 2)
        expFuncList(ExpFunc.f_year) = NewExpFunction(AddressOf ef_year, EFN_YEAR, 1)
        expFuncList(ExpFunc.f_month) = NewExpFunction(AddressOf ef_month, EFN_MONTH, 1)
        expFuncList(ExpFunc.f_len) = NewExpFunction(AddressOf ef_len, EFN_LEN, 1)
    End If
    GetExpFuncList = expFuncList
End Function

Public Function GetExpFuncByName(ByVal funcName As String) As Variant
    Dim FuncList As Variant: FuncList = GetExpFuncList
    Dim i As Long
    For i = 1 To UBound(FuncList)
        If StrComp(funcName, FuncList(i)(ExpArgu.funcName), vbTextCompare) = 0 Then
            GetExpFuncByName = FuncList(i)
            Exit For
        End If
    Next
End Function

Public Function GetExpFunc(ByVal WhichFunc As ExpFunc) As Variant
    GetExpFunc = GetExpFuncList()(WhichFunc)
End Function

Public Function CallFunc(ByVal FunctionAddr As Long, vArguments As Variant, pErrDesc As String) As Variant
    On Error GoTo eh
    Dim vRet As Variant
    CallFunction FunctionAddr, VarPtr(vArguments), VarPtr(vRet), VarPtr(pErrDesc), 0
    CallFunc = vRet
    Exit Function
eh:
    pErrDesc = EXP_ERROR & ": " & Err.Description
    Err.Clear
End Function

Private Function thc_equal(v1 As Variant, v2 As Variant) As Boolean
    If IsNumeric(v1) Then
        If IsNumeric(v2) Then
            thc_equal = (Val(v1) = Val(v2))
        ElseIf IsDate(v2) Then
            thc_equal = (Val(v1) = CDate(v2))
        Else
            thc_equal = (StrComp(v1, v2, vbTextCompare) = 0)
        End If
    ElseIf IsDate(v1) Then
        If IsNumeric(v2) Then
            thc_equal = (CDate(v1) = Val(v2))
        ElseIf IsDate(v2) Then
            thc_equal = (Val(v1) = Val(v2))
        Else
            thc_equal = (StrComp(v1, v2, vbTextCompare) = 0)
        End If
    Else
        thc_equal = (StrComp(v1, v2, vbTextCompare) = 0)
    End If
End Function

Private Function thc_larger(v1 As Variant, v2 As Variant) As Boolean
    If IsNumeric(v1) Then
        If IsNumeric(v2) Then
            thc_larger = (Val(v1) > Val(v2))
        ElseIf IsDate(v2) Then
            thc_larger = (Val(v1) > CDate(v2))
        Else
            thc_larger = (StrComp(v1, v2, vbTextCompare) > 0)
        End If
    ElseIf IsDate(v1) Then
        If IsNumeric(v2) Then
            thc_larger = (CDate(v1) > Val(v2))
        ElseIf IsDate(v2) Then
            thc_larger = (Val(v1) > Val(v2))
        Else
            thc_larger = (StrComp(v1, v2, vbTextCompare) > 0)
        End If
    Else
        thc_larger = (StrComp(v1, v2, vbTextCompare) > 0)
    End If
End Function

Private Function thc_smaller(v1 As Variant, v2 As Variant) As Boolean
    If IsNumeric(v1) Then
        If IsNumeric(v2) Then
            thc_smaller = (Val(v1) < Val(v2))
        ElseIf IsDate(v2) Then
            thc_smaller = (Val(v1) < CDate(v2))
        Else
            thc_smaller = (StrComp(v1, v2, vbTextCompare) < 0)
        End If
    ElseIf IsDate(v1) Then
        If IsNumeric(v2) Then
            thc_smaller = (CDate(v1) < Val(v2))
        ElseIf IsDate(v2) Then
            thc_smaller = (Val(v1) < Val(v2))
        Else
            thc_smaller = (StrComp(v1, v2, vbTextCompare) < 0)
        End If
    Else
        thc_smaller = (StrComp(v1, v2, vbTextCompare) < 0)
    End If
End Function

Private Function thc_IsBool(v, Optional ByRef vRet As Boolean) As Boolean
    On Error GoTo eh
    vRet = CBool(v)
    thc_IsBool = True
eh: Err.Clear
End Function

Private Function ef_getreportdate(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    pReturn = Date
End Function

Private Function ef_join_str(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    Dim i As Long
    Dim retVal As String
    For i = 0 To UBound(pArguments)
        retVal = retVal & pArguments(i)
    Next
    pReturn = retVal
End Function

Private Function ef_plus(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    Dim i As Long
    Dim retVal As Double
    For i = 0 To UBound(pArguments)
        If i = 0 Then
            If Not IsEmpty(pArguments(i)) Then
                If Trim$(pArguments(i)) <> "" Then
                    retVal = CDbl(pArguments(i))
                End If
            End If
        Else
            retVal = retVal + CDbl(pArguments(i))
        End If
    Next
    pReturn = retVal
End Function

Private Function ef_minus(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    Dim i As Long
    Dim retVal As Double
    For i = 0 To UBound(pArguments)
        If i = 0 Then
            If Not IsEmpty(pArguments(i)) Then
                If Trim$(pArguments(i)) <> "" Then
                    retVal = CDbl(pArguments(i))
                End If
            End If
        Else
            retVal = retVal - CDbl(pArguments(i))
        End If
    Next
    pReturn = retVal
End Function

Private Function ef_if(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    If CBool(pArguments(0)) Then
        pReturn = pArguments(1)
    Else
        pReturn = pArguments(2)
    End If
End Function

Private Function ef_edate(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    Dim lb_ As Long
    lb_ = LBound(pArguments)
    pReturn = DateAdd("m", CDbl(pArguments(lb_ + 1)), CDate(pArguments(lb_)))
End Function

Private Function ef_larger(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    Dim i As Long
    Dim tmp As Boolean
    tmp = thc_larger(pArguments(0), pArguments(1))
    For i = 2 To UBound(pArguments)
        tmp = thc_larger(tmp, pArguments(i))
        'pErrDesc = EXP_WORNING & ": not reasonable expression"
    Next
    pReturn = tmp
End Function

Private Function ef_smaller(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    Dim i As Long
    Dim tmp As Boolean
    tmp = thc_smaller(pArguments(0), pArguments(1))
    For i = 2 To UBound(pArguments)
        tmp = thc_smaller(tmp, pArguments(i))
        'pErrDesc = EXP_WORNING & ": not reasonable expression"
    Next
    pReturn = tmp
End Function

Private Function ef_equal(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    Dim i As Long
    Dim tmp As Boolean
    tmp = thc_equal(pArguments(0), pArguments(1))
    For i = 2 To UBound(pArguments)
        tmp = thc_equal(tmp, pArguments(i))
        'pErrDesc = EXP_WORNING & ": not reasonable expression"
    Next
    pReturn = tmp
End Function

Private Function ef_multiply(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    Dim i As Long
    Dim retVal As Double
    retVal = 1#
    For i = 0 To UBound(pArguments)
        retVal = retVal * CDbl(pArguments(i))
    Next
    pReturn = retVal
End Function

Private Function ef_divide(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    Dim i As Long
    Dim retVal As Double
    retVal = 1# * CDbl(pArguments(0))
    For i = 1 To UBound(pArguments)
        retVal = retVal / pArguments(i)
    Next
    pReturn = retVal
End Function

Private Function ef_iserror(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    pReturn = IsError(pArguments(0))
End Function

Private Function ef_is_number(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    If IsEmpty(pArguments(0)) Then
        pReturn = False
    Else
        pReturn = IsNumeric(pArguments(0))
    End If
End Function

Private Function ef_int(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    pReturn = Int(pArguments(0))
End Function

Private Function ef_value(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    pReturn = Val(pArguments(0))
End Function

Private Function ef_left(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    pReturn = Left$(pArguments(0), pArguments(1))
End Function

Private Function ef_mid(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    Dim ub As Long
    ub = UBound(pArguments)
    If ub < 2 Then
        pReturn = Mid$(pArguments(0), pArguments(1))
    Else
        pReturn = Mid$(pArguments(0), pArguments(1), pArguments(2))
    End If
End Function

Private Function ef_date(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    pReturn = DateSerial(CInt(pArguments(0)), CInt(pArguments(1)), CInt(pArguments(2)))
End Function

Private Function ef_text(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    pReturn = Format(Val(pArguments(0)), pArguments(1))
End Function

Private Function ef_year(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    pReturn = Year(pArguments(0))
End Function

Private Function ef_month(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    pReturn = Format(pArguments(0))
End Function

Private Function ef_void(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    ef_void = pArguments(0)
End Function

Private Function ef_len(pArguments As Variant, pReturn As Variant, pErrDesc As String, Unused As Long) As Integer
    pReturn = Len(pArguments(0))
End Function
