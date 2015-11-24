Attribute VB_Name = "MConfig"
Option Explicit

Public Enum TValueCheck
    BOF_
    srcField
    srcFieldId
    Desc
    rule
    data
    uppers
    lowers
    EOF_
End Enum

Public Enum VCRule
    BOF_
    MoreThan
    LessThan
    Between
    NULL_
    NOTNULL
    Custom
    equal
    EOF_
End Enum

Public Enum TCondition
    BOF_
    ValidField
    hasValidField
    ValidHeadId
    heads
    body
    switchs
    cases
    outs ' output heads
    outExps ' output expressions
    EOF_
End Enum

Public Enum TSort
    BOF_
    heads
    body
    EOF_
End Enum

Public Enum TAccept
    BOF_
    srcField
    srcFieldId
    body
    EOF_
End Enum

Public Enum TCSType 'config section type
    BOF_
    normals
    conditions
    display
    balanceHead
    Sort
    faked
    tHeads
    obs
    accepts
    Validation
    EOF_
End Enum

Public Enum TCSection 'config section
    BOF_
    type_
    detail
    EOF_
End Enum

Public Enum TDataConfig
    BOF_
    srcName
    srcStart
    destName
    destStart
    isOptional
    beginRow
    endRow
    normals
    conditions
    display
    balanceHead
    Sort
    faked
    tHeads
    obs
    accepts
    Validation
    EOF_
End Enum

Public Enum TConfig
    BOF_
    allTHeads
    MSRDataSheet
    ConfigMain
'    srcName
'    srcStart
'    destName
'    destStart
'    isOptional
'    beginRow
'    endRow
'    normals
'    conditions
'    display
'    balanceHead
'    Sort
'    faked
'    tHeads
'    obs
'    accepts
'    Validation
    
    data
    EOF_
End Enum

Public Type UCaseKeyWords
    FieldsMap As String
    RelatedDataSheet As String
    NormalFields As String
    ConditionalField As String
    DisplayFields As String
    BalanceField As String
    SortOutput As String
    FakedFields As String
    THCHeadDescription As String
    OBSItems As String
    AcceptableValues As String
    ValueCheckFields As String
    MSRDataSheet As String
    Optional_ As String
End Type

Public Enum OutputField
    BOF_
    id
    name_
    Desc
    EOF_
End Enum

Public Enum DataField
    BOF_
    id
    name_
    Desc
    EOF_
End Enum

Public Enum TExp
    BOF_
    id
    outId
    Text
    MainFunc
    FuncList
    FuncCount
    EOF_
End Enum

Public Enum ExecErr
    BOF_
    rowId
    expression
    errDesc
    EOF_
End Enum

Public Const KW_FieldsMap As String = "Fields Map"
Public Const KW_RelatedDataSheet As String = "Related Data Sheet:"
Public Const KW_NormalFields As String = "Normal Fields"
Public Const KW_ConditionalField As String = "Conditional Field"
Public Const KW_DisplayFields As String = "DisplayFields"
Public Const KW_BalanceField As String = "BalanceField"
Public Const KW_SortOutput As String = "SortOutput"
Public Const KW_FakedFields As String = "FakedFields"
Public Const KW_THCHeadDescription As String = "THCHeadDescription"
Public Const KW_OBSItems As String = "OBS Items"
Public Const KW_AcceptableValues As String = "AcceptableValues"
Public Const KW_ValueCheckFields As String = "ValueCheckFields"
Public Const KW_MSRDataSheet As String = "MSR Data Sheet"
Public Const KW_Optional As String = "Optional"

Public Function GetUCaseKeyWords() As UCaseKeyWords
    Dim ret As UCaseKeyWords
    With ret
        .FieldsMap = UCase(KW_FieldsMap)
        .RelatedDataSheet = UCase(KW_RelatedDataSheet)
        .NormalFields = UCase(KW_NormalFields)
        .ConditionalField = UCase(KW_ConditionalField)
        .DisplayFields = UCase(KW_DisplayFields)
        .BalanceField = UCase(KW_BalanceField)
        .SortOutput = UCase(KW_SortOutput)
        .FakedFields = UCase(KW_FakedFields)
        .THCHeadDescription = UCase(KW_THCHeadDescription)
        .OBSItems = UCase(KW_OBSItems)
        .AcceptableValues = UCase(KW_AcceptableValues)
        .ValueCheckFields = UCase(KW_ValueCheckFields)
        .MSRDataSheet = UCase(KW_MSRDataSheet)
        .Optional_ = UCase(KW_Optional)
    End With
    GetUCaseKeyWords = ret
End Function

Public Function New_TCondition()
    Dim ret() As Variant: ReDim ret(TCondition.BOF_ + 1 To TCondition.EOF_ - 1) As Variant
    New_TCondition = ret
End Function

Public Function New_TSort()
    Dim ret() As Variant: ReDim ret(TSort.BOF_ + 1 To TSort.EOF_ - 1) As Variant
    New_TSort = ret
End Function

Public Function New_TAccept()
    Dim ret() As Variant: ReDim ret(TAccept.BOF_ + 1 To TAccept.EOF_ - 1) As Variant
    New_TAccept = ret
End Function

Public Function New_TValueCheck()
    Dim ret() As Variant: ReDim ret(TValueCheck.BOF_ + 1 To TValueCheck.EOF_ - 1) As Variant
    New_TValueCheck = ret
End Function

Public Function New_TExp()
    Dim ret() As Variant: ReDim ret(TExp.BOF_ + 1 To TExp.EOF_ - 1) As Variant
    New_TExp = ret
End Function

Public Function New_OutputField()
    Dim ret() As Variant: ReDim ret(OutputField.BOF_ + 1 To OutputField.EOF_ - 1) As Variant
    New_OutputField = ret
End Function

Public Function New_ExecErr()
    Dim ret() As Variant: ReDim ret(ExecErr.BOF_ + 1 To ExecErr.EOF_ - 1) As Variant
    New_ExecErr = ret
End Function

Public Function ParsePredefineStr(ByVal str)
    Dim s As String: s = LCase(CStr(str))
    If s = "empty" Then
        ParsePredefineStr = Empty
    ElseIf s = "*empty" Then
        ParsePredefineStr = Empty
    Else
        ParsePredefineStr = str
    End If
End Function
