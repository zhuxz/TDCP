Attribute VB_Name = "MTDCP"
Option Explicit

Public Const UI_MARGIN = 60

Public Const SHEET_CONFIG As String = "config"

Public Const DEBUG_FILE As String = "000000"

Public Enum TField
    BOF_
    id
    type_
    Name_
    Desc
    src
    parseRet
    EOF_
End Enum

Public Enum TFieldType
    src
    output
    errTHC
End Enum

Public Function IsDebugApp() As Boolean
    If Dir(App.Path & "\" & DEBUG_FILE) = DEBUG_FILE Then
        IsDebugApp = True
    Else
        IsDebugApp = False
    End If
End Function

Public Function NewField(Optional ByVal id As Long, _
    Optional ByVal type_ As Long = TFieldType.output, _
    Optional ByVal Name_ As String, _
    Optional ByVal Desc As String, _
    Optional ByVal src As String, _
    Optional ByVal parseRet As Variant = Empty) As Variant
    Dim ret(TField.BOF_ + 1 To TField.EOF_ - 1) As Variant
    ret(TField.type_) = type_
    ret(TField.Name_) = Name_
    ret(TField.Desc) = Desc
    ret(TField.src) = src
    ret(TField.parseRet) = parseRet
End Function

Public Function DataConfig2Pathfile( _
    SourceData As Variant, _
    ByVal reportdate As Variant, _
    ByVal PathFileName As Variant, _
    ByVal ConfigFilePath As Variant, _
    ByVal BuilderFilePath As Variant, _
    ByRef ReturnValue As Variant, _
    Optional ByRef pSrcXML As Variant, _
    Optional ByRef pExtraXML As Variant, _
    Optional ByRef bPathfile As Boolean = False)

End Function

Private Function readExcelConfig(ConfigSheet As Worksheet)
    Dim data As Variant: data = MExcel.GetSafeSheetValues(ConfigSheet)
    
    readExcelConfig = data
End Function
