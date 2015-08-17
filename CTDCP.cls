VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTDCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isDebug As Boolean
'Private m_xlsApp As Excel.Application
Private m_config As CConfig

Private Sub Class_Initialize()
    m_isDebug = MTDCP.IsDebugApp()
    
    'Set m_config = New CConfigs
End Sub

Private Sub Class_Terminate()
    Set m_config = Nothing
End Sub

Public Function BuildPathFile( _
    SourceData As Variant, _
    ByVal reportdate As Variant, _
    ByVal PathFileName As Variant, _
    ByVal ConfigFilePath As Variant, _
    ByVal BuilderFilePath As Variant, _
    ByRef ReturnValue As Variant, _
    Optional ByRef pSrcXML As Variant, _
    Optional ByRef pExtraXML As Variant, _
    Optional ByRef bPathfile As Boolean = False)
    
    On Error GoTo eh
    
    BuildPathFile = False
eh:
End Function

Private Function getDataHead(oSheet As Worksheet, ByVal startRow As Long)
    Dim MaxBlankColumn As Long: MaxBlankColumn = 50
    Dim nColumn As Long: nColumn = oSheet.UsedRange.Columns.Count
    Dim iCol As Long
    Dim cellVal As Variant
    Dim vHeads As Variant
    Dim nBlankColumn As Long: nBlankColumn = 0
    Dim nLastColumn As Long: nLastColumn = 0
    Dim ret()
    
    With oSheet
        Do
            cellVal = MExcel.GetExcelErrorValue(Trim$(CStr(.Cells(startRow, iCol).Value)))
            If Len(cellVal) = 0 Then
                nBlankColumn = nBlankColumn + 1
                If nBlankColumn > MaxBlankColumn Then Exit Do
            Else
                nBlankColumn = 0
                nLastColumn = iCol
            End If
            
            MFunc.VarArrAppend vHeads, cellVal
            
            iCol = iCol + 1
            If iCol > nColumn Then Exit Do
        Loop
        
        If nLastColumn > 0 Then
            ReDim ret(1 To nLastColumn)
            For iCol = 1 To nLastColumn
                ret(iCol) = vHeads(iCol - 1)
            Next
            getDataHead = ret
        End If
    End With
End Function

Public Function Build(ByVal dataPath As String, ByVal configPath As String) As Boolean
On Error GoTo eh
    Dim xlsApp As Excel.Application
    MExcel.LoadExcelApp xlsApp
    
    Dim xlsWB As Excel.Workbook: Set xlsWB = xlsApp.Workbooks.Open(configPath, , True)
    Dim xlsWS As Excel.Worksheet: Set xlsWS = MExcel.GetExcelSheet(xlsWB, SHEET_CONFIG)
    Dim srcData As Variant: srcData = MExcel.GetSafeSheetValues(xlsWS, 100, 100)
    Dim oConfig As New CConfig: oConfig.PreviewData srcData
    
    xlsWB.Close
    
    Dim oDataConfig As CDataConfig
    Dim dataHeads As Variant
    Dim errDesc As String
    Dim lastCol As Long
    Dim oBuilder As CBuilder
    Dim dataRowVals
    Dim outHeads()
    Dim outRowVals()
    
    Set xlsWB = xlsApp.Workbooks.Open(configPath, , True)
    For Each xlsWS In xlsWB.Sheets
        Set oDataConfig = oConfig.GetDataConfig(Trim$(xlsWS.Name), errDesc)
        
        
        If Len(errDesc) = 0 Then
            dataHeads = getDataHead(xlsWS, oDataConfig.m_srcBeginRow)
            lastCol = UBound(dataHeads)
            
            oDataConfig.m_reportDate = Date
            oDataConfig.Apply dataHeads
            
            outHeads = oDataConfig.OutputHeads
            ReDim outRowVals(oDataConfig.OutputHeadCount)
            
            oBuilder.SetConfig oDataConfig.vSections
            
            oBuilder.Launch dataRowVals, outRowVals
            'oDataConfig.vSections
        Else
            
        End If
        
        With xlsWS
            'for icol = 1 to odataconfig.m_srcBeginRow
            'MFunc.VarArrAppend dataHeads
        End With
        'oDataConfig.Apply()
    Next
Exit Function
eh:
End Function

Public Function DataConfig2Pathfile(ByVal dataPath As String, ByVal configPath As String, xlsWBPathfile As Excel.Workbook)
    'xlsWBData As Excel.Workbook, _
    xlsWBConfig As Excel.Workbook, xlsWBPathfile As Excel.Workbook
End Function



