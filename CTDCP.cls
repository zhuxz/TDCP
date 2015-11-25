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

Private m_vDataErrs() As Variant
Private m_nDataErr As Long

Private m_xmlLog As String

Private Sub Class_Initialize()
    m_isDebug = MTDCP.IsDebugApp()
End Sub

Private Sub Class_Terminate()
    '''
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
    Dim iCol As Long: iCol = 1
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

Private Sub cacheDataError(errInfo As String, Optional rowId As Long, Optional columnId As Long)
    If m_nDataErr = 0 Then
        ReDim m_vDataErrs(m_nDataErr)
    Else
        ReDim Preserve m_vDataErrs(m_nDataErr)
    End If
    m_vDataErrs(m_nDataErr) = "(" & rowId & ", " & columnId & ")" & errInfo
    m_nDataErr = m_nDataErr + 1
End Sub

Public Function BuildSinglePathFileToSheet(shtData As Excel.Worksheet, _
    oDataConfig As CDataConfig, shtOutput As Excel.Worksheet, _
    Optional ByVal THCCommonHeads As Variant = Empty)
    Dim dataHeads: dataHeads = getDataHead(shtData, oDataConfig.m_srcStart)
    Dim nLastCol As Long: nLastCol = UBound(dataHeads)
    
    oDataConfig.m_reportDate = #9/30/2015# 'Date
    oDataConfig.Apply dataHeads
    
    Dim iRow As Long, nLastRow As Long
    Dim iCol As Long
    Dim nBlankCell As Long, nBlankRow As Long
    Dim shtCellVals
    Dim inputRowCells
    Dim outputRowCells
    Dim outputRows(), nOutputRow As Long
    
    With shtData
        nLastRow = .UsedRange.row + .UsedRange.rows.Count - 1
        inputRowCells = MFunc.NewVarArray(0, oDataConfig.DataHeadCount - 1)
        outputRowCells = MFunc.NewVarArray(0, oDataConfig.OutputHeadCount - 1)
        
        For iRow = oDataConfig.m_srcStart + 1 To nLastRow
            nBlankCell = 0
            shtCellVals = .Range("A" & iRow & ":" & Int2ABC(nLastCol) & iRow).Value2
            
            For iCol = 1 To nLastCol
                If isEmpty(shtCellVals(1, iCol)) Then
                    nBlankCell = nBlankCell + 1
                    inputRowCells(iCol - 1) = Empty
                Else
                    If IsError(shtCellVals(1, iCol)) Then
                        inputRowCells(iCol - 1) = MExcel.GetExcelErrorValue(CStr(shtCellVals(1, iCol)))
                        cacheDataError CStr(shtCellVals(1, iCol)), iRow, iCol
                    Else
                        inputRowCells(iCol - 1) = shtCellVals(1, iCol)
                    End If
                End If
            Next
            
            If nBlankCell = nLastCol Then
                nBlankRow = nBlankRow + 1
                If nBlankRow > 500 Then Exit For
            Else
                nBlankRow = 0
                If iRow = 110 Then
                    Debug.Print iRow
                End If
                oDataConfig.LaunchXLSRow inputRowCells, outputRowCells, iRow
                If nOutputRow = 0 Then
                    ReDim outputRows(nOutputRow)
                Else
                    ReDim Preserve outputRows(nOutputRow)
                End If
                outputRows(nOutputRow) = outputRowCells
                nOutputRow = nOutputRow + 1
            End If
        Next
    End With
    
    'populate
    Dim outputMatrix()
    Dim outputHeads
    Dim iTHCHead
    Dim startRow As Long: startRow = 1
    
    If nOutputRow > 0 Then
        With shtOutput
            iRow = startRow
            .Cells(iRow, 1).Value = "Group"
            .Cells(iRow, 2).Value = "Yes"
            
            iRow = iRow + 1
            .Cells(iRow, 1).Value = "Report Date: "
            .Cells(iRow, 2).Value = oDataConfig.m_reportDate
            
            iRow = iRow + 1
            outputHeads = oDataConfig.outputHeads
            If isEmpty(THCCommonHeads) Then
                For iCol = 1 To oDataConfig.OutputHeadCount
                    .Cells(iRow, iCol).Value = outputHeads(iCol - 1)(OutputField.name_)
                    .Cells(iRow + 1, iCol).Value = outputHeads(iCol - 1)(OutputField.Desc)
                Next
            Else
                For iCol = 0 To oDataConfig.OutputHeadCount - 1
                    .Cells(iRow, iCol + 1).Value = outputHeads(iCol)(OutputField.name_)
                    If isEmpty(outputHeads(iCol)(OutputField.Desc)) Then
                        For iTHCHead = 0 To UBound(THCCommonHeads)
                            If LCase(outputHeads(iCol)(OutputField.name_)) = THCCommonHeads(iTHCHead)(0) Then
                                outputHeads(iCol)(OutputField.Desc) = THCCommonHeads(iTHCHead)(1)
                                .Cells(iRow + 1, iCol + 1).Value = outputHeads(iCol)(OutputField.Desc)
                            End If
                        Next
                    Else
                        .Cells(iRow + 1, iCol + 1).Value = outputHeads(iCol)(OutputField.Desc)
                    End If
                Next
            End If
        End With
            
        ReDim outputMatrix(1 To nOutputRow, 1 To oDataConfig.OutputHeadCount)
        For iRow = 0 To nOutputRow - 1
            For iCol = 0 To oDataConfig.OutputHeadCount - 1
                outputMatrix(iRow + 1, iCol + 1) = outputRows(iRow)(iCol)
            Next
        Next
        
        startRow = 5
        With shtOutput
            .Range("A" & startRow & ":" & Int2ABC(oDataConfig.OutputHeadCount) & (nOutputRow + startRow - 1)).Value2 = outputMatrix
        End With
    End If
End Function

Public Function Build(ByVal dataPath As String, ByVal configPath As String) As Boolean
On Error GoTo eh
    Dim xlsApp As Excel.Application
    MExcel.LoadExcelApp xlsApp
    
    Dim wbConfig As Excel.Workbook: Set wbConfig = xlsApp.Workbooks.Open(configPath, , True)
    Dim xlsWS As Excel.Worksheet: Set xlsWS = MExcel.GetExcelSheet(wbConfig, SHEET_CONFIG)
    Dim srcData As Variant: srcData = MExcel.GetSafeSheetValues(xlsWS, 100, 100)
    Dim oConfig As New CConfig: oConfig.PreviewData srcData
    Dim wbData As Excel.Workbook: Set wbData = xlsApp.Workbooks.Open(dataPath, , True)
    Dim oDataConfig As CDataConfig
    Dim errDesc As String
    Dim wbOutput As Excel.Workbook: Set wbOutput = xlsApp.Workbooks.Add
    Dim shtOutput As Excel.Worksheet
    Dim xmlLog() As String
    
    For Each xlsWS In wbData.Sheets
        Set oDataConfig = oConfig.GetDataConfig(Trim$(xlsWS.Name), errDesc)
        If Len(errDesc) = 0 Then
            Set shtOutput = wbOutput.Sheets(1)
            BuildSinglePathFileToSheet xlsWS, oDataConfig, shtOutput, oConfig.THCHeads
            
            VarArrAppend xmlLog, "<data name=""" & xlsWS.Name & """>"
            VarArrAppend xmlLog, oDataConfig.GetExecErrXML()
            VarArrAppend xmlLog, oDataConfig.GetMissFieldXML()
            VarArrAppend xmlLog, oDataConfig.GetParseErrXML()
            VarArrAppend xmlLog, "</data>"
        Else
            
        End If
    Next
    
    If UBound(xmlLog) >= 0 Then
        m_xmlLog = Join(xmlLog, "")
    End If
eh:
    If Err.Number = 0 Then
        Exit Function
    Else
        Debug.Print Err.Description
        'Resume Next
    End If
    
    Set xlsWS = Nothing
    
    If (Not wbData Is Nothing) Then
        wbData.Close
        Set wbData = Nothing
    End If
    
    If (Not wbConfig Is Nothing) Then
        wbConfig.Close
        Set wbConfig = Nothing
    End If
    
'    If (Not wbOutput Is Nothing) Then
'        wbOutput.Close
'        Set wbOutput = Nothing
'    End If
'
'    If (Not xlsApp Is Nothing) Then
'        xlsApp.Quit
'    End If
End Function

Public Function DataConfig2Pathfile(ByVal dataPath As String, ByVal configPath As String, xlsWBPathfile As Excel.Workbook)
    'xlsWBData As Excel.Workbook, _
    xlsWBConfig As Excel.Workbook, xlsWBPathfile As Excel.Workbook
End Function

Public Property Get xmlLog() As String
    xmlLog = m_xmlLog
End Property



