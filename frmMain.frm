VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Parse Expression"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   15615
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDataPath 
      Height          =   390
      Left            =   2520
      TabIndex        =   6
      Text            =   "Data:"
      Top             =   5520
      Width           =   12615
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox txtConfigPath 
      Height          =   390
      Left            =   2520
      TabIndex        =   4
      Text            =   "Formula:"
      Top             =   4920
      Width           =   12615
   End
   Begin VB.CommandButton cmdReadConfig 
      Caption         =   "Read Config"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton cmdParseExpression 
      Caption         =   "Parse Expression"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtParseResult 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmMain.frx":0000
      Top             =   720
      Width           =   15135
   End
   Begin VB.TextBox txtFormula 
      Height          =   390
      Left            =   2640
      TabIndex        =   0
      Text            =   "Formula:"
      Top             =   120
      Width           =   12615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_xlsApp As Excel.Application
Private m_isDebug As Boolean

Private Sub TerminateForm()
    If Not m_xlsApp Is Nothing Then
        m_xlsApp.Quit
        Set m_xlsApp = Nothing
    End If
End Sub

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

Private Function getDataRow()

End Function

Private Sub cmdBuild_Click()
On Error GoTo eh
    Dim configPath As String: configPath = Trim$(Me.txtConfigPath.Text)
    
    Dim xlsApp As Excel.Application
    MExcel.LoadExcelApp xlsApp
    
    Dim xlsWB As Excel.Workbook: Set xlsWB = xlsApp.Workbooks.Open(configPath, , True)
    Dim xlsWS As Excel.Worksheet: Set xlsWS = MExcel.GetExcelSheet(xlsWB, SHEET_CONFIG)
    Dim srcData As Variant: srcData = MExcel.GetSafeSheetValues(xlsWS, 100, 100)
    Dim oConfig As New CConfig: oConfig.PreviewData srcData
    
    xlsWB.Close
    
    Dim dataPath As String: dataPath = Trim$(Me.txtDataPath.Text)
    Dim oDataConfig As CDataConfig
    Dim dataHeads As Variant
    Dim errDesc As String
    Dim lastCol As Long
    
    Set xlsWB = xlsApp.Workbooks.Open(configPath, , True)
    For Each xlsWS In xlsWB.Sheets
        Set oDataConfig = oConfig.GetDataConfig(Trim$(xlsWS.Name), errDesc)
        
        If Len(errDesc) = 0 Then
            dataHeads = getDataHead(xlsWS, oDataConfig.m_srcBeginRow)
            lastCol = UBound(dataHeads)
            
            oDataConfig.m_reportDate = Date
            oDataConfig.Apply dataHeads
            oDataConfig.vSections
            
            
        Else
            
        End If
        
        With xlsWS
            'for icol = 1 to odataconfig.m_srcBeginRow
            'MFunc.VarArrAppend dataHeads
        End With
        'oDataConfig.Apply()
    Next
Exit Sub
eh:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Terminate()
    TerminateForm
End Sub

Private Sub Form_Resize()
    With Me.cmdParseExpression
        .Top = UI_MARGIN
        .Left = UI_MARGIN
        .Height = Me.txtFormula.Height
    End With
    
    With Me.txtFormula
        .Left = Me.cmdParseExpression.Left + Me.cmdParseExpression.Width + UI_MARGIN
        .Top = Me.cmdParseExpression.Top
        .Width = Me.ScaleWidth - .Left - UI_MARGIN
    End With
    
    With Me.txtParseResult
        .Left = UI_MARGIN
        .Top = Me.cmdParseExpression.Top + Me.cmdParseExpression.Height + UI_MARGIN
        .Width = Me.ScaleWidth - .Left - UI_MARGIN
    End With
    
    With Me.cmdReadConfig
        .Top = Me.txtParseResult.Top + Me.txtParseResult.Height + UI_MARGIN
        .Left = UI_MARGIN
        .Height = Me.txtConfigPath.Height
    End With
    
    With Me.txtConfigPath
        .Left = Me.cmdReadConfig.Left + Me.cmdReadConfig.Width + UI_MARGIN
        .Top = Me.cmdReadConfig.Top
        .Width = Me.ScaleWidth - .Left - UI_MARGIN
    End With
    
    With Me.cmdBuild
        .Top = Me.cmdReadConfig.Top + Me.cmdReadConfig.Height + UI_MARGIN
        .Left = UI_MARGIN
        .Height = Me.txtDataPath.Height
    End With
    
    With Me.txtDataPath
        .Left = Me.cmdBuild.Left + Me.cmdBuild.Width + UI_MARGIN
        .Top = Me.cmdBuild.Top
        .Width = Me.ScaleWidth - .Left - UI_MARGIN
    End With
End Sub

Private Sub cmdReadConfig_Click()
    Dim oConf As New CConfig
    Dim xlsApp As Excel.Application
    Dim xlsWB As Excel.Workbook
    Dim xlsWS As Excel.Worksheet
    Dim srcData As Variant
    
On Error GoTo eh:
    Set xlsApp = GetXLSApp()
    Set xlsWB = xlsApp.Workbooks.Open(Trim$(Me.txtConfigPath.Text), , True)
    Set xlsWS = MExcel.GetExcelSheet(xlsWB, SHEET_CONFIG)
    srcData = MExcel.GetSafeSheetValues(xlsWS, 100, 100)
    oConf.PreviewData srcData
    Set oConf = Nothing
    'oConf.ReadDataConfig "data1"
eh:
    Set xlsWS = Nothing
    If Not xlsWB Is Nothing Then
        xlsWB.Close False
        Set xlsWB = Nothing
    End If
    
    If Err.Number = 0 Then
        MsgBox "ok"
    Else
        MsgBox "read config error"
    End If
End Sub

Private Sub cmdParseExpression_Click()
    Dim oExp As New CExpression
    oExp.Parse Me.txtFormula.Text
    If oExp.errDesc <> "" Then
        Me.txtParseResult.Text = oExp.errDesc
    Else
        Me.txtParseResult.Text = oExp.ToXML()
    End If
End Sub

Private Sub txtConfigPath_DblClick()
    'Me.txtConfigPath.Text = selectFile(Trim$(Me.txtConfigPath.Text))
End Sub

Private Sub txtParseResult_GotFocus()
    Me.txtParseResult.SelStart = 0
    Me.txtParseResult.SelLength = Len(Me.txtParseResult.Text)
    Me.txtParseResult.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        TerminateForm
        End
    End If
End Sub

Private Function GetXLSApp() As Excel.Application
    If m_xlsApp Is Nothing Then Set m_xlsApp = MExcel.GetExcelApp()
    If m_isDebug Then
        m_xlsApp.Visible = True
    Else
        m_xlsApp.Visible = False
    End If
    Set GetXLSApp = m_xlsApp
End Function

Private Sub Form_Initialize()
    m_isDebug = MTDCP.IsDebugApp()
    With Me
        .txtFormula.Text = "mid(F1, int(f2) + int(f3), f4 + f5)"
        '[UPB($)]-   (_F([Bal])/""100""   + _C(""Name"")) + Mid(dd, left(DD), int(Text(XX))) + (((RRR)))
        .txtParseResult = ""
        .txtConfigPath = App.Path & "\sample.xlsx"
        .txtDataPath = App.Path & "\data1.xlsx"
    End With
End Sub

Private Function selectFile(ByVal DefaultPath As String) As String
'    On Error GoTo eh
'    Dim ft As String, fn As String
'    With CommonDialog1
'        .ShowOpen
'        .CancelError = True
'        ft = .FileTitle
'        fn = .FileName
'    End With
'eh:
'    If Len(ft) > 0 Then
'        selectFile = fn
'    Else
'        selectFile = DefaultPath
'    End If
'    If Err.Number = 0 Then Exit Function
'    Err.Clear
End Function


