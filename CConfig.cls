VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lastErr As String

Private m_srcData As Variant
Private m_srcDataProp As ArrayProp

Private m_allHeads() As Variant
Private m_headCount As Variant

Private m_msrs() As Variant
Private m_msrCount As Long

Private m_configs As Collection

Private Sub Class_Initialize()
    m_headCount = 0
    m_msrCount = 0
    Set m_configs = New Collection
End Sub

Private Sub Class_Terminate()
    Me.Clear
    Set m_configs = Nothing
End Sub

Public Function Clear()
    Dim data_config As CDataConfig
    With m_configs
        While .Count > 0
            Set data_config = .item(1)
            Set data_config = Nothing
            .Remove 1
        Wend
    End With
End Function

Private Function AppendCommonTHCHead(ByVal rowId As Long) As Long
    Dim head(1) As String
    Dim startCol As Long: startCol = 1
    head(0) = Trim$(m_srcData(rowId)(startCol))
    head(1) = Trim$(m_srcData(rowId)(startCol + 1))
    
    If m_headCount = 0 Then
        ReDim m_allHeads(m_headCount) As Variant
    Else
        ReDim Preserve m_allHeads(m_headCount) As Variant
    End If
    
    m_allHeads(m_headCount) = head
    AppendCommonTHCHead = m_headCount
    
    m_headCount = m_headCount + 1
End Function

Private Function AppendMSRConfig(ByVal rowId As Long) As Long
    Dim dataName As String
    Dim startCol As Long: startCol = 1
    dataName = Trim$(m_srcData(rowId)(startCol))
    
    If m_msrCount = 0 Then
        ReDim m_msrs(m_msrCount) As Variant
    Else
        ReDim Preserve m_msrs(m_msrCount) As Variant
    End If
    
    m_msrs(m_msrCount) = dataName
    AppendMSRConfig = m_msrCount
    
    m_msrCount = m_msrCount + 1
End Function

Private Function AppendDataConfig(rowId, Optional beginCol As Long = 1) As Long
    ''' to do
End Function

Public Function PreviewData(SourceData, Optional ByRef errInfo As String) As Variant
On Error GoTo eh
    m_srcData = SourceData
    If Not MFunc.CheckArray(SourceData, m_srcDataProp) Then
        errInfo = "config data is not an array"
        Exit Function
    End If
    
    Dim UKW As UCaseKeyWords: UKW = MConfig.GetUCaseKeyWords()
    Dim beginCol As Long: beginCol = LBound(SourceData(m_srcDataProp.ub))
    
    Dim section As TConfig
    Dim iRow As Long
    Dim str As String
    Dim oDataConfig As CDataConfig
    
    For iRow = m_srcDataProp.lb To m_srcDataProp.ub
        str = UCase(Trim(SourceData(iRow)(beginCol)))
        Select Case str
            Case UKW.FieldsMap
                section = TConfig.allTHeads
            Case UKW.RelatedDataSheet
                section = TConfig.ConfigMain
                Set oDataConfig = New CDataConfig
                oDataConfig.m_srcName = Trim$(SourceData(iRow)(beginCol + 1))
                oDataConfig.m_srcStart = MFunc.Var2Long(SourceData(iRow)(beginCol + 2), 1)
                oDataConfig.m_destName = Trim$(SourceData(iRow)(beginCol + 3))
                oDataConfig.m_destStart = MFunc.Var2Long(SourceData(iRow)(beginCol + 4), 1)
                oDataConfig.m_isOptional = IIf(UCase(SourceData(iRow)(beginCol + 5)) = UCase(KW_Optional), True, False)
                oDataConfig.m_srcBeginRow = iRow
                m_configs.Add oDataConfig
            Case UKW.MSRDataSheet
                section = TConfig.MSRDataSheet
                AppendMSRConfig iRow
            Case Else
                If section = TConfig.allTHeads Then
                    AppendCommonTHCHead iRow
                ElseIf section = TConfig.ConfigMain Then
                    oDataConfig.m_srcEndRow = iRow
                End If
        End Select
    Next
Exit Function
eh:
    errInfo = Err.Description
    m_lastErr = errInfo
    Err.Clear
End Function

Private Function GetConfigId(ByVal srcName As String) As Long
    GetConfigId = -1
    Dim iConf As Long
    For iConf = 1 To m_configs.Count
        If StrComp(srcName, m_configs.item(iConf).m_srcName, vbTextCompare) = 0 Then
            GetConfigId = iConf
            Exit For
        End If
    Next
End Function

Public Function GetDataConfig(ByVal srcName As String, Optional ByRef errInfo As String) As CDataConfig
On Error GoTo eh
    Dim iConf As Long: iConf = GetConfigId(srcName)
    If iConf < 0 Then
        errInfo = "cannot find config"
        Exit Function
    End If
    
    Dim iRow As Long
    Dim oDataConf As CDataConfig: Set oDataConf = m_configs.item(iConf)
    With oDataConf
        Dim data(): ReDim data(.m_srcBeginRow + 1 To .m_srcEndRow)
        For iRow = .m_srcBeginRow + 1 To .m_srcEndRow
            data(iRow) = m_srcData(iRow)
        Next
        oDataConf.AcceptData data
    End With
    
    Set GetDataConfig = oDataConf
Exit Function
eh:
    errInfo = Err.Description
    m_lastErr = errInfo
    Err.Clear
End Function

Public Property Get THCHeads() As Variant
    If m_headCount > 0 Then THCHeads = m_allHeads
End Property

Public Property Get Result() As Collection
    Set Result = m_configs
End Property

Public Property Get MSRs() As Variant
    If m_msrCount > 0 Then MSRs = m_msrs
End Property


