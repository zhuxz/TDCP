VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main"
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   15615
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
Private Sub cmdParseExpression_Click()
    Dim oExp As New CExpression
    oExp.Parse Me.txtFormula.Text
    If oExp.ErrDesc <> "" Then
        Me.txtParseResult.Text = oExp.ErrDesc
    Else
        Me.txtParseResult.Text = oExp.ToXML()
    End If
End Sub

Private Sub Form_Initialize()
    With Me
        .txtFormula.Text = "[UPB($)]-   (_F([Bal])/""100""   + _C(""Name""))"
        .txtParseResult = ""
    End With
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
End Sub

Private Sub txtParseResult_GotFocus()
    Me.txtParseResult.SelStart = 0
    Me.txtParseResult.SelLength = Len(Me.txtParseResult.Text)
    Me.txtParseResult.SetFocus
End Sub

