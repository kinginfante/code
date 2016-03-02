VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form bView 
   Caption         =   "绩效会议报表"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdL 
      Caption         =   "<-"
      Height          =   315
      Left            =   8010
      TabIndex        =   9
      Top             =   8760
      Width           =   405
   End
   Begin VB.CommandButton cmdR 
      Caption         =   "->"
      Height          =   315
      Left            =   10020
      TabIndex        =   8
      Top             =   8760
      Width           =   405
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14610
      Picture         =   "bView.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8580
      Width           =   585
   End
   Begin VB.CommandButton cmdBm 
      Caption         =   "部门"
      Height          =   585
      Left            =   13110
      Picture         =   "bView.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8580
      Width           =   705
   End
   Begin VB.CommandButton cmdXuan 
      Caption         =   "选 取"
      Height          =   285
      Left            =   30
      TabIndex        =   3
      Top             =   8100
      Width           =   825
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复 制"
      Height          =   285
      Left            =   870
      TabIndex        =   2
      Top             =   8100
      Width           =   825
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "查询"
      Height          =   585
      Left            =   13830
      Picture         =   "bView.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8580
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBB 
      Height          =   8010
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   14129
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker txtM 
      Height          =   345
      Left            =   8400
      TabIndex        =   7
      Top             =   8730
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy年MM月"
      Format          =   54788099
      CurrentDate     =   39415
   End
   Begin VB.Label Label5 
      Caption         =   "月份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7470
      TabIndex        =   10
      Top             =   8790
      Width           =   585
   End
   Begin VB.Label lblFw 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10560
      TabIndex        =   6
      Top             =   8640
      Width           =   2475
   End
End
Attribute VB_Name = "bView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoBview As ADODB.Recordset

Private Sub cmdBack_Click()
Me.Visible = False

b1.Enabled = True
b1.ZOrder 0
End Sub

Private Sub cmdBm_Click()
    Set Ren.XForm = New bView
    Call mod1.RenXz("bView", Me, 0)
End Sub

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText dtgBB.Clip

dtgBB.FixedRows = 0
dtgBB.MergeCol(1) = True
dtgBB.MergeCol(2) = True
dtgBB.MergeCol(5) = True
dtgBB.MergeCells = 3
dtgBB.FixedRows = 1
End Sub

Private Sub cmdL_Click()
txtM.Value = DateSerial(Year(txtM.Value), Month(txtM.Value) - 1, Day(txtM.Value))
End Sub

Private Sub cmdR_Click()
txtM.Value = DateSerial(Year(txtM.Value), Month(txtM.Value) + 1, Day(txtM.Value))
End Sub


Private Sub cmdView_Click()

Dim tt As String
Dim ii As Integer
On Error Resume Next
Set bView.adoBview = New ADODB.Recordset
If lblFw.ToolTipText <> "" Then
    tt = "select 部门,姓名,专项工作内容,完成期限,完成情况 from bview where uid='" & Trim(lblFw.ToolTipText) & "'"
Else
    tt = "select 部门,姓名,专项工作内容,完成期限,完成情况 from bview where 部门='" & Trim(lblFw.Caption) & "'"
End If
Set bView.adoBview = New ADODB.Recordset
bView.adoBview.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If bView.adoBview.RecordCount = 0 Then
    Set bView.dtgBB.DataSource = bView.adoBview
    bView.dtgBB.Rows = 2
    bView.dtgBB.FixedRows = 0
    bView.dtgBB.FixedRows = 1
Else
    bView.dtgBB.FixedRows = 1
    Set bView.dtgBB.DataSource = bView.adoBview
    bView.dtgBB.FixedRows = 0
    bView.dtgBB.MergeCol(1) = True
    bView.dtgBB.MergeCol(2) = True
    bView.dtgBB.MergeCol(5) = True
    bView.dtgBB.MergeCells = 3
    bView.dtgBB.FixedRows = 1
End If


End Sub

Private Sub cmdXuan_Click()
dtgBB.FixedRows = 0
       dtgBB.MergeCells = 0
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
dtgBB.ColWidth(0) = 300
dtgBB.ColWidth(3) = 5500
dtgBB.ColWidth(5) = 5500
End Sub
