VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGjwV 
   Caption         =   "施工计划查询"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdAll 
      Caption         =   "合部显示"
      Height          =   375
      Left            =   11250
      TabIndex        =   3
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "详     情"
      Height          =   345
      Left            =   11220
      TabIndex        =   2
      Top             =   270
      Width           =   3645
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   555
      Left            =   14520
      Picture         =   "frmGjwV.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8580
      Width           =   645
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8565
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   15108
      _Version        =   393216
      BackColor       =   -2147483634
      BackColorBkg    =   -2147483636
      FillStyle       =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmGjwV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoGJW As ADODB.Recordset

Private Sub cmdAll_Click()
Dim tt As String
On Error Resume Next
If mod1.DName = "张寅" Then
    tt = "select * from GjwV where comid=0 order by gid desc"
'ElseIf mod1.DName = "郑刚" Then
'    tt = "select * from GjwV where comid=0 and qy<>'上海' order by gid desc"
ElseIf mod1.DName = "彭海翔" Then
    tt = "select * from GjwV where comid=1 order by gid desc"
Else '组长
    tt = "select * from GjwV where 组长='" & mod1.DName & "' order by gid desc"
End If
Set frmGjwV.adoGJW = New ADODB.Recordset
frmGjwV.adoGJW.Close
frmGjwV.adoGJW.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGjwV.dtgBr.DataSource = frmGjwV.adoGJW
If frmGjwV.adoGJW.RecordCount > 0 Then
    frmGjwV.dtgBr.FixedRows = 0
    frmGjwV.dtgBr.MergeCol(2) = True
    frmGjwV.dtgBr.MergeCol(3) = True
    frmGjwV.dtgBr.MergeCol(4) = True
    frmGjwV.dtgBr.MergeCells = 3
    frmGjwV.dtgBr.FixedRows = 1
End If
frmGjwV.Visible = True
frmGjwV.Enabled = True
frmGjwV.ZOrder 0
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
frmZu.ZOrder 0
End Sub

Private Sub cmdOpen_Click()
Dim Gid As Long
On Error Resume Next
dtgBr.Col = 5
Gid = dtgBr.Text
Call modGjw.GjwQing
Call modGjw.GjwOpen(Gid)
frmGJW.Show
frmGJW.ZOrder 0
frmGjwV.Enabled = False

End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
Set adoGJW = New ADODB.Recordset
dtgBr.ColWidth(0) = 300
dtgBr.ColWidth(1) = 2500
dtgBr.ColWidth(2) = 4000
dtgBr.ColWidth(4) = 1200
dtgBr.ColWidth(5) = 0
dtgBr.ColWidth(6) = 0
dtgBr.ColWidth(7) = 0
dtgBr.ColWidth(8) = 0
dtgBr.ColWidth(9) = 0
dtgBr.ColWidth(10) = 0
dtgBr.ColWidth(11) = 0
End Sub
