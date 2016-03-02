VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form KCBB 
   Caption         =   "库存报表"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   555
      Left            =   14610
      Picture         =   "KCBB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8640
      Width           =   645
   End
   Begin VB.CommandButton cmdREF 
      Caption         =   "刷   新"
      Height          =   285
      Left            =   13080
      TabIndex        =   3
      Top             =   1290
      Width           =   2115
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgKcbb 
      Height          =   9105
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12945
      _ExtentX        =   22834
      _ExtentY        =   16060
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   13080
      TabIndex        =   1
      Top             =   600
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   503
      _Version        =   393216
      Format          =   67895297
      CurrentDate     =   38915
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   13080
      TabIndex        =   0
      Top             =   240
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      Format          =   67895297
      CurrentDate     =   38915
   End
End
Attribute VB_Name = "KCBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoKcbb As ADODB.Recordset

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
frmZu.ZOrder 0

End Sub

Private Sub cmdREF_Click()
Dim tt As String
On Error Resume Next
tt = "select * from kcbb where comid=" & mod1.comId & " order by 品牌,编号,货品名称"
adoKcbb.Close
adoKcbb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgKcbb.DataSource = adoKcbb

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Set adoKcbb = New ADODB.Recordset
dtgKcbb.ColWidth(0) = 300
dtgKcbb.ColWidth(1) = 1500
dtgKcbb.ColWidth(2) = 1500
dtgKcbb.ColWidth(3) = 2000
dtgKcbb.ColWidth(6) = 0
dtgKcbb.ColWidth(7) = 3000
dtgKcbb.ColWidth(9) = 0
End Sub
