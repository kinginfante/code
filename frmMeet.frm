VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmMeet 
   BackColor       =   &H00C0FFC0&
   Caption         =   "会议记录查询"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   555
      Left            =   8400
      TabIndex        =   7
      Top             =   8520
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   979
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdGB 
      Caption         =   "关闭"
      Height          =   315
      Left            =   13980
      TabIndex        =   5
      Top             =   8670
      Width           =   735
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "查询"
      Height          =   285
      Left            =   5550
      TabIndex        =   4
      Top             =   8640
      Width           =   945
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmMeet.frx":0000
      Left            =   3150
      List            =   "frmMeet.frx":0013
      TabIndex        =   3
      Text            =   "关键字"
      Top             =   8610
      Width           =   1095
   End
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   4260
      TabIndex        =   2
      Top             =   8610
      Width           =   1185
   End
   Begin VB.CommandButton cmdMe 
      Caption         =   "我的会议"
      Height          =   375
      Left            =   330
      TabIndex        =   1
      Top             =   8580
      Width           =   1665
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMeet 
      Height          =   8355
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   14737
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "查询方式"
      Height          =   255
      Left            =   2340
      TabIndex        =   6
      Top             =   8640
      Width           =   735
   End
End
Attribute VB_Name = "frmMeet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGB_Click()
Me.Visible = False
End Sub

Private Sub Command1_Click()



End Sub

Private Sub cmdMe_Click()
Dim tt As String
Dim Ra
tt = "select rq,lx,mc,nlb,mid from meetView where zren='" & mod1.DName & "' group by mid,rq,lx,mc,nlb order by rq desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Call Me.dtgBound(Ra)
End Sub

Private Sub dtgMeet_DblClick()
dtgN.Row = dtgMeet.Row
dtgN.Col = 4
If Val(dtgN.Text) = 0 Then Exit Sub
Call frmMeetDetail.Bound(Val(dtgN.Text))
frmMeetDetail.Show
End Sub

Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
'Call dtgMeetFF
Call cmdMe_Click
End Sub

Public Sub dtgMeetFF()
dtgMeet.Clear
dtgMeet.Cols = 5
dtgMeet.Rows = 50
dtgMeet.Row = 0
dtgMeet.Col = 0: dtgMeet.Text = "日期": dtgMeet.CellFontBold = True
dtgMeet.Col = 1: dtgMeet.Text = "会议性质": dtgMeet.CellFontBold = True
dtgMeet.Col = 2: dtgMeet.Text = "会议名称": dtgMeet.CellFontBold = True
dtgMeet.Col = 3: dtgMeet.Text = "会议摘要": dtgMeet.CellFontBold = True
dtgMeet.ColWidth(0) = 1500
dtgMeet.ColWidth(1) = 1500
dtgMeet.ColWidth(2) = 3000
dtgMeet.ColWidth(3) = 8670
dtgMeet.ColWidth(4) = 0

dtgN.Clear
dtgN.Cols = 5
dtgN.Rows = 50
dtgN.Row = 0
dtgN.Col = 0: dtgN.Text = "日期": dtgN.CellFontBold = True
dtgN.Col = 1: dtgN.Text = "会议性质": dtgN.CellFontBold = True
dtgN.Col = 2: dtgN.Text = "会议名称": dtgN.CellFontBold = True
dtgN.Col = 3: dtgN.Text = "会议摘要": dtgN.CellFontBold = True
dtgN.ColWidth(0) = 1500
dtgN.ColWidth(1) = 1500
dtgN.ColWidth(2) = 3000
dtgN.ColWidth(3) = 8670
dtgN.ColWidth(4) = 0

End Sub

Public Sub dtgBound(Ra)
Dim La As Long
Call dtgMeetFF
La = UBound(Ra, 2) + 1
For oo = 1 To La
    dtgMeet.Row = oo
    dtgMeet.Col = 0: dtgMeet.Text = Ra(0, oo - 1)
    dtgMeet.Col = 1: dtgMeet.Text = Ra(1, oo - 1)
    dtgMeet.Col = 2: dtgMeet.Text = Ra(2, oo - 1)
    dtgMeet.Col = 3: dtgMeet.Text = Ra(3, oo - 1)
    dtgMeet.Col = 4: dtgMeet.Text = Ra(4, oo - 1)
    
    dtgN.Row = oo
    dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1)
    dtgN.Col = 1: dtgN.Text = Ra(1, oo - 1)
    dtgN.Col = 2: dtgN.Text = Ra(2, oo - 1)
    dtgN.Col = 3: dtgN.Text = Ra(3, oo - 1)
    dtgN.Col = 4: dtgN.Text = Ra(4, oo - 1)
Next



End Sub
