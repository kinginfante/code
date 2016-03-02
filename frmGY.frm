VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGY 
   BackColor       =   &H00C0FFC0&
   Caption         =   "供应商查询"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "frmGY.frx":0000
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.TextBox txtZZ 
      Height          =   300
      Left            =   30
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   8700
      Width           =   5685
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBN 
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   8700
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox txtZ 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   3840
      TabIndex        =   7
      Top             =   8700
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H0080FF80&
      Caption         =   "查询"
      Height          =   315
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8700
      Width           =   1065
   End
   Begin VB.ComboBox comLx 
      BackColor       =   &H00C0FFC0&
      Height          =   300
      ItemData        =   "frmGY.frx":0442
      Left            =   1440
      List            =   "frmGY.frx":0452
      TabIndex        =   4
      Top             =   8700
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.CommandButton cmdQX 
      BackColor       =   &H00FFFFC0&
      Caption         =   "权限"
      Height          =   555
      Left            =   13830
      Picture         =   "frmGY.frx":0478
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8580
      Width           =   675
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0FFFF&
      Caption         =   "返回"
      Height          =   555
      Left            =   14550
      Picture         =   "frmGY.frx":08BA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8580
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgB 
      Height          =   8325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   14684
      _Version        =   393216
      BackColor       =   12648384
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "值"
      Height          =   225
      Left            =   3570
      TabIndex        =   5
      Top             =   8760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "查询类别"
      Height          =   255
      Left            =   510
      TabIndex        =   3
      Top             =   8760
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmGY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True

End Sub

Private Sub cmdC_Click()
Dim tt As String
Dim Ra
Dim La
Dim oo As Long
On Error Resume Next
''''''Select Case Trim(comLx.Text)

'''''Case "客户名称"
    tt = "select yz,mc,dj,fw,gid from gymxc where mc like '%" & txtZZ.Text & "%'"
'''''Case "联系人"
'''''    tt = "select yz,mc,lxr,jfw,gid from gymxc where lxr like '%" & txtzz.Text & "%'"
'''''Case "经营范围"
'''''    tt = "select yz,mc,lxr,jfw,gid from gymxc where jfw like '%" & txtzz.Text & "%'"
'''''Case "全部"
'''''    tt = "select yz,mc,lxr,jfw,gid from gymxc"
'''''End Select
tt = tt & " order by gid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2) + 2
    dtgB.Rows = La: dtgBN.Rows = La
    dtgBN.Cols = dtgB.Cols
    dtgB.Visible = False
    Call Me.dtgBFF
    For oo = 1 To La
        dtgB.Row = oo: dtgBN.Row = oo
        dtgB.Col = 0: dtgBN.Col = 0
        dtgB.Text = Ra(0, oo - 1): dtgBN.Text = Ra(0, oo - 1)
        dtgB.Col = 1: dtgBN.Col = 1
        dtgB.Text = Ra(1, oo - 1): dtgBN.Text = Ra(1, oo - 1)
        dtgB.Col = 2: dtgBN.Col = 2
        dtgB.Text = Ra(2, oo - 1): dtgBN.Text = Ra(2, oo - 1)
        dtgB.Col = 3: dtgBN.Col = 3
        dtgB.Text = Ra(3, oo - 1): dtgBN.Text = Ra(3, oo - 1)
        dtgB.Col = 4: dtgBN.Col = 4
        dtgB.Text = Ra(4, oo - 1): dtgBN.Text = Ra(4, oo - 1)
    Next
    dtgB.Visible = True
End Sub

Private Sub cmdQX_Click()
frmGyQX.Show
frmGyQX.ZOrder 0
End Sub

Private Sub comLx_Click()
Dim oo As Long
Dim tt As String
Dim Ra
Dim La
If Me.Visible = False Then Exit Sub
On Error Resume Next
For oo = 1000 To 0 Step -1
    txtZ.RemoveItem oo
Next
Select Case comLx.Text

Case "客户名称"
Case "联系人"

Case "全部"
    txtZ.Text = ""
End Select
End Sub

Private Sub dtgB_DblClick()
On Error Resume Next
Dim Gid As Long
dtgBN.Row = dtgB.Row
dtgBN.Col = 4
Gid = Val(dtgBN.Text)
'If Gid = 0 Then Exit Sub
Call frmGyDetail.Qing
Call frmGyDetail.Bound(Gid)
frmGyDetail.cmdSave.Enabled = False
frmGyDetail.Show
frmGyDetail.ZOrder 0
End Sub


Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
End Sub

Public Sub dtgBFF()
dtgB.Cols = 5
dtgB.Clear
dtgB.ColWidth(0) = 2745
dtgB.ColWidth(1) = 3000
dtgB.ColWidth(2) = 1440
dtgB.ColWidth(3) = 7635
dtgB.ColWidth(4) = 0



On Error Resume Next

dtgB.Row = 0
dtgB.Col = 0
dtgB.Text = "类别": dtgB.CellFontBold = True
dtgB.Col = 1
dtgB.Text = "客户名称及部门": dtgB.CellFontBold = True
dtgB.Col = 2
dtgB.Text = "联系人": dtgB.CellFontBold = True
dtgB.Col = 3
dtgB.Text = "经营范围": dtgB.CellFontBold = True




End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
frmZu.Enabled = True
frmZu.ZOrder 0
End Sub
