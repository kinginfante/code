VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmComJian 
   BackColor       =   &H00C0FFC0&
   Caption         =   "IT设备状态表"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15180
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15180
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   465
      Left            =   3750
      TabIndex        =   6
      Top             =   300
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   820
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   585
      Left            =   14445
      Picture         =   "frmComJian.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8490
      Width           =   675
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制"
      Height          =   585
      Left            =   13740
      Picture         =   "frmComJian.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "点击后,打开EXCEL,可将表格内容粘贴."
      Top             =   8490
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMa 
      Height          =   7725
      Left            =   8490
      TabIndex        =   2
      Top             =   660
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   13626
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   8
      Cols            =   3
      BackColorFixed  =   15728356
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgZong 
      Height          =   7725
      Left            =   0
      TabIndex        =   3
      Top             =   660
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   13626
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   8
      Cols            =   3
      BackColorFixed  =   16761024
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "领用明细"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   8610
      TabIndex        =   5
      Top             =   210
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IT设备一览"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00305AED&
      Height          =   345
      Left            =   330
      TabIndex        =   4
      Top             =   210
      Width           =   1875
   End
End
Attribute VB_Name = "frmComJian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Initialize()
Call Me.ZongInitialize
Call Me.MaInitialize
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub dtgZong_Click()
Dim tt As String
Dim Rb
Dim Lb As Integer
Dim Cid As Long
On Error Resume Next
dtgN.Col = 0: dtgN.Row = dtgZong.Row
Cid = Val(dtgN.Text)
tt = "select aid,atime,operating,ywy,qm from fadetail where cid=" & Cid & " order by aid desc"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lb = UBound(Rb, 2) + 1
Call Me.MaInitialize
Call MaBound(Rb, Lb)
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
End Sub

Public Sub Bound()
Dim tt As String
Dim Ra, Rb
Dim La As Integer
Dim Lb As Integer
Call Me.Initialize
tt = "select cid,lb,cname,cbh,status,bz from FixedAssets order by lb,cid;" & _
    "select aid,atime,operating,ywy,qm from fadetail order by aid desc"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
La = UBound(Ra, 2) + 1
Lb = UBound(Rb, 2) + 1
mod1.HTP.Close
Set mod1.HTP = Nothing
Call Me.ZongBound(Ra, La)
Call Me.MaBound(Rb, Lb)
End Sub

Public Sub ZongInitialize()
dtgZong.Clear
dtgZong.Cols = 6
dtgZong.Row = 0
dtgZong.Col = 1: dtgZong.Text = "物品名称": dtgZong.CellFontBold = True
dtgZong.Col = 2: dtgZong.Text = "物品型号": dtgZong.CellFontBold = True
dtgZong.Col = 3: dtgZong.Text = "编号": dtgZong.CellFontBold = True
dtgZong.Col = 4: dtgZong.Text = "状态(使用管理者)": dtgZong.CellFontBold = True
dtgZong.Col = 5: dtgZong.Text = "备注": dtgZong.CellFontBold = True
dtgZong.ColWidth(0) = 0: dtgZong.ColWidth(1) = 1785
dtgZong.ColWidth(2) = 1515: dtgZong.ColWidth(3) = 1215
dtgZong.ColWidth(4) = 1770
dtgZong.ColWidth(5) = 1725
dtgZong.RowHeight(0) = 405

dtgN.Clear
dtgN.Cols = 6



End Sub

Public Sub MaInitialize()
dtgMa.Clear
dtgMa.Cols = 5
dtgMa.Row = 0
dtgMa.Col = 1: dtgMa.Text = "执行日期": dtgMa.CellFontBold = True
dtgMa.Col = 2: dtgMa.Text = "操作": dtgMa.CellFontBold = True
dtgMa.Col = 3: dtgMa.Text = "姓名": dtgMa.CellFontBold = True
dtgMa.Col = 4: dtgMa.Text = "签名备注": dtgMa.CellFontBold = True
dtgMa.ColWidth(0) = 0
dtgMa.ColWidth(1) = 1875
dtgMa.ColWidth(2) = 1020
dtgMa.RowHeight(0) = 405
dtgMa.ColWidth(3) = 1350
dtgMa.ColWidth(4) = 2010


End Sub

Public Sub ZongBound(Ra, La As Integer)
Dim oo As Long
On Error Resume Next
dtgZong.Visible = False
dtgZong.Rows = La + 20: dtgN.Rows = dtgZong.Rows
For oo = 1 To La
    dtgZong.Row = oo: dtgN.Row = oo
    dtgN.Col = 0
    dtgZong.Col = 0: dtgZong.Text = Ra(0, oo - 1): dtgN.Text = Ra(0, oo - 1)
    dtgZong.Col = 1: dtgZong.Text = Ra(1, oo - 1)
    dtgZong.Col = 2: dtgZong.Text = Ra(2, oo - 1): dtgZong.CellAlignment = 1
    dtgZong.Col = 3: dtgZong.Text = Ra(3, oo - 1)
    dtgZong.Col = 4: dtgZong.Text = Ra(4, oo - 1)
    dtgZong.Col = 5: dtgZong.Text = Ra(5, oo - 1)
Next
dtgZong.Visible = True
End Sub

Public Sub MaBound(Rb, Lb As Integer)
Dim oo As Long
On Error Resume Next
dtgMa.Visible = False
If Lb = 0 Then
    dtgMa.Visible = True
    Exit Sub
End If
dtgMa.Rows = Lb + 30
For oo = 1 To Lb
    dtgMa.Row = oo
    dtgMa.Col = 0: dtgMa.Text = Rb(0, oo - 1) 'Aid
    dtgMa.Col = 1: dtgMa.Text = Rb(1, oo - 1): dtgMa.CellAlignment = 1
    dtgMa.Col = 2: dtgMa.Text = Rb(2, oo - 1): dtgMa.CellAlignment = 1
    dtgMa.Col = 3: dtgMa.Text = Rb(3, oo - 1) '
    dtgMa.Col = 4: dtgMa.Text = Rb(4, oo - 1) '

Next
dtgMa.Visible = True
End Sub
