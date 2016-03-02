VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form frmGxBNew 
   BackColor       =   &H00C0FFC0&
   Caption         =   "未处理询价"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   765
      Left            =   14430
      Picture         =   "frmGxBNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "返回"
      Top             =   8280
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   2325
      Left            =   13470
      TabIndex        =   2
      Top             =   4110
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   4101
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin NiceFormControl.NiceButton NiceButton1 
      Height          =   315
      Left            =   12930
      TabIndex        =   1
      Top             =   240
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   556
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmGxBNew.frx":0102
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "详    情"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   9045
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   15954
      _Version        =   393216
      BackColor       =   12648384
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin NiceFormControl.NiceButton NiceButton2 
      Height          =   315
      Left            =   12930
      TabIndex        =   4
      Top             =   750
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   556
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmGxBNew.frx":011E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "刷    新"
   End
End
Attribute VB_Name = "frmGxBNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bid As Long
Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub dtgBr_Click()
dtgN.Row = dtgBr.Row
dtgN.Col = 5
Bid = Val(dtgN.Text)
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
End Sub

Public Sub dtgbrFF()
dtgBr.Clear
dtgBr.Cols = 6
dtgBr.Rows = 100
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "项目名称": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "类型": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "询价日期": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "业务员": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "编号": dtgBr.CellFontBold = True
dtgBr.Col = 5: dtgBr.Text = "Bid": dtgBr.CellFontBold = True
dtgBr.ColWidth(5) = 0
dtgBr.ColWidth(0) = 6195
dtgBr.ColWidth(1) = 2025
dtgBr.ColWidth(2) = 2055
dtgN.Clear
dtgN.Cols = 6
dtgN.Rows = 100

End Sub

Public Sub Bound()
Dim tt As String
Dim Ra
Dim La As Long
Dim oo As Long
Call dtgbrFF
dtgBr.Visible = False
tt = "select xmmc,zl,rq,ywy,'XJD'+cast(bid as nvarchar(20)),bid from xunjiaD where lcren='' and delf=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
For oo = 1 To La
    dtgBr.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1)
    dtgBr.Col = 1: dtgBr.Text = Ra(1, oo - 1)
    dtgBr.Col = 2: dtgBr.Text = Ra(2, oo - 1)
    dtgBr.Col = 3: dtgBr.Text = Ra(3, oo - 1)
    dtgBr.Col = 4: dtgBr.Text = Ra(4, oo - 1)
    dtgBr.Col = 5: dtgBr.Text = Ra(5, oo - 1)
    
    dtgN.Row = oo
    dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1)
    dtgN.Col = 1: dtgN.Text = Ra(1, oo - 1)
    dtgN.Col = 2: dtgN.Text = Ra(2, oo - 1)
    dtgN.Col = 3: dtgN.Text = Ra(3, oo - 1)
    dtgN.Col = 4: dtgN.Text = Ra(4, oo - 1)
    dtgN.Col = 5: dtgN.Text = Ra(5, oo - 1)
Next
dtgBr.Visible = True
End Sub

Private Sub NiceButton1_Click()
If Bid = 0 Then Exit Sub
    mod1.BTZ = 36
    Call FmxcXJ.Bound(Bid)
''''''    dtgLx.Col = 3: Call FmxcXJ.SDJE(Val(dtgLx.Text))
    FmxcXJ.Show
    FmxcXJ.ZOrder 0
End Sub


Private Sub NiceButton2_Click()
Call Me.Bound
End Sub
