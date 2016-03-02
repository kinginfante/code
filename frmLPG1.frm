VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmLPG 
   BackColor       =   &H00C0FFC0&
   Caption         =   "杰升零件系统"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdBasic 
      BackColor       =   &H00C0FFC0&
      Caption         =   "机组基础数据"
      Height          =   435
      Left            =   12810
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4590
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   585
      Left            =   12420
      TabIndex        =   12
      Top             =   8640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1032
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14550
      Picture         =   "frmLPG1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8580
      Width           =   675
   End
   Begin VB.OptionButton optBrand 
      BackColor       =   &H00C00000&
      Caption         =   "麦克维尔"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   4
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3540
      Width           =   2115
   End
   Begin VB.OptionButton optBrand 
      BackColor       =   &H00C00000&
      Caption         =   "顿汉布什"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   3
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2790
      Width           =   2115
   End
   Begin VB.OptionButton optBrand 
      BackColor       =   &H00C00000&
      Caption         =   "特  灵"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   2
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   2115
   End
   Begin VB.OptionButton optBrand 
      BackColor       =   &H00C00000&
      Caption         =   "开  利"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   1
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1290
      Width           =   2115
   End
   Begin VB.OptionButton optBrand 
      BackColor       =   &H00C00000&
      Caption         =   "约  克"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   0
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   540
      Width           =   2115
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查询"
      Height          =   315
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8700
      Width           =   1185
   End
   Begin VB.ComboBox txtZ 
      Height          =   300
      Left            =   4620
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   8670
      Width           =   2715
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   8640
      Width           =   2685
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8355
      Left            =   -30
      TabIndex        =   5
      Top             =   0
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   14737
      _Version        =   393216
      BackColor       =   12648384
      Rows            =   10
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      BackColorUnpopulated=   8454016
      GridColorUnpopulated=   8454016
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "值"
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Top             =   8700
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "查询方式:"
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   8730
      Width           =   1005
   End
End
Attribute VB_Name = "frmLPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub

Private Sub dtgBr_DblClick()
Dim Hid As Long
dtgN.Row = dtgBr.Row
If dtgN.Row = 0 Then Exit Sub
dtgN.Col = 5
Hid = Val(dtgN.Text)
Call frmLPGDetail.Initialize
Call frmLPGDetail.BoundHM(Hid)
frmLPGDetail.Show: frmLPGDetail.ZOrder 0
frmLPGDetail.frmRealNumbers.Visible = False
frmLPGDetail.frmSupplier.Visible = False
End Sub


Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
Call Me.Initialize
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub


Private Sub optBrand_Click(Index As Integer)
Dim tt As String
Dim Ra
Dim La As Long
Select Case Index
Case 0
    tt = "select HMNumbers,originallyNumbers,partName,replaceNumber1,replaceNumber2,hid from Nlpg"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2) + 1
    Call Me.REF(Ra, La)
End Select
End Sub

Public Sub REF(Ra, La As Long)
Dim oo As Long
dtgBr.Visible = False
Call Initialize
dtgBr.Rows = La + 1: dtgN.Rows = La + 1
For oo = 1 To La
    dtgBr.Row = oo: dtgN.Row = oo
    dtgBr.Col = 0: dtgN.Col = 0
    dtgBr.Text = Ra(0, oo - 1): dtgN.Text = Ra(0, oo - 1)
    dtgBr.Col = 1: dtgN.Col = 1
    dtgBr.Text = Ra(1, oo - 1): dtgN.Text = Ra(1, oo - 1)
    dtgBr.Col = 2: dtgN.Col = 2
    dtgBr.Text = Ra(2, oo - 1): dtgN.Text = Ra(2, oo - 1)
    dtgBr.Col = 3: dtgN.Col = 3
    dtgBr.Text = Ra(3, oo - 1): dtgN.Text = Ra(3, oo - 1)
    dtgBr.Col = 4: dtgN.Col = 4
    dtgBr.Text = Ra(4, oo - 1): dtgN.Text = Ra(4, oo - 1)
    dtgBr.Col = 5: dtgN.Col = 5
    dtgBr.Text = Ra(5, oo - 1): dtgN.Text = Ra(5, oo - 1)
Next
dtgBr.Visible = True
End Sub

Public Sub Initialize()
dtgBr.Clear
dtgN.Clear
dtgN.Cols = dtgBr.Cols
dtgBr.ColWidth(0) = 2235
dtgBr.ColWidth(1) = 1515
dtgBr.ColWidth(2) = 2220
dtgBr.ColWidth(3) = 1875
dtgBr.ColWidth(4) = 2205
dtgBr.ColWidth(5) = 0
dtgBr.Row = 0: dtgBr.Col = 0: dtgBr.Text = "豪曼零配件编号": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "原厂编号": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "零件名字": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "渠道替代编号": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "功能替代编号": dtgBr.CellFontBold = True
End Sub
