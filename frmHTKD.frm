VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmHTKD 
   BackColor       =   &H00C0FFC0&
   Caption         =   "本月开单"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14520
      Picture         =   "frmHTKD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7620
      Width           =   675
   End
   Begin VB.CommandButton cmdRight 
      BackColor       =   &H00C0FFC0&
      Caption         =   ">"
      Height          =   345
      Left            =   14670
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   540
      Width           =   405
   End
   Begin VB.CommandButton cmdLeft 
      BackColor       =   &H00C0FFC0&
      Caption         =   "<"
      Height          =   345
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   540
      Width           =   405
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgA 
      Height          =   6465
      Left            =   30
      TabIndex        =   0
      Top             =   1020
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11404
      _Version        =   393216
      BackColor       =   12648384
      Rows            =   10
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      BackColorUnpopulated=   8454016
      GridColorUnpopulated=   8454016
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin MSComCtl2.DTPicker dtpM 
      Height          =   315
      Left            =   13800
      TabIndex        =   1
      Top             =   90
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   8454016
      CalendarTitleBackColor=   16711808
      CalendarTrailingForeColor=   -2147483635
      CustomFormat    =   "yyyy-MM"
      Format          =   64356355
      CurrentDate     =   38797
   End
   Begin VB.Label lblBM 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   750
      TabIndex        =   6
      Top             =   90
      Width           =   3105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "部门:"
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   150
      Width           =   555
   End
End
Attribute VB_Name = "frmHTKD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub

Private Sub Form_Load()
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Me.Height = 8760
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
dtgA.Cols = 9
dtgA.ColWidth(0) = 4335
dtgA.ColWidth(1) = 2715
dtgA.ColWidth(2) = 1200
dtgA.ColWidth(3) = 1200
dtgA.ColWidth(4) = 1200
dtgA.ColWidth(5) = 1200
dtgA.ColWidth(6) = 1200
dtgA.ColWidth(7) = 1200
dtgA.ColWidth(8) = 0
End Sub

Public Sub Initialize()
lblBm.Caption = ""
dtgA.Clear

dtgA.Rows = 30
dtgA.Row = 0: dtgA.Col = 0: dtgA.Text = "项目名称": dtgA.CellFontBold = True
dtgA.Row = 0: dtgA.Col = 1: dtgA.Text = "合同编号": dtgA.CellFontBold = True
dtgA.Row = 0: dtgA.Col = 2: dtgA.Text = "业务员": dtgA.CellFontBold = True
dtgA.Row = 0: dtgA.Col = 3: dtgA.Text = "合同金额": dtgA.CellFontBold = True
dtgA.Row = 0: dtgA.Col = 4: dtgA.Text = "应收日期": dtgA.CellFontBold = True
dtgA.Row = 0: dtgA.Col = 5: dtgA.Text = "应收金额": dtgA.CellFontBold = True
dtgA.Row = 0: dtgA.Col = 6: dtgA.Text = "开单日期": dtgA.CellFontBold = True
dtgA.Row = 0: dtgA.Col = 7: dtgA.Text = "开单金额": dtgA.CellFontBold = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
frmZu.Enabled = True
Cancel = True
End Sub

Public Sub Bound()
Dim tt As String
Dim oo As Long

End Sub
