VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmOR 
   Caption         =   "您与   的私聊记录"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制"
      Height          =   405
      Left            =   14100
      TabIndex        =   2
      Top             =   8730
      Width           =   555
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14760
      Picture         =   "frmOR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8700
      Width           =   435
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgOld 
      Height          =   8655
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   15266
      _Version        =   393216
      Rows            =   40
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      WordWrap        =   -1  'True
      SelectionMode   =   1
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
If frmOL.Visible = True Then
    frmOL.ZOrder 0
End If
End Sub

Private Sub cmdCopy_Click()
'dtgOld.Col = 0
'dtgOld.Row = 1
'dtgOld.ColSel = 2
'dtgOld.RowSel = dtgOld.Rows - 1
Clipboard.Clear
Clipboard.SetText dtgOld.Clip
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
dtgOld.ColWidth(0) = 2000
dtgOld.ColWidth(2) = 12000

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmOL.Show
If frmOL.Visible = True Then
    frmOL.ZOrder 0
End If
End Sub
