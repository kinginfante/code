VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmCWBBA 
   Caption         =   "������ϸ��"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   11190
   Begin VB.CommandButton cmdCopy 
      Caption         =   "����"
      Height          =   525
      Left            =   9420
      TabIndex        =   2
      Top             =   5400
      Width           =   705
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   555
      Left            =   10530
      Picture         =   "frmCWBBA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5370
      Width           =   645
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgL 
      Height          =   5325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   9393
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmCWBBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoL As Object
Public BCol As Integer '�������ӱ�ŵ���ID
Private Sub cmdBack_Click()
Me.Visible = False
frmCWBB.Enabled = True
frmCWBB.ZOrder 0
End Sub

Private Sub cmdCopy_Click()
dtgL.FixedRows = 0
dtgL.FixedCols = 0
dtgL.Col = 0
dtgL.Row = 0
''''''If comLx.Text = "�Ŷӷ���" Then
''''''    dtgL.ColSel = 11
''''''ElseIf comLx.Text = "���˷���" Then
''''''    dtgL.ColSel = 23
''''''
''''''ElseIf comLx.Text = "���˸��� ���" Then
    dtgL.ColSel = 10
'''''ElseIf comLx.Text = "��˾������ϸ" Then
'''''    dtgL.ColSel = 13
'''''ElseIf comLx.Text = "Ӧ���ʿ�" Then
'''''    dtgL.ColSel = 6
'''''
'''''End If
    dtgL.RowSel = dtgL.Rows - 3
Clipboard.Clear
Clipboard.SetText dtgL.Clip
dtgL.FixedRows = 1
'''''If comLx.Text = "��˾������ϸ" Then
'''''    dtgBB.FixedCols = 1
'''''ElseIf comLx.Text = "Ӧ���ʿ�" Then
'''''    dtgBB.FixedCols = 0
'''''Else
'''''    dtgBB.FixedCols = 2
'''''End If
'''''    dtgBB.MergeCol(0) = True
'''''    dtgBB.MergeCells = 3
End Sub

Private Sub dtgL_DblClick()
Dim Bxid As String
dtgL.Col = BCol
Bxid = Trim(dtgL.Text)
If Bxid = "" Then Exit Sub
frmFYBX.Show
frmFYBX.ZOrder 0
Call ModBx.FyQing
Call ModBx.fydBound(Bxid)
Me.Enabled = False
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = 6465
Me.Width = 11310
dtgL.Cols = 3
dtgL.Rows = 100
dtgL.FixedCols = 0
dtgL.Row = 0
dtgL.Col = 0
dtgL.Text = "ǩ������"
dtgL.Col = 1
dtgL.Text = "������"
dtgL.Col = 2
dtgL.Text = "���������"
dtgL.ColWidth(1) = 5190
End Sub
