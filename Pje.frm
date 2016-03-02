VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Pje 
   Caption         =   "评审建议"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9675
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   9675
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgA 
      Height          =   585
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1032
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgPje 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   9
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.TextBox txtXQ 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4830
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   "详情:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   4890
      Width           =   705
   End
End
Attribute VB_Name = "Pje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoPje As Object

Private Sub dtgPje_Click()
On Error Resume Next
dtgA.Row = dtgPje.Row
dtgA.Col = 4
txtXQ.Text = dtgA.Text
End Sub

'''Private Sub dtgPje_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
''''''On Error Resume Next
''''''txtXQ.Text = adoPje.Fields("bz").Value
'''End Sub

Private Sub Form_Load()
Me.Height = 6615
Me.Width = 9795
Set adoPje = CreateObject("adodb.recordset")
dtgPje.ColWidth(0) = 300
dtgPje.ColWidth(1) = 2000
dtgPje.ColWidth(4) = 3500
End Sub
