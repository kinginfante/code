VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcCGY 
   BackColor       =   &H00C0FFC0&
   Caption         =   "待采购货品"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form2"
   ScaleHeight     =   5895
   ScaleWidth      =   9990
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   585
      Left            =   9300
      Picture         =   "FmxcCGY.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5310
      Width           =   585
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9234
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
End
Attribute VB_Name = "FmxcCGY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub dtgBRFF()
dtgBr.Clear
dtgBr.Rows = 100
dtgBr.Cols = 5
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "执行通知日期": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "编号": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "货品": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "数量": dtgBr.CellFontBold = True

End Sub

