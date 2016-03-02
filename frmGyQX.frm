VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGyQX 
   BackColor       =   &H00C0FFFF&
   Caption         =   "供应商权限设置"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   10470
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "权限设置"
      Height          =   7695
      Left            =   -30
      TabIndex        =   4
      Top             =   0
      Width           =   5805
      Begin VB.ComboBox comRen 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   990
         TabIndex        =   5
         Top             =   330
         Width           =   3165
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgQX 
         Height          =   6435
         Left            =   960
         TabIndex        =   6
         Top             =   810
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   11351
         _Version        =   393216
         BackColor       =   12648447
         FixedCols       =   0
         BackColorFixed  =   16777152
         BackColorBkg    =   12648447
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "员工"
         Height          =   345
         Left            =   270
         TabIndex        =   8
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "经营范围一览"
         Height          =   495
         Left            =   210
         TabIndex        =   7
         Top             =   990
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BFFFE2&
      Caption         =   "查询日志"
      Height          =   7695
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   4725
      Begin VB.CommandButton cmdback 
         BackColor       =   &H00C0FFFF&
         Caption         =   "返回"
         Height          =   645
         Left            =   4050
         Picture         =   "frmGyQX.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7020
         Width           =   645
      End
      Begin VB.ComboBox comR 
         BackColor       =   &H00BFFFE2&
         Height          =   300
         Left            =   1020
         TabIndex        =   2
         Top             =   300
         Width           =   2985
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRZ 
         Height          =   6135
         Left            =   0
         TabIndex        =   1
         Top             =   810
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   10821
         _Version        =   393216
         BackColor       =   15728356
         FixedCols       =   0
         BackColorFixed  =   16777152
         BackColorBkg    =   12582882
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "员工"
         Height          =   255
         Left            =   390
         TabIndex        =   3
         Top             =   390
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmGyQX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()
Dim tt As String
Dim oo As Long
Dim Ra, Rb
Dim La, Lb
Me.Width = 10590
Me.Height = 8175
Call dtgQxFF
Call dtgRZFF





tt = "select username from worker where bm='零件事业部' and zzf=1;" & _
   "select jfw,max(gid) from gongying where jfw<>'' group by jfw "
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Lb = UBound(Rb, 2) + 1
For oo = 0 To La - 1
    comRen.AddItem Ra(0, oo), oo
    comR.AddItem Ra(0, oo), oo
Next
comR.AddItem "全部"
comR.Text = "全部"

On Error Resume Next
dtgQX.Rows = Lb + 1
'经营范围
For oo = 1 To Lb + 1
    dtgQX.Row = oo
    dtgQX.Col = 0: dtgQX.Text = Rb(0, oo)
    dtgQX.Col = 2: dtgQX.Text = Rb(1, oo)
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
End Sub

Public Sub dtgQxFF()
dtgQX.Cols = 3
dtgQX.ColWidth(2) = 0
dtgQX.ColWidth(0) = 3450
dtgQX.Row = 0
dtgQX.Col = 0: dtgQX.Text = "经营项目": dtgQX.CellFontBold = True
dtgQX.Col = 1: dtgQX.Text = "选择": dtgQX.CellFontBold = True
dtgQX.Rows = 30
End Sub

Public Sub dtgRZFF()
dtgRZ.Cols = 4
dtgRZ.Rows = 30
dtgRZ.Row = 0
dtgRZ.Col = 0: dtgRZ.Text = "查询时间": dtgRZ.CellFontBold = True
dtgRZ.Col = 1: dtgRZ.Text = "供应商": dtgRZ.CellFontBold = True
dtgRZ.Col = 2: dtgRZ.Text = "联系人": dtgRZ.CellFontBold = True
dtgRZ.ColWidth(3) = 0
dtgRZ.ColWidth(1) = 2280
End Sub

