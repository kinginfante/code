VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmZhangTao 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请选择帐套"
   ClientHeight    =   2265
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgZT 
      Height          =   2235
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   3942
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
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   345
      Left            =   5220
      TabIndex        =   1
      Top             =   450
      Width           =   765
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   345
      Left            =   5220
      TabIndex        =   0
      Top             =   60
      Width           =   765
   End
End
Attribute VB_Name = "frmZhangTao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public ZT As Integer '帐套号
Public WDF As Boolean

Private Sub CancelButton_Click()
Me.Visible = False
End
End Sub

Private Sub dtgZT_Click()
If dtgZT.Row = 2 Then
    mod1.ZT = "HBData"
    WDF = True
Else
    mod1.ZT = "HMData"
    WDF = False
    If dtgZT.Row = 3 Then
        WDF = True
    End If
End If
End Sub

Private Sub Form_Load()
dtgZT.Cols = 2
dtgZT.Rows = 10
dtgZT.Row = 0
dtgZT.Col = 0: dtgZT.Text = "帐套名称": dtgZT.CellFontBold = True
dtgZT.Col = 1: dtgZT.Text = "说明": dtgZT.CellFontBold = True
dtgZT.Row = 1
dtgZT.Col = 0: dtgZT.Text = "上海豪曼": dtgZT.Col = 1: dtgZT.Text = "上海豪曼制冷空调服务有限公司"
dtgZT.Row = 2
dtgZT.Col = 0: dtgZT.Text = "北京豪曼必克": dtgZT.Col = 1: dtgZT.Text = "北京豪曼必克制冷空调服务有限公司"
dtgZT.Row = 3
dtgZT.Col = 0: dtgZT.Text = "豪曼外地办": dtgZT.Col = 1: dtgZT.Text = "豪曼2015年外地办新流程ERP"
dtgZT.ColWidth(1) = 4035
End Sub

Private Sub OKButton_Click()
'''''''If adoZt.Recordset.Fields("cid").Value = 1 Then
'''''''    Me.ZT = 1
'''''''Else
'''''''    Me.ZT = 2
'''''''End If
'''''''Me.Visible = False
'''''''frmLogin.lblZT.Caption = frmZhangTao.adoZt.Recordset.Fields("conM").Value
mod1.workKK = "Provider=SQLOLEDB.1;Password=kate135mxc;Persist Security Info=True;User ID=sa;Initial Catalog=" & mod1.ZT & ";Data Source=" & mod1.IP
Set mod1.cc = CreateObject("adodb.connection")
mod1.cc.Open mod1.workKK
Me.Visible = False
Form1.Show

End Sub
