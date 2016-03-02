VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPF 
   Caption         =   "豪曼排行榜"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form2"
   ScaleHeight     =   5565
   ScaleWidth      =   9690
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   30
      ScaleHeight     =   4875
      ScaleWidth      =   2805
      TabIndex        =   8
      Top             =   30
      Width           =   2865
   End
   Begin VB.TextBox txtPLY 
      Height          =   3555
      Left            =   3030
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1980
      Width           =   3555
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frmPF.frx":0000
      Left            =   4710
      List            =   "frmPF.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1500
      Width           =   1935
   End
   Begin MSDataListLib.DataList dtgYwy 
      Height          =   5520
      Left            =   6720
      TabIndex        =   0
      Top             =   0
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   9737
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "请你预测，下一张大单子，将由哪个业务员开出？"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3030
      TabIndex        =   7
      Top             =   150
      Width           =   3795
   End
   Begin VB.Label Label2 
      Caption         =   "理由："
      Height          =   375
      Left            =   3060
      TabIndex        =   4
      Top             =   1500
      Width           =   1305
   End
   Begin VB.Label lblUid 
      Caption         =   "lblUid"
      Height          =   345
      Left            =   4080
      TabIndex        =   3
      Top             =   630
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lblYwy 
      Height          =   345
      Left            =   4710
      TabIndex        =   2
      Top             =   990
      Width           =   1845
   End
   Begin VB.Label Label1 
      Caption         =   "业务员选择："
      Height          =   375
      Left            =   3060
      TabIndex        =   1
      Top             =   990
      Width           =   1275
   End
End
Attribute VB_Name = "frmPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoYwy As ADODB.Recordset

Private Sub dtgYwy_Click()
lblYwy.Caption = dtgYwy.Text
lblYwy.ToolTipText = dtgYwy.BoundText
End Sub
