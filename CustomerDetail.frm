VERSION 5.00
Begin VB.Form CustomerDetail 
   BackColor       =   &H00C0FFC0&
   Caption         =   "客户信息"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   LinkTopic       =   "Form2"
   ScaleHeight     =   9945
   ScaleWidth      =   13710
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtXmmc 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   4635
   End
   Begin VB.TextBox txtXmAdr 
      DataField       =   "xmAdr"
      DataSource      =   "adodm1"
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   900
      Width           =   4635
   End
   Begin VB.Label Label27 
      Caption         =   "项目名称"
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   60
      Width           =   765
   End
   Begin VB.Label Label34 
      Caption         =   "项目代码"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   510
      Width           =   885
   End
   Begin VB.Label lblXid 
      Caption         =   "lblXid"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   510
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "地  址"
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   900
      Width           =   585
   End
End
Attribute VB_Name = "CustomerDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
