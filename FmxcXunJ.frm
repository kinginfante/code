VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form FmxcXunJ 
   BackColor       =   &H00C0FFC0&
   Caption         =   "新版人工询价单"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frame1"
      Height          =   7305
      Left            =   60
      TabIndex        =   17
      Top             =   930
      Width           =   15135
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   6645
         Left            =   180
         TabIndex        =   18
         Top             =   330
         Width           =   14775
      End
   End
   Begin VB.CommandButton cmdHT 
      BackColor       =   &H00C0FFC0&
      Caption         =   "合同评审单"
      Height          =   345
      Left            =   13050
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   30
      Width           =   1785
   End
   Begin VB.TextBox txtXmmc 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   4500
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   30
      Width           =   5205
   End
   Begin VB.TextBox txtYfadr 
      BackColor       =   &H00FFFFC0&
      Height          =   270
      Left            =   4500
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   390
      Width           =   5205
   End
   Begin VB.TextBox txtBrq 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   11070
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   390
      Width           =   1485
   End
   Begin NiceFormControl.NiceButton cmdDht 
      Height          =   345
      Left            =   13020
      TabIndex        =   0
      Top             =   30
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FmxcXunJ.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Style           =   9
      Caption         =   "导入合同"
   End
   Begin MSComCtl2.DTPicker dtpBrq 
      Height          =   315
      Left            =   11160
      TabIndex        =   5
      Top             =   390
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   8454016
      CalendarTitleBackColor=   16711808
      CalendarTrailingForeColor=   -2147483635
      Format          =   102367233
      CurrentDate     =   38797
   End
   Begin VB.Label lblRq 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      Height          =   285
      Left            =   11070
      TabIndex        =   16
      Top             =   30
      Width           =   1485
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "询价日期"
      Height          =   195
      Left            =   10230
      TabIndex        =   15
      Top             =   90
      Width           =   885
   End
   Begin VB.Label lblZl 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   13620
      TabIndex        =   14
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "性质"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   13080
      TabIndex        =   13
      Top             =   480
      Width           =   585
   End
   Begin VB.Label lblBh 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   285
      Left            =   870
      TabIndex        =   12
      Top             =   30
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "编号"
      Height          =   285
      Left            =   180
      TabIndex        =   11
      Top             =   60
      Width           =   435
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "项目名称"
      Height          =   285
      Left            =   3450
      TabIndex        =   10
      Top             =   90
      Width           =   795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "业务员"
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   450
      Width           =   645
   End
   Begin VB.Label lblYwy 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   870
      TabIndex        =   8
      Top             =   390
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "地址"
      Height          =   285
      Left            =   3810
      TabIndex        =   7
      Top             =   450
      Width           =   555
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "报价有效期"
      Height          =   315
      Left            =   10110
      TabIndex        =   6
      Top             =   420
      Width           =   1065
   End
End
Attribute VB_Name = "FmxcXunJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
