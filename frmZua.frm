VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form frmZu 
   BackColor       =   &H00C0FFC0&
   Caption         =   "业务导航图"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   13875
   Begin VB.Timer timRev 
      Interval        =   5000
      Left            =   3330
      Top             =   5310
   End
   Begin VB.CommandButton cmdBack 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   13440
      Picture         =   "frmZua.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4980
      Width           =   435
   End
   Begin MSComctlLib.Toolbar TBa 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "公告栏"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "报销"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "事务"
            ImageIndex      =   16
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "2012年度奖项说明"
            ImageIndex      =   19
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "公司制度"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "电话簿"
            ImageIndex      =   12
            Style           =   2
            Object.Width           =   4500
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "会议记录"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "在线更新"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmTX 
      BackColor       =   &H00C0FFC0&
      Caption         =   "请选择头像"
      Height          =   4875
      Left            =   5040
      TabIndex        =   2
      Top             =   1440
      Width           =   4515
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消"
         Height          =   345
         Left            =   3750
         TabIndex        =   6
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   3750
         TabIndex        =   5
         Top             =   4380
         Width           =   645
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   4635
         Left            =   3120
         TabIndex        =   4
         Top             =   270
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   8176
         _Version        =   393216
         Orientation     =   1
         Max             =   20
      End
      Begin MSComctlLib.Toolbar tb2 
         Height          =   5130
         Left            =   90
         TabIndex        =   3
         Top             =   360
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   9049
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   43
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button40 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button41 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button42 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button43 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin VB.Image imgTX 
         Height          =   765
         Left            =   3720
         Top             =   1920
         Width           =   735
      End
   End
   Begin VB.Timer timOline 
      Interval        =   20000
      Left            =   3330
      Top             =   6180
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1890
      Top             =   4650
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   78
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":09DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":10B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":198C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":2266
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":2907
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":2DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":36A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":3F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":4860
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":553A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":5B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":642A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":6D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":6F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":74A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":760C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":82E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":8907
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":91E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":9ABF
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":A033
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":A90F
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":B1EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":B351
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":B8CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":C1A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":C7CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":C902
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":D1DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":D34F
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":DC2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":E507
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":EDE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":F6BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":FF9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":10877
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":11153
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":11A2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":1230B
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":147BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":15099
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":15975
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":16251
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":16B2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":17409
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":17CE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":185C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":18E9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":19779
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":19A93
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":2EC05
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":43D77
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":58EE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":6E05B
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":831CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":837FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":83DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":8436F
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":8490F
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":84F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":87602
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":89E04
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":8C438
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":8C9F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":8CFF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":8D5D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":8DC9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":90430
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":92B9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":95333
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":978E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":9A021
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":9C73A
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":9ED6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":A15C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":A3D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":A66D7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5730
      Top             =   5910
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":A9180
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":A95D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":A9A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":A9E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AA2C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AA71A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AAB6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AAFBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AB410
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AB862
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":ABCB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AC106
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AC558
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AC9AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":ACDFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AD24E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AD6A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":ADAF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":ADF44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   4260
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   60
      ImageHeight     =   60
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":AE396
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":B0E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":B389E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":B6322
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":B8DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":BB82A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   1440
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   70
      ImageHeight     =   60
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":BE2AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":C2416
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":C6052
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":C9AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":CCF55
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZua.frx":D0F42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolNew 
      Align           =   3  'Align Left
      Height          =   7200
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   12700
      ButtonWidth     =   2037
      ButtonHeight    =   2223
      Appearance      =   1
      Style           =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "维保业务"
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "采购服务"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "行政人事"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "维保记录"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "财务统计"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "管理信息"
            ImageIndex      =   6
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin NiceFormControl.NiceContainr NC 
      Height          =   7275
      Left            =   1230
      TabIndex        =   8
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   12832
      HeaderLightColor=   14214600
      HeaderDarkColor =   11651986
      BackLightColor  =   15857131
      BackDarkColor   =   16777152
      BorderColor     =   11191944
      TextColor       =   255
      Caption         =   "维保业务"
      Theme           =   2
      Begin VB.CommandButton cmdFocus 
         Caption         =   "Command2"
         Height          =   285
         Left            =   -840
         TabIndex        =   18
         Top             =   630
         Width           =   825
      End
      Begin NiceFormControl.NiceButton cmdBg 
         Height          =   735
         Index           =   0
         Left            =   270
         TabIndex        =   17
         Top             =   6300
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1296
         BTYPE           =   1
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmZua.frx":D463A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Caption         =   "管理"
      End
      Begin NiceFormControl.NiceButton cmdBf 
         Height          =   735
         Index           =   0
         Left            =   270
         TabIndex        =   16
         Top             =   5340
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1296
         BTYPE           =   1
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmZua.frx":D4656
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Caption         =   "财务"
      End
      Begin NiceFormControl.NiceButton cmdBe 
         Height          =   735
         Index           =   0
         Left            =   270
         TabIndex        =   15
         Top             =   4200
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1296
         BTYPE           =   1
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmZua.frx":D4672
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Caption         =   "产品"
      End
      Begin NiceFormControl.NiceButton cmdBB 
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   3210
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1296
         BTYPE           =   1
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmZua.frx":D468E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Caption         =   "行政"
      End
      Begin NiceFormControl.NiceButton cmdBc 
         Height          =   735
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   2100
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1296
         BTYPE           =   1
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmZua.frx":D46AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Caption         =   "服务"
      End
      Begin NiceFormControl.NiceButton cmdBu 
         Height          =   735
         Index           =   0
         Left            =   750
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1296
         BTYPE           =   1
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmZua.frx":D46C6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Caption         =   "维保业务"
      End
      Begin NiceFormControl.NiceContainr NiceContainr2 
         Height          =   6975
         Left            =   7530
         TabIndex        =   9
         Top             =   360
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   12303
         HeaderLightColor=   15526633
         HeaderDarkColor =   14276307
         BackLightColor  =   15857131
         BackDarkColor   =   16777152
         BorderColor     =   12632319
         TextColor       =   4867908
         Style           =   1
         Theme           =   10
         Begin NiceFormControl.NiceButton cmdXz 
            Height          =   825
            Left            =   4230
            TabIndex        =   11
            Top             =   5940
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1455
            BTYPE           =   14
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmZua.frx":D46E2
            PICN            =   "frmZua.frx":D46FE
            PICH            =   "frmZua.frx":D4FD8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            Style           =   4
            Caption         =   "豪曼军营"
         End
         Begin NiceFormControl.NiceButton NR 
            Height          =   825
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   450
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1455
            BTYPE           =   1
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmZua.frx":D58B2
            PICN            =   "frmZua.frx":D58CE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            Style           =   21
            Caption         =   "马晓聪"
         End
      End
      Begin VB.Label lblDtg 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   240
         Left            =   180
         TabIndex        =   19
         Top             =   390
         Visible         =   0   'False
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmZu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Bjxt As Object


Dim Location As Long

Dim MaxTop As Long
Dim MinTop As Long
Public meIndex As Integer
Dim ORa(100, 200) As String 'QQ按钮数组（防止重复刷新）
Dim OCount As Integer '以前的QQ人总数

Private Sub cmdBack_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next

        MDI.Cq = True
        Unload MDI
        mod1.FiR = False
        Form1.Show
        Form1.Fa1.GotoFrame (160)
        Call mod1.zhuLK '退出时取消注册
        Call mod1.DelDKZ
        If mod1.Mname = "马晓聪" Then
            ii = MsgBox("想不想做个游戏？ ：）", vbQuestion + vbYesNo, "轻松一下吧")
            If ii = vbNo Then
                End
            End If
            tt = "select username,userid from worker where xlx=0 and zzf=1 order by wid"
            Set frmPF.adoYwy = CreateObject("adodb.recordset")
            frmPF.adoYwy.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            Set frmPF.dtgYwy.RowSource = frmPF.adoYwy
            frmPF.dtgYwy.ListField = "username"
            frmPF.dtgYwy.BoundColumn = "userid"
            frmPF.Show
            
        End If
End Sub

Private Sub cmdBB_Click(Index As Integer)
Dim Ra
Dim La
Dim oo As Integer: Dim ii As Integer: Dim YY As Integer
Dim tt As String
On Error Resume Next

If cmdBB(Index).ToolTipText = "没有权限" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
    Exit Sub
End If

Call frmZu.OLOff
Select Case Trim(cmdBB(Index).Caption)
Case "课程统计"
    frmPei.Show
    Call frmPei.Qing
    Call frmPei.JLBound
    frmPei.ZOrder 0
Case "员工培训"
    frmPeiView.Show
    Call frmPeiView.dtgbrFF
    Call frmPeiView.dtgDeFF
    frmPeiView.ZOrder 0
Case "合同评审"

'''    htBrowG.Visible = False
'''
'''        tt = "Select * from htView1  order by 合同日期 desc"
'''
'''    htBrowG.adoBr.Close
'''    htBrowG.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''    Set htBrowG.dtgBr.DataSource = htBrowG.adoBr
'''    If htBrowG.adoBr.RecordCount > 0 Then
'''        htBrowG.dtgBr.FixedRows = 0
'''        htBrowG.dtgBr.MergeCol(1) = True
'''        htBrowG.dtgBr.MergeCol(2) = True
'''        htBrowG.dtgBr.MergeCol(3) = True
'''        htBrowG.dtgBr.MergeCol(4) = True
'''        htBrowG.dtgBr.MergeCol(7) = True
'''        htBrowG.dtgBr.MergeCol(13) = True
'''        htBrowG.dtgBr.MergeCells = 3
'''        htBrowG.dtgBr.FixedRows = 1
'''    End If
    If mod1.KhK = 1 Then
        htBrowG.lblFw.Caption = mod1.Bm
    Else
        htBrowG.lblFw.Caption = "业务部"
    End If

    htBrowG.Visible = True
    htBrowG.ZOrder 0
Case "员工档案"
frmRen.Visible = True
If mod1.Qy = "上海" Then
    tt = "select userid as 工号,username as 姓名,qy as 区域,bm as 部门,userzw as 职务,nx as 工作年限 from worker where zzf=1 order by userid"
Else
    tt = "select userid as 工号,username as 姓名,qy as 区域,bm as 部门,userzw as 职务,nx as 工作年限 from worker where zzf=1 and qy='" & mod1.Qy & "' order by userid"
End If
'''''''''frmRen.adoRen.Close
'''''''''frmRen.adoRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''Set frmRen.dtgRen.DataSource = frmRen.adoRen
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
La = UBound(Ra, 2) + 1
mod1.HTP.Close
Set mod1.HTP = Nothing
Call frmRen.RenQing
frmRen.dtgRen.Clear: frmRen.dtgRN.Clear
frmRen.dtgRen.Visible = False
frmRen.dtgRen.Rows = La + 10: frmRen.dtgRen.Cols = 10
frmRen.dtgRen.Row = 0: frmRen.dtgRen.Col = 1: frmRen.dtgRen.Text = "工号": frmRen.dtgRen.Col = 2: frmRen.dtgRen.Text = "姓名"
frmRen.dtgRen.Col = 3: frmRen.dtgRen.Text = "区域": frmRen.dtgRen.Col = 4: frmRen.dtgRen.Text = "部门":
frmRen.dtgRen.Col = 5: frmRen.dtgRen.Text = "职务": frmRen.dtgRen.Col = 6: frmRen.dtgRen.Text = "工作年限"
frmRen.dtgRN.Rows = frmRen.dtgRen.Rows: frmRen.dtgRN.Cols = frmRen.dtgRen.Cols
For oo = 1 To La + 1
    frmRen.dtgRen.Row = oo: frmRen.dtgRN.Row = oo
    For ii = 1 To 10
        frmRen.dtgRen.Col = ii: frmRen.dtgRN.Col = ii
        frmRen.dtgRen.Text = Ra(ii - 1, oo - 1)
        frmRen.dtgRN.Text = frmRen.dtgRen.Text
    Next
Next
frmRen.dtgRen.Visible = True
frmRen.frmMod.Enabled = False
frmRen.txtTang.Visible = True
Case "固定资产"
    frmComputer.Show
    Call frmComputer.OlineF
    frmComputer.ZOrder 0
Case Else
    MsgBox ("正在建设中2")
    Call frmZu.OLOn
End Select
End Sub

Private Sub cmdBBzx_Click(Index As Integer)
If mod1.DName <> "潘明峰" And mod1.BmJl = True Or mod1.DName = "徐瑛" Then
    frmBB.Show
End If
End Sub


Private Sub cmdBc_Click(Index As Integer)
Dim tt As String
Dim oo As Integer
Dim jj As Integer
Dim PP As String
Dim Ra, Rb
Dim ua As Long
On Error Resume Next

If cmdBc(Index).ToolTipText = "没有权限" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
    Exit Sub
End If

frmZu.Enabled = False
frmGxBiao.cmdCreat.Visible = True
frmGxBiao.cmdDx.Visible = True
frmGxBiao.cmdNew.Visible = True
Call frmZu.OLOff
Select Case Trim(cmdBc(Index).Caption)
Case "采购合同"
    Call FmxcCG.CGBound
    Call FmxcCG.CDBound
    FmxcCG.Show
    FmxcCG.ZOrder 0
Case "货品资料"
    frmHPBR.frmZX.Visible = False
    frmHPBR.frmZX.Visible = False
    If mod1.DName <> "倪东海" And mod1.DName <> "李午阳" And mod1.DName <> "邹晨" And mod1.DName <> "马晓聪" And mod1.DName <> "货品录入员" Then
        Call frmHPBR.dtgLPFF
        frmHPBR.Show
        frmHPBR.ZOrder 0
        Exit Sub
    End If
    Call frmHPZL.Qing
    Call frmHPZL.BoundL1
    Call frmHPZL.dtgL2FF
    Call frmHPZL.dtgL3FF
    frmHPZL.dtgL1.Visible = False
    frmHPZL.Show
        Call frmHPBR.dtgLPFF
        Call frmHPBR.dtgFF
        frmHPBR.Show
        frmHPBR.ZOrder 0
        If mod1.DName = "邹晨" Then
            frmHPZL.frm3.Visible = True
            frmHPZL.frm2.Visible = False
        End If
        Exit Sub
Case "网上销售"
    tt = "select tb_order.addtime,tb_member.username,tb_usergroup.name,tb_order.sum,tb_order.status,tb_order.id,tb_usergroup.groupid,tb_order.uid from tb_order left outer join tb_member on " & _
        "tb_order.uid=tb_member.uid left outer join tb_usergroup on tb_member.groupid=tb_usergroup.groupid"
'''''        tt = "select * from dbo.tb_member"
'''''tt = "select tb_order.addtime,tb_member.username,tb_usergroup.name,tb_order.sum,tb_order.status,tb_order.id,'aaa','ok' from tb_order left outer join tb_member on " & _
'''''        "tb_order.uid=tb_member.uid left outer join tb_usergroup on tb_member.groupid=tb_usergroup.groupid"
'''''        tt = "select * from dbo.tb_order"
    Call FmxcNet.Bound(tt)
    FmxcNet.Show
Case "成本追加单"
    If mod1.Qy = "上海" Then
        tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc from htzuiView where fbf=0 order by ztime desc,zid desc"
    Else
        tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc from htzuiView  where qy='" & mod1.Qy & "' and fbf=0  order by ztime desc,zid desc"
    End If
    Call FmxcZuiBrow.Bound(tt)
    FmxcZuiBrow.tt = tt
    FmxcZuiBrow.Show
    FmxcZuiBrow.ZOrder 0
Case "供应商"
    Call frmGy.dtgBFF
    frmGy.Show
    frmGy.ZOrder 0
Case "合同执行"
    If mod1.Bm = "零件事业部" Then
    tt = "SELECT dbo.htzx.zid, dbo.htPing.htBh, dbo.htPing.khMc, dbo.htzx.xz, dbo.htzx.BZ, dbo.htzx.zTime, SUM(dbo.htzFk.amount) AS amount, dbo.htzx.ZT" & _
        " FROM dbo.htPing INNER JOIN dbo.htzx ON dbo.htPing.Hid = dbo.htzx.hid INNER JOIN dbo.htzFk ON dbo.htzx.zid = dbo.htzFk.zid" & _
        " Where (dbo.htzFk.Pwf = 1 and dbo.htzx.xz='配件' ) GROUP BY dbo.htzx.zid, dbo.htPing.htBh, dbo.htPing.khMc, dbo.htzx.xz, dbo.htzx.BZ, dbo.htzx.zTime, dbo.htzx.ZT" & _
        " order by dbo.htzx.ztime desc"
    Else
    tt = "SELECT dbo.htzx.zid, dbo.htPing.htBh, dbo.htPing.khMc, dbo.htzx.xz, dbo.htzx.BZ, dbo.htzx.zTime, SUM(dbo.htzFk.amount) AS amount, dbo.htzx.ZT" & _
        " FROM dbo.htPing INNER JOIN dbo.htzx ON dbo.htPing.Hid = dbo.htzx.hid INNER JOIN dbo.htzFk ON dbo.htzx.zid = dbo.htzFk.zid" & _
        " Where (dbo.htzFk.Pwf = 1 and dbo.htzx.xz='产品' ) GROUP BY dbo.htzx.zid, dbo.htPing.htBh, dbo.htPing.khMc, dbo.htzx.xz, dbo.htzx.BZ, dbo.htzx.zTime, dbo.htzx.ZT" & _
        " order by dbo.htzx.ztime desc"
    End If
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        ua = UBound(Ra, 2) + 1
        Call frmHtZX.Bound(Ra, ua)
    frmHtZX.Show
    frmHtZX.ZOrder 0
Case "询价单"
''''''''    mod1.BTZ = 36
''''''''    frmGxBiao.Visible = False
''''''''    frmGxBiao.cmdQH.Visible = False
''''''''    frmGxBiao.cmdZF.Visible = False
''''''''    If mod1.DName = "徐瑛" Then
''''''''        tt = "select top 100 * from xunjiaview order by 询价日期 desc"
''''''''        frmGxBiao.adoXj.Close
''''''''        frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''        Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
''''''''    ElseIf mod1.BM = "零件事业部" Or mod1.DName = "周春云" Then
''''''''        tt = "select top 100 * from xunjiaview  order by bid desc"
''''''''        frmGxBiao.adoXj.Close
''''''''        frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''        Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
''''''''        If frmGxBiao.adoXj.RecordCount > 1 Then
''''''''            frmGxBiao.dtgXj.FixedRows = 0
''''''''            frmGxBiao.dtgXj.MergeCol(1) = True
''''''''            frmGxBiao.dtgXj.MergeCol(2) = True
''''''''            frmGxBiao.dtgXj.MergeCol(3) = True
''''''''            frmGxBiao.dtgXj.MergeCol(4) = True
''''''''            frmGxBiao.dtgXj.MergeCol(5) = True
''''''''            frmGxBiao.dtgXj.MergeCells = 3
''''''''            frmGxBiao.dtgXj.FixedRows = 1
''''''''        End If
''''''''    ElseIf mod1.Qy = "北京" Or mod1.DName = "周春云" Then
''''''''        tt = "select top 100 * from xunjiaview  order by bid desc"
''''''''        frmGxBiao.adoXj.Close
''''''''        frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''        Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
''''''''        If frmGxBiao.adoXj.RecordCount > 1 Then
''''''''            frmGxBiao.dtgXj.FixedRows = 0
''''''''            frmGxBiao.dtgXj.MergeCol(1) = True
''''''''            frmGxBiao.dtgXj.MergeCol(2) = True
''''''''            frmGxBiao.dtgXj.MergeCol(3) = True
''''''''            frmGxBiao.dtgXj.MergeCol(4) = True
''''''''            frmGxBiao.dtgXj.MergeCol(5) = True
''''''''            frmGxBiao.dtgXj.MergeCells = 3
''''''''            frmGxBiao.dtgXj.FixedRows = 1
''''''''        End If
''''''''
''''''''    Else
''''''''
''''''''        tt = "select * from xunjiaView where lx=0 and qy='" & mod1.Qy & "' order by bid desc"
''''''''        frmGxBiao.adoXj.Close
''''''''        frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''        Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
''''''''        frmGxBiao.dtgXj.FixedRows = 0
''''''''        frmGxBiao.dtgXj.MergeCol(1) = True
''''''''        frmGxBiao.dtgXj.MergeCol(3) = True
''''''''        frmGxBiao.dtgXj.MergeCol(4) = True
''''''''        frmGxBiao.dtgXj.MergeCells = 3
''''''''        frmGxBiao.dtgXj.FixedRows = 1
''''''''    End If
        If mod1.Qy <> "上海" Then
            tt = "select * from xunjiaView where qy='" & mod1.Qy & "' order by bid desc"
        Else
            tt = "select * from xunjiaView order by bid desc"
        End If
        frmGxBiao.adoXj.Close
        frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
        frmGxBiao.dtgXj.FixedRows = 0
        frmGxBiao.dtgXj.MergeCol(1) = True
        frmGxBiao.dtgXj.MergeCol(3) = True
        frmGxBiao.dtgXj.MergeCol(4) = True
        frmGxBiao.dtgXj.MergeCells = 3
        frmGxBiao.dtgXj.FixedRows = 1
    frmGxBiao.Visible = True
    frmGxBiao.frmNew.Visible = False
'    '取得新建维保询价单及购销询价单的流程参数
'    tt = "xunJiaBut('" & mod1.DName & "','" & mod1.DHid & "','维保询价')"
'    Set mod1.HTP = CreateObject("adodb.recordset")
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'    frmGxBiao.cmdNew.Tag = mod1.HTP.Fields("nlb").Value
'    frmGxBiao.cmdNew.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
'    frmGxBiao.cmdDx.Tag = mod1.HTP.Fields("nlb").Value                          '大修的流程同维保
'    frmGxBiao.cmdDx.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
'    tt = "xunJiaBut('" & mod1.DName & "','" & mod1.DHid & "','购销')"
'    Set mod1.HTP = CreateObject("adodb.recordset")
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'    frmGxBiao.cmdCreat.Tag = mod1.HTP.Fields("nlb").Value
'    frmGxBiao.cmdCreat.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
    frmGxBiao.cmdCreat.Visible = False
    frmGxBiao.cmdDx.Visible = False
    frmGxBiao.cmdNew.Visible = False
    frmGxBiao.frmC.Visible = True
Case "项目执行"
    'frmGxNew.Show
    frmHtZxG.Visible = False
    'tt = "Select * from htView where (业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
    "') order by 合同日期 desc"
    tt = "select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 业务员='马晓聪' and 状态='执行' order by 合同日期 desc"
    frmHtZxG.adoBr.Close
    frmHtZxG.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmHtZxG.dtgBr.DataSource = frmHtZxG.adoBr
    If frmHtZxG.adoBr.RecordCount > 0 Then
        frmHtZxG.dtgBr.FixedRows = 0
        frmHtZxG.dtgBr.MergeCol(1) = True
        frmHtZxG.dtgBr.MergeCol(2) = True
        frmHtZxG.dtgBr.MergeCol(3) = True
        frmHtZxG.dtgBr.MergeCells = 3
        frmHtZxG.dtgBr.FixedRows = 1
    End If
    
    frmHtZxG.lblHtbh.Caption = ""
    tt = "select * from PldView where htbh='" & frmHtZxG.lblHtbh.Caption & "' order by 编号"
    frmHtZxG.adoPld.Close
    frmHtZxG.adoPld.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmHtZxG.dtgPld.DataSource = frmHtZxG.adoPld
    frmHtZxG.Show
    'frmHtZXg.optY(0).Value = True
    frmHtZxG.ZOrder 0

Case "合同评审"
'''''''    FR = DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda) - 14)
'''''''    htBrowG.Visible = False
'''''''    If mod1.KhK = 1 Then
'''''''        tt = "Select 项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr from htView1 where ggl='" & mod1.DHid & "' and htrq>='" & FR & "' and 状态<>'评审'  order by htrq desc"
'''''''
'''''''    ElseIf mod1.KhK = 2 Then
'''''''        tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htView1 where not(部门='维销部3' or 部门='产品部1' or 部门='产品部2')  and htrq>='" & FR & "' and 状态<>'评审' order by htrq desc"
'''''''    ElseIf mod1.KhK = 3 Then
'''''''        tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htView1 where htrq>='" & FR & "' and 状态<>'评审' order by htrq desc"
'''''''    End If
''''''''''''    If mod1.BM = "工程二部" Then
''''''''''''        tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0   from htViewP where 状态<>'评审'   order by 部门,htrq desc"
''''''''''''    End If
'''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''''    Ra = mod1.HTP.GetRows
'''''''    mod1.HTP.Close
'''''''    Set mod1.HTP = Nothing
'''''''    ua = UBound(Ra, 2)
''''''''''    htBrowG.dtgBr.Clear
''''''''''    htBrowG.dtgBr.Row = 0: htBrowG.dtgBr.Col = 1: htBrowG.dtgBr.Text = "项目归属人"
''''''''''    htBrowG.dtgBr.Col = 2: htBrowG.dtgBr.Text = "项目名称": htBrowG.dtgBr.Col = 3: htBrowG.dtgBr.Text = "合同日期": htBrowG.dtgBr.Col = 4: htBrowG.dtgBr.Text = "合同性质"
''''''''''    htBrowG.dtgBr.Col = 5: htBrowG.dtgBr.Text = "合同金额": htBrowG.dtgBr.Col = 6: htBrowG.dtgBr.Text = "合同编号": htBrowG.dtgBr.Col = 7: htBrowG.dtgBr.Text = "状态"
'''''''
'''''''    Call htBrowG.Bref(Ra, ua + 1)
'''''''    If mod1.KhK = 1 Then
'''''''        htBrowG.lblFw.Caption = mod1.BM
'''''''    Else
'''''''        htBrowG.lblFw.Caption = "业务部"
'''''''    End If

    htBrowG.Visible = True
    htBrowG.ZOrder 0
Case "零配件库"
    Call frmLPNew.dtgFF
    frmLPNew.Show
'''''    frmLPG.Show
    frmZu.Enabled = False
Case "收款情况"
    On Error Resume Next
    PP = InputBox("请键入合同编号(后五位数字部分)", "请录入")
    If Len(PP) > 5 Then
        tt = "select htze from htping where htbh='" & Trim(PP) & "';" & _
           "select sum(je) from htAview where htbh='" & Trim(PP) & "' and lc=100 group by htbh"
    Else
        tt = "declare @htbh nvarchar(22),@LcUid nvarchar(22);" & _
        "select @htbh=htbh,@LcUid=lcuid from htping where hid=" & Val(PP) & ";" & _
        "select htze from htping where htbh=@htbh;" & _
        "select sum(je) from htAview where htbh=@htbh and lc=100 group by htbh"
    End If
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        Set mod1.HTP = mod1.HTP.NextRecordset
        Rb = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        jj = MsgBox("收款额度为:" & Str(Round(Val(Rb(0, 0)) / Val(Ra(0, 0)) * 100, 2)) & "%", vbInformation, "Hello")
        frmZu.Enabled = True
Case Else
    MsgBox "正在建设中!3"
    frmZu.Enabled = True
    Call frmZu.OLOn
End Select
End Sub

Private Sub cmdBd_Click(Index As Integer)
'Select Case Trim(cmdBd(Index).Caption)
'Case "合同评审"
'
'    htBrowG.Visible = False
'
'        tt = "Select * from htView1  order by 合同日期 desc"
'
'    htBrowG.adoBr.Close
'    htBrowG.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    Set htBrowG.dtgBr.DataSource = htBrowG.adoBr
'    If htBrowG.adoBr.RecordCount > 0 Then
'        htBrowG.dtgBr.FixedRownyns = 0
'        htBrowG.dtgBr.MergeCol(1) = True
'        htBrowG.dtgBr.MergeCol(2) = True
'        htBrowG.dtgBr.MergeCol(3) = True
'        htBrowG.dtgBr.MergeCol(4) = True
'        htBrowG.dtgBr.MergeCol(7) = True
'        htBrowG.dtgBr.MergeCol(13) = True
'        htBrowG.dtgBr.MergeCells = 3
'        htBrowG.dtgBr.FixedRows = 1
'    End If
'    If mod1.KhK = 1 Then
'        htBrowG.lblFw.Caption = mod1.Bm
'    Else
'        htBrowG.lblFw.Caption = "业务部"
'    End If
'
'    htBrowG.Visible = True
'    htBrowG.ZOrder 0
'Case Else
'    MsgBox ("正在建设中")
'End Select
End Sub


Private Sub cmdBe_Click(Index As Integer)
Dim tt As String
Dim oo As Integer
Dim ii As Integer
Dim jj As Integer
Dim Ra: Dim La As Integer
Dim ua As Long
On Error Resume Next

If cmdBe(Index).ToolTipText = "没有权限" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
    Exit Sub
End If

frmZu.Enabled = False
frmGxBiao.cmdCreat.Visible = True
frmGxBiao.cmdDx.Visible = True
frmGxBiao.cmdNew.Visible = True
Call frmZu.OLOff
Select Case Trim(cmdBe(Index).Caption)
Case "出工统计"
frmWork.Show
frmWork.ZOrder 0
Case "合同执行"
    tt = "SELECT dbo.htzx.zid, dbo.htPing.htBh, dbo.htPing.khMc, dbo.htzx.xz, dbo.htzx.BZ, dbo.htzx.zTime, SUM(dbo.htzFk.amount) AS amount, dbo.htzx.ZT" & _
        " FROM dbo.htPing INNER JOIN dbo.htzx ON dbo.htPing.Hid = dbo.htzx.hid INNER JOIN dbo.htzFk ON dbo.htzx.zid = dbo.htzFk.zid" & _
        " Where (dbo.htzFk.Pwf = 1 and dbo.htzx.xz<>'配件' and dbo.htzx.xz<>'产品') GROUP BY dbo.htzx.zid, dbo.htPing.htBh, dbo.htPing.khMc, dbo.htzx.xz, dbo.htzx.BZ, dbo.htzx.zTime, dbo.htzx.ZT" & _
        " order by dbo.htzx.ztime desc"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        ua = UBound(Ra, 2) + 1
        Call frmHtZX.Bound(Ra, ua)
    frmHtZX.Show
    frmHtZX.ZOrder 0
Case "人工询价"

    mod1.BTZ = 36
    frmGxBiao.frmNew.Visible = False
    frmGxBiao.Visible = False
    frmGxBiao.comLx.Text = "项目名称"
    frmGxBiao.comLx.Locked = True
    If mod1.Qy = "北京" Or mod1.DName = "周春云" Then
       ' tt = "select * from xunjiaView where 类型<>'购销' and 类型<>'配件' and 类型<>'产品' and comid=0 and lc>=4 order by 询价日期 desc"
        tt = "select * from xunjiaView where 类型<>'购销' and 类型<>'配件' and 类型<>'产品' and comid=0 order by 询价日期 desc"
'    ElseIf mod1.DName = "郑刚" Then
'        tt = "select * from xunjiaView where qy<>'上海' and 类型<>'购销' and comid=0 and lc>=4"

    Else '组长
        tt = "select * from xunjiaView where 类型<>'购销' and 类型<>'配件' and 类型<>'产品' and comid=0 and lc>=4 order by 询价日期 desc"
    End If
''''''    frmGxBiao.adoXj.Close
''''''    frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''    Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
''''''    If frmGxBiao.adoXj.RecordCount > 1 Then
''''''        frmGxBiao.dtgXj.FixedRows = 0
''''''        frmGxBiao.dtgXj.MergeCol(1) = True
''''''        frmGxBiao.dtgXj.MergeCol(3) = True
''''''        frmGxBiao.dtgXj.MergeCol(4) = True
''''''        frmGxBiao.dtgXj.MergeCells = 3
''''''        frmGxBiao.dtgXj.FixedRows = 1
''''''    End If
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2) + 1
    Call frmGxBiao.XJBound(Ra, La)
    frmGxBiao.cmdZF.Visible = False
    frmGxBiao.cmdQH.Visible = False
    frmGxBiao.dtgXj.Row = 0: frmGxBiao.dtgXj.Col = 3
    frmGxBiao.dtgXj.Text = "类型"
    frmGxBiao.dtgXj.ColWidth(1) = 1000
    frmGxBiao.dtgXj.ColWidth(2) = 2500
'''''''''    '显示工程部询价表
'''''''''    If mod1.DName = "张寅" Or mod1.DName = "徐瑛" Then
'''''''''        tt = "select * from xunjiagcv where comid=0"
''''''''''    ElseIf mod1.DName = "郑刚" Then
''''''''''        tt = "select * from xunjiagcv where qy<>'上海' and comid=0"
'''''''''    ElseIf mod1.DName = "彭海翔" Then
'''''''''        tt = "select * from xunjiagcv where comid=1"
'''''''''    Else '组长
'''''''''        tt = "select * from xunjiagcv where uid='" & mod1.DHid & "'"
'''''''''    End If
'''''''''    frmGxBiao.adoGc.Close
'''''''''    frmGxBiao.adoGc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''    Set frmGxBiao.dtgGc.DataSource = frmGxBiao.adoGc
'''''''''    If frmGxBiao.adoGc.RecordCount > 1 Then
'''''''''        frmGxBiao.dtgGc.FixedRows = 0
'''''''''        frmGxBiao.dtgGc.MergeCol(1) = True
'''''''''        frmGxBiao.dtgGc.MergeCol(3) = True
'''''''''        frmGxBiao.dtgGc.MergeCol(4) = True
'''''''''        frmGxBiao.dtgGc.MergeCells = 3
'''''''''        frmGxBiao.dtgGc.FixedRows = 1
'''''''''    End If
    
    frmGxBiao.Visible = True
    frmGxBiao.cmdCreat.Visible = False
    frmGxBiao.cmdDx.Visible = False
    frmGxBiao.cmdNew.Visible = False

Case "工作单检验"
    If Not (mod1.DName = "" Or mod1.Bq2 = True) Then
        Exit Sub
    End If
    Me.Enabled = False
    frmGZDJY.Show
    frmGZDJY.ZOrder 0
    tt = "select * from gzdView where qy='" & mod1.Qy & "' and trq is null order by gid desc"
    frmGZDJY.adoGZD.Close
    frmGZDJY.adoGZD.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGZDJY.dtgGzd.DataSource = frmGZDJY.adoGZD
    For oo = 1 To 8
        frmGZDJY.optLx(oo).Value = False
    Next
'    If mod1.comId = 1 Then
'        frmGZDJY.lblFw.Caption = "陈文珍"
'        frmGZDJY.lblFw.ToolTipText = "HMG012"
'    End If
    
    If frmGZDJY.adoGZD.RecordCount > 0 Then
        frmGZDJY.dtgGzd.FixedRows = 0
        frmGZDJY.dtgGzd.MergeCol(2) = True
        'frmGZDJY.dtgGzd.MergeCol(3) = True
        'frmGZDJY.dtgGzd.MergeCol(4) = True
        frmGZDJY.dtgGzd.MergeCells = 3
        frmGZDJY.dtgGzd.FixedRows = 1
    End If
Case "合同评审"
'''''    htBrowG.Visible = False
'''''    tt = "select * from htview1 where 合同性质='I love you' "
'''''
'''''    htBrowG.adoBr.Close
'''''    htBrowG.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''    Set htBrowG.dtgBr.DataSource = htBrowG.adoBr
'''''    If htBrowG.adoBr.RecordCount > 0 Then
'''''        htBrowG.dtgBr.FixedRows = 0
'''''        htBrowG.dtgBr.MergeCol(1) = True
'''''        htBrowG.dtgBr.MergeCol(2) = True
'''''        htBrowG.dtgBr.MergeCol(3) = True
'''''        htBrowG.dtgBr.MergeCol(4) = True
'''''        htBrowG.dtgBr.MergeCol(7) = True
'''''        htBrowG.dtgBr.MergeCol(13) = True
'''''        htBrowG.dtgBr.MergeCells = 3
'''''        htBrowG.dtgBr.FixedRows = 1
'''''    End If
'''''    If mod1.KhK = 1 Then
'''''        htBrowG.lblFw.Caption = mod1.BM
'''''    Else
'''''        htBrowG.lblFw.Caption = "业务部"
'''''    End If
    If mod1.Bm = "配送中心" Or mod1.Bm = "配送中心" Then
         tt = "select 项目名称,询价日期,人工费,最低销售价,类型,编号,业务员,bid,uid from xunjiaZC order by bid desc"
    Else '工程部不能看见配件
         tt = "select 项目名称,询价日期,人工费,最低销售价,类型,编号,业务员,bid,uid from xunjiaZC where NOT (类型 = '配件' OR 类型 = '产品' OR 类型= '购销' OR 类型= '零配件') order by bid desc"
    End If
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2)
    htBrowG.dtgBr.Row = 0: htBrowG.dtgBr.Col = 1: htBrowG.dtgBr.Text = "项目名称": htBrowG.dtgBr.Col = 2: htBrowG.dtgBr.Text = "询价日期"
    htBrowG.dtgBr.Col = 3: htBrowG.dtgBr.Text = "人工费": htBrowG.dtgBr.Col = 4: htBrowG.dtgBr.Text = "最低销售价": htBrowG.dtgBr.Col = 5: htBrowG.dtgBr.Text = "类型"
    htBrowG.dtgBr.Col = 6: htBrowG.dtgBr.Text = "编号": htBrowG.dtgBr.Col = 7: htBrowG.dtgBr.Text = "业务员": htBrowG.dtgBr.ColWidth(8) = 0: htBrowG.dtgBr.ColWidth(9) = 0
    htBrowG.dtgBr.ColWidth(1) = 2500: htBrowG.dtgBr.ColWidth(2) = 1000
    For oo = 1 To La + 1
        htBrowG.dtgBr.Row = oo
        For ii = 1 To 9
            htBrowG.dtgBr.Col = ii
            htBrowG.dtgBr.Text = Ra(ii - 1, oo - 1)
        Next
    Next
    htBrowG.Visible = True
    htBrowG.ZOrder 0
Case "出工统计"
    frmGZDBR.Show
    frmGZDBR.frmFw.Visible = True
Case "报销管理"
    mod1.BTZ = 23
    frmBxV.Visible = False
    frmBxV.mtA.Value = mod1.DQda
    tt = "select * from FydBrowG where 部门 like '%工程部%' and comid=" & mod1.comId & " order by 签收日期 desc"
    frmBxV.adoBxV.Close
    frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmBxV.dtgBx.DataSource = frmBxV.adoBxV
    If frmBxV.adoBxV.RecordCount > 0 Then
        frmBxV.dtgBx.FixedRows = 0
        frmBxV.dtgBx.MergeCol(1) = True
        frmBxV.dtgBx.MergeCol(2) = True
        frmBxV.dtgBx.MergeCol(3) = True
        frmBxV.dtgBx.MergeCol(4) = True
        frmBxV.dtgBx.MergeCol(5) = True
        frmBxV.dtgBx.MergeCol(7) = True
        frmBxV.dtgBx.MergeCells = 3
        frmBxV.dtgBx.FixedRows = 1
    End If
    frmBxV.Show
    frmBxV.ZOrder 0
Case "施工计划"
    If mod1.DName = "张寅" Then
        tt = "select * from GjwV where (trq is null or trq>='" & Date & "') and comid=0 "
'    ElseIf mod1.DName = "郑刚" Then
'        tt = "select * from GjwV where (trq is null or trq>='" & Date & "') and comid=0 and qy<>'上海'"
    ElseIf mod1.DName = "彭海翔" Then
        tt = "select * from GjwV where (trq is null or trq>='" & Date & "') and comid=1"
    Else '组长
        tt = "select * from GjwV where (trq is null or trq>='" & Date & "') and 组长='" & mod1.DName & "'"
    End If
    Set frmGjwV.adoGJW = CreateObject("adodb.recordset")
    frmGjwV.adoGJW.Close
    frmGjwV.adoGJW.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGjwV.dtgBr.DataSource = frmGjwV.adoGJW
    If frmGjwV.adoGJW.RecordCount > 0 Then
        frmGjwV.dtgBr.FixedRows = 0
        frmGjwV.dtgBr.MergeCol(2) = True
        frmGjwV.dtgBr.MergeCol(3) = True
        frmGjwV.dtgBr.MergeCol(4) = True
        frmGjwV.dtgBr.MergeCells = 3
        frmGjwV.dtgBr.FixedRows = 1
    End If
    frmGjwV.Visible = True
    frmGjwV.Enabled = True
    frmGjwV.ZOrder 0
Case "维保记录"
    frmWBjl.lblZNAME.Caption = ""
    frmWBjl.comLx.Text = "项目名称"
    frmWBjl.txtZ.Text = ""
    frmWBjl.cmdSave.Enabled = False
    frmWBjl.frmMod.Visible = False
    tt = "select * from wbjl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "' and yy=" & Year(mod1.DQda) & " order by jid"
    frmWBjl.adoNr.Recordset.Close
    If mod1.Zuf = 1 Then
        frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        frmWBjl.cmdMod.Enabled = True
    Else
        frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmWBjl.cmdMod.Enabled = False
    End If
    Set frmWBjl.dtpNr.DataSource = frmWBjl.adoNr
    If mod1.Zuf = False Or mod1.DName = "张寅" Then
        frmWBjl.frmRen.Visible = True
        frmWBjl.cmdMod.Enabled = False
    Else
        frmWBjl.frmRen.Visible = False
        frmWBjl.lblZNAME = mod1.DName
        frmWBjl.cmdMod.Enabled = True
    End If
    frmWBjl.Show
    frmWBjl.ZOrder 0
Case "财务报表"
    fyBB.Show
    fyBB.ZOrder 0
    Me.Enabled = False
Case "采购信息"
    frmTDCG.Show
    frmTDCG.ZOrder 0
    Me.Enabled = False
    tt = "select * from xunjiagcj where month(询价日期)=" & Month(mod1.DQda) & " and year(询价日期)=" & Year(mod1.DQda) & " order by 询价日期 desc"
    frmTDCG.adoGCJ.Close
    frmTDCG.adoGCJ.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If frmTDCG.adoGCJ.RecordCount > 0 Then
        Set frmTDCG.dtgCGj.DataSource = frmTDCG.adoGCJ
        frmTDCG.dtgCGj.FixedRows = 1
        frmTDCG.dtgCGj.Row = frmTDCG.adoGCJ.RecordCount - 1
    Else
        'Set frmTDCG.dtgCGj.DataSource = frmTDCG.adoGCJ
        frmTDCG.dtgCGj.Rows = 2
        frmTDCG.dtgCGj.FixedRows = 1
        frmTDCG.dtgCGj.Row = 1
        For oo = 0 To 10
            frmTDCG.dtgCGj.Col = oo
            frmTDCG.dtgCGj.Text = ""
        Next
    End If
    frmTDCG.monDate.Value = Date
    frmTDCG.Drq = DateSerial(Year(mod1.DQda), Month(mod1.DQda), 1)
Case "人工资料"
    fmxcRGB.Show
    fmxcRGB.ZOrder 0
    frmZu.Enabled = True
Case Else
    MsgBox "正在建设中!3"
    frmZu.Enabled = True
    Call frmZu.OLOn
End Select
End Sub

Private Sub cmdBf_Click(Index As Integer)
Dim Ra
Dim ua As Long
Dim tt As String
On Error Resume Next

If cmdBf(Index).ToolTipText = "没有权限" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
    Exit Sub
End If

frmZu.Enabled = False
'MsgBox cmdBu(Index).Caption
Call frmZu.OLOff
Select Case Trim(cmdBf(Index).Caption)
Case "执行报表"
    fmxcZfile.Show
    fmxcZfile.Visible = True
    fmxcZfile.ZOrder 0
    
    Call fmxcZfile.Bound
Case "付款沟通"
    fmxcZC.Show
    Call fmxcZC.Qing
    FmxcZcBr.Show 0
Case "财务到帐"
    If mod1.Qy = "上海" Then
        FMXCYBR.tt = "select top 50 khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy,mbf from htAView order by aid desc"
    Else
        FMXCYBR.tt = "select top 50 khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy,mbf from htAView where qy='" & mod1.Qy & "' order by aid desc"
    End If
    Call FMXCYBR.REF(FMXCYBR.tt)
    FMXCYBR.Show
Case "成本追加单"
    If mod1.DName = "邹晨" Or mod1.DName = "徐瑛" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
    frmZu.Enabled = True
    Exit Sub
    End If
    If mod1.Qy = "上海" Then
        tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc from htzuiView order by ztime desc,zid desc"
    Else
        tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc from htzuiView  where (qy='" & mod1.Qy & "'  order by ztime desc,zid desc"
    End If
    Call FmxcZuiBrow.Bound(tt)
    FmxcZuiBrow.tt = tt
    FmxcZuiBrow.Show
    FmxcZuiBrow.ZOrder 0
Case "销售管理"
    If mod1.DName = "邹晨" Or mod1.DName = "徐瑛" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
    frmZu.Enabled = True
    Exit Sub
    End If
    frmGzbN.Show
    frmGzbN.ZOrder 0
    frmGzbN.cmdXZ.Visible = True
Case "应收帐款"
    If mod1.DName = "邹晨" Or mod1.DName = "徐瑛" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
        frmZu.Enabled = True

    Exit Sub
    End If
        frmCWBB.cmdVnew.Visible = True
        frmCWBB.cmdV.Visible = False
            frmCWBB.comLx = "应收帐款"
        If mod1.KhK = 1 Then
            frmCWBB.comLx.Enabled = False
            frmCWBB.comLx = "应收帐款"
        Else
            frmCWBB.comLx.Enabled = True
        End If
            frmCWBB.Show: frmCWBB.ZOrder 0
Case "报销单" '
    If mod1.DName = "邹晨" Or mod1.DName = "徐瑛" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
        frmZu.Enabled = True

    Exit Sub
    End If
    mod1.BTZ = 23
    frmBxV.Visible = False
    frmBxV.mtA.Value = mod1.DQda
'    If mod1.Qy <> "上海" Then
        'tt = "select * from FydBrowG where qy='" & mod1.Qy & "' and (month(qrq)>= order by 签收日期 desc"
'    Else
'        tt = "Select * from FydBrowG order by 签收日期 desc"
'    End If
    tt = "FydVG('" & mod1.Qy & "','" & frmBxV.mtA.Value & "')"
    frmBxV.adoBxV.Close
    frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set frmBxV.dtgBx.DataSource = frmBxV.adoBxV
    If frmBxV.adoBxV.RecordCount > 0 Then
        frmBxV.dtgBx.FixedRows = 0
        frmBxV.dtgBx.MergeCol(1) = True
        frmBxV.dtgBx.MergeCol(2) = True
        frmBxV.dtgBx.MergeCol(3) = True
        frmBxV.dtgBx.MergeCol(4) = True
        frmBxV.dtgBx.MergeCol(5) = True
        frmBxV.dtgBx.MergeCol(7) = True
        frmBxV.dtgBx.MergeCells = 3
        frmBxV.dtgBx.FixedRows = 1
    End If
    frmBxV.Show
    frmBxV.ZOrder 0
Case "合同评审"
    If mod1.DName = "邹晨" Or mod1.DName = "徐瑛" Then
    MsgBox "没有权限!"
        frmZu.Enabled = True

    cmdFocus.SetFocus
    Exit Sub
    End If
'''''    htBrowG.Visible = False
'''''
'''''        tt = "Select * from htView1 where 项目归属人='马晓聪' order by 合同日期 desc"
'''''
'''''    htBrowG.adoBr.Close
'''''    htBrowG.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''    Set htBrowG.dtgBr.DataSource = htBrowG.adoBr
'''''    If htBrowG.adoBr.RecordCount > 0 Then
'''''        htBrowG.dtgBr.FixedRows = 0
'''''        htBrowG.dtgBr.MergeCol(1) = True
'''''        htBrowG.dtgBr.MergeCol(2) = True
'''''        htBrowG.dtgBr.MergeCol(3) = True
'''''        htBrowG.dtgBr.MergeCol(4) = True
'''''        htBrowG.dtgBr.MergeCol(7) = True
'''''        htBrowG.dtgBr.MergeCol(13) = True
'''''        htBrowG.dtgBr.MergeCells = 3
'''''        htBrowG.dtgBr.FixedRows = 1
'''''    End If

        htBrowG.lblFw.Caption = "业务部"


    htBrowG.Visible = True
    htBrowG.ZOrder 0
    
Case "合同执行"
    If mod1.DName = "邹晨" Or mod1.DName = "徐瑛" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
        frmZu.Enabled = True

    Exit Sub
    End If
    tt = "SELECT dbo.htzx.zid, dbo.htPing.htBh, dbo.htPing.khMc, dbo.htzx.xz, dbo.htzx.BZ, dbo.htzx.zTime, SUM(dbo.htzFk.amount) AS amount, dbo.htzx.ZT" & _
        " FROM dbo.htPing INNER JOIN dbo.htzx ON dbo.htPing.Hid = dbo.htzx.hid INNER JOIN dbo.htzFk ON dbo.htzx.zid = dbo.htzFk.zid" & _
        " Where (dbo.htzFk.Pwf = 1) GROUP BY dbo.htzx.zid, dbo.htPing.htBh, dbo.htPing.khMc, dbo.htzx.xz, dbo.htzx.BZ, dbo.htzx.zTime, dbo.htzx.ZT" & _
        " order by dbo.htzx.ztime desc"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        ua = UBound(Ra, 2) + 1
        Call frmHtZX.Bound(Ra, ua)
    frmHtZX.Show
    frmHtZX.ZOrder 0

Case Else
    MsgBox "正在建设中!3"
    frmZu.Enabled = True
    Call frmZu.OLOn
End Select
End Sub

Private Sub cmdBG_Click(Index As Integer)
Dim tt As String
Dim ii As Integer
Dim oo As Integer

Dim FR As Date

Dim FHg As Double
Dim Obmid As Integer
Dim Ra: Dim ua As Long: Dim Rb: Dim ub
Dim oClo


On Error Resume Next

If cmdBG(Index).ToolTipText = "没有权限" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
    Exit Sub
End If


frmZu.Enabled = False
'MsgBox cmdBu(Index).Caption
Call frmZu.OLOff
Select Case Trim(cmdBG(Index).Caption)
Case "成本追加单"
    If mod1.Qy = "上海" Then
        tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc,ztime from htzuiView order by ztime desc,zid desc"
    Else
        tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc,ztime from htzuiView  where (qy='" & mod1.Qy & "' or ggl='" & mod1.DHid & "') order by ztime desc,zid desc"
    End If
    Call FmxcZuiBrow.Bound(tt)
    FmxcZuiBrow.tt = tt
    FmxcZuiBrow.Show
    FmxcZuiBrow.ZOrder 0
Case "本月开单"
    frmHTKD.dtpM.Value = mod1.DQda
    Call frmHTKD.Initialize
    frmHTKD.Show
    frmHTKD.ZOrder 0
Case "财务评定"
    FmxcCw.txtHtbh.Text = ""
    Call FmxcCw.HtInitialize
    Call FmxcCw.Initialize
    FmxcCw.Show
    FmxcCw.ZOrder 0
Case "固定资产"
    frmComputer.Show
    Call frmComputer.OlineF
    frmComputer.ZOrder 0
Case "项目资料" '项目资料
    mod1.BTZ = 1
    frmKhbrG.Visible = False
    frmKhbrG.Left = 0
    frmKhbrG.Top = 0
    frmKhbrG.ZOrder 0
    frmKhbrG.Enabled = True
    frmKhbrG.ZOrder 0
    Set frmKhbrG.adoRenBr = CreateObject("adodb.recordset")
    Set frmKhbrG.adoKhBr = CreateObject("adodb.recordset")
    If mod1.KhK = 1 Then
        tt = "Select * from XmView where ggl='" & mod1.DHid & "' order by 业务员"
    ElseIf mod1.KhK = 3 Then
        tt = "Select * from xmView  order by comid,部门,业务员"
    ElseIf mod1.KhK = 2 And mod1.comId <> 0 Then
        tt = "Select * from xmView where comid=" & mod1.comId & " order by 部门,业务员"
    ElseIf mod1.KhK = 2 And mod1.comId = 0 Then '倪旭
        tt = "Select * from xmView where comid=" & mod1.comId & " and not(部门='维销部3' or 部门='产品部1' or 部门='产品部2')  order by 部门,业务员"
    End If
    frmKhbrG.adoKhBr.Close
    frmKhbrG.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmKhbrG.dtgKh.DataSource = frmKhbrG.adoKhBr
    frmKhbrG.dtgKh.FixedRows = 0
    frmKhbrG.dtgKh.MergeCol(4) = True
    frmKhbrG.dtgKh.MergeCol(12) = True
    frmKhbrG.dtgKh.MergeCol(14) = True
    frmKhbrG.dtgKh.MergeCells = 3
    frmKhbrG.dtgKh.FixedRows = 1
    frmKhbrG.tabCx.Tab = 0
    If mod1.KhK = 1 Then
        frmKhbrG.lblFw.Caption = mod1.Bm
    Else
        frmKhbrG.lblFw.Caption = "业务部"
    End If
    frmKhbrG.Visible = True
    frmKhbrG.lblYwy.Caption = ""
'    '设置业务员下拉框
'    If mod1.KhK = 1 Then
'        tt = "select username,userid from DlName where bm='" & mod1.Bm & "'"
'    ElseIf mod1.KhK = 2 Then
'        tt = "select username,userid from workerNew where kqf=1 order by bm desc"
'    End If
'    Set frmKhBrg. = CreateObject("adodb.recordset")
'    frmKhBrg.adoYwy.Close
'    frmKhBrg.adoYwy.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    Set frmKhBrg.comYwy.RowSource = frmKhBrg.adoYwy
'    frmKhBrg.comYwy.ListField = "username"
'    frmKhBrg.comYwy.BoundColumn = "userid"
Case "工作报告"
    frmGzbN.Show
    frmGzbN.ZOrder 0
    frmGzbN.cmdXZ.Visible = True
'''''''''    mod1.BTZ = 4
'''''''''    frmGzBG.Visible = False
'''''''''
'''''''''
'''''''''    Dim FR As Date '一周的起始日期（星期一）
'''''''''    Dim LR As Date '一周的截至日期（星期二）
'''''''''
'''''''''
'''''''''    '设置默认日期和星期
'''''''''    frmGzBG.dtpRq.Value = mod1.HMDa
'''''''''    frmGzBG.lblWeek.Caption = modXmGz.dayWeek(frmGzBG.dtpRq.DayOfWeek)
'''''''''
'''''''''
'''''''''    '打开工作报告表(先取得当前周日期)
'''''''''
'''''''''    Select Case frmGzBG.dtpRq.DayOfWeek
'''''''''    Case 1 '星期日
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 6)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa))
'''''''''
'''''''''    Case 2 '星期一
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa))
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 6)
'''''''''    Case 3
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 1)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 5)
'''''''''    Case 4
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 2)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 4)
'''''''''    Case 5
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 3)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 3)
'''''''''    Case 6
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 4)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 2)
'''''''''    Case 7
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 5)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 1)
'''''''''    End Select
'''''''''
'''''''''    frmGzBG.lblFr.Caption = FR
'''''''''    frmGzBG.lblLR.Caption = LR
'''''''''    modXmGz.FR = FR
'''''''''    modXmGz.LR = LR
'''''''''
'''''''''    tt = "Select atime,khqc,xmfy,NewF,gid from xmgz where ywy='" & mod1.DName & "' and aTime>='" & FR & _
'''''''''    "' and aTime <='" & LR & "' and lb=1 order by aTime"
'''''''''    frmGzBG.adoXm.Close
'''''''''    frmGzBG.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''
'''''''''    Set frmGzBG.dtgXmgz.DataSource = frmGzBG.adoXm
'''''''''
'''''''''    tt = "Select atime,khqc,newF,gid from xmgz where ywy='" & mod1.DName & "' and aTime>='" & FR & _
'''''''''    "' and aTime <='" & LR & "' and lb=0 order by aTime"
'''''''''    frmGzBG.adoJi.Close
'''''''''    frmGzBG.adoJi.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''    Set frmGzBG.dtgJi.DataSource = frmGzBG.adoJi
'''''''''    frmGzBG.dtpXr.Value = DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda))
'''''''''
''''''''''     tt = "select xmmc as 项目名称,khjb as 项目平台,xid as 编号,xmfy as 费用,xid from xmzl where ywy='" & mod1.DName & "'"
''''''''''    Set frmGzBG.AdoKh = CreateObject("adodb.recordset")
''''''''''    frmGzBG.AdoKh.Close
''''''''''    frmGzBG.AdoKh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''    Set frmGzBG.dtgKH.DataSource = frmGzBG.AdoKh
'''''''''    frmGzBG.lblYwy.Caption = mod1.DName
'''''''''
'''''''''    '取得销售日记流程参数
'''''''''    tt = "ZBut('" & mod1.DName & "','" & mod1.DHid & "','销售日记')"
'''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'''''''''    frmGzBG.cmdOK.Tag = mod1.HTP.Fields("nlb").Value
'''''''''    frmGzBG.cmdOK.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
'''''''''    frmGzBG.Show
'''''''''    frmGzBG.comKhmc.Text = ""
'''''''''    Set frmGzBG.dtgHt.DataSource = Nothing
'''''''''    frmGzBG.lblHtbh.Caption = ""
'''''''''    frmGzBG.cmdFw.Visible = True
'''''''''    frmGzBG.lblFw.Visible = True
Case "合同评审"
    
    FR = DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda) - 14)
    htBrowG.Visible = False
    If mod1.Qy <> "北京" Then
            If mod1.KhK = 1 Then
        
                tt = "Select 项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid from htView1 where ggl='" & mod1.DHid & "' and htrq>='" & FR & "' and 状态<>'评审'  order by htrq desc"
            
            ElseIf mod1.KhK = 2 Then
                tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where not(部门='维销部3' or 部门='产品部1' or 部门='产品部2')  and htrq>='" & FR & "' and 状态<>'评审' order by htrq desc"
            ElseIf mod1.KhK = 3 Then
                tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where htrq>='" & FR & "' and 状态<>'评审' order by htrq desc"
            End If
        '''''    If mod1.BM = "工程二部" Then
        '''''        tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0   from htViewP where 状态<>'评审'   order by 部门,htrq desc"
        '''''    End If
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
            Ra = mod1.HTP.GetRows
            mod1.HTP.Close
            Set mod1.HTP = Nothing
            ua = UBound(Ra, 2)
            Call htBrowG.Bref(Ra, ua + 1)
            If mod1.KhK = 1 Then
                htBrowG.lblFw.Caption = mod1.Bm
            Else
                htBrowG.lblFw.Caption = "业务部"
            End If
    ElseIf mod1.Qy = "北京" Then

                tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where htrq>='" & FR & "' and 状态<>'评审' and 区域='北京' order by htrq desc"

            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
            Ra = mod1.HTP.GetRows
            mod1.HTP.Close
            Set mod1.HTP = Nothing
            ua = UBound(Ra, 2)
            Call htBrowG.Bref(Ra, ua + 1)
            If mod1.KhK = 1 Then
                htBrowG.lblFw.Caption = mod1.Bm
            Else
                htBrowG.lblFw.Caption = "北京项目1部"
            End If
    End If

    htBrowG.Visible = True
    htBrowG.ZOrder 0
Case "合同执行"

    tt = "SELECT dbo.htzx.zid, dbo.htPing.htBh, dbo.htPing.khMc, dbo.htzx.xz, dbo.htzx.BZ, dbo.htzx.zTime, SUM(dbo.htzFk.amount) AS amount, dbo.htzx.ZT" & _
        " FROM dbo.htPing INNER JOIN dbo.htzx ON dbo.htPing.Hid = dbo.htzx.hid INNER JOIN dbo.htzFk ON dbo.htzx.zid = dbo.htzFk.zid" & _
        " Where (dbo.htzFk.Pwf = 1) GROUP BY dbo.htzx.zid, dbo.htPing.htBh, dbo.htPing.khMc, dbo.htzx.xz, dbo.htzx.BZ, dbo.htzx.zTime, dbo.htzx.ZT" & _
        " order by dbo.htzx.ztime desc"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        ua = UBound(Ra, 2) + 1
        Call frmHtZX.Bound(Ra, ua)
    frmHtZX.Show
    frmHtZX.ZOrder 0
Case "库存报表"
    KCBB.Show
Case "维保记录"
    frmGZDBR.Show
    frmGZDBR.frmFw.Visible = True
Case "报价单"
    Set frmBJD.adoBJD = CreateObject("adodb.recordset")
    tt = "select * from bjdV where baoid=-1"
    frmBJD.adoBJD.Close
    frmBJD.adoBJD.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmBJD.dtgBJD.DataSource = frmBJD.adoBJD
    frmBJD.Show
    frmBJD.ZOrder 0
Case "施工计划"
    If mod1.KhK = 1 Then
        tt = "select * from GjwV1 where bm='" & mod1.Bm & "' order by gid desc"
    ElseIf mod1.KhK = 2 Then
        tt = "select * from GjwV1 where comid=" & mod1.comId & " order by gid desc"
    ElseIf mod1.KhK = 3 Then
        tt = "select * from GjwV1 where comid=" & mod1.comId & " order by gid desc"
    ElseIf mod1.DName = "金强" Then
        tt = "select * from GjwV1 where comid=" & mod1.comId & " order by gid desc"
    End If
    Set frmGjwV.adoGJW = CreateObject("adodb.recordset")
    frmGjwV.adoGJW.Close
    frmGjwV.adoGJW.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGjwV.dtgBr.DataSource = frmGjwV.adoGJW
    If frmGjwV.adoGJW.RecordCount > 0 Then
        frmGjwV.dtgBr.FixedRows = 0
        frmGjwV.dtgBr.MergeCol(2) = True
        frmGjwV.dtgBr.MergeCol(3) = True
        frmGjwV.dtgBr.MergeCol(4) = True
        frmGjwV.dtgBr.MergeCells = 3
        frmGjwV.dtgBr.FixedRows = 1
    End If
    frmGjwV.Visible = True
    frmGjwV.Enabled = True
    frmGjwV.ZOrder 0
Case "费用报销" '
    If mod1.DName = "孟智峰" Then
        fyBB.Show
        fyBB.ZOrder 0
        Me.Enabled = False
    End If
    mod1.BTZ = 23
    frmBxV.Visible = False
    frmBxV.mtA.Value = mod1.DQda
'    If mod1.Qy <> "上海" Then
        'tt = "select * from FydBrowG where qy='" & mod1.Qy & "' and (month(qrq)>= order by 签收日期 desc"
'    Else
'        tt = "Select * from FydBrowG order by 签收日期 desc"
'    End If
    tt = "FydVG('" & mod1.Qy & "','" & frmBxV.mtA.Value & "')"
    frmBxV.adoBxV.Close
    frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set frmBxV.dtgBx.DataSource = frmBxV.adoBxV
    If frmBxV.adoBxV.RecordCount > 0 Then
        frmBxV.dtgBx.FixedRows = 0
        frmBxV.dtgBx.MergeCol(1) = True
        frmBxV.dtgBx.MergeCol(2) = True
        frmBxV.dtgBx.MergeCol(3) = True
        frmBxV.dtgBx.MergeCol(4) = True
        frmBxV.dtgBx.MergeCol(5) = True
        frmBxV.dtgBx.MergeCol(7) = True
        frmBxV.dtgBx.MergeCells = 3
        frmBxV.dtgBx.FixedRows = 1
    End If
    frmBxV.Show
    frmBxV.ZOrder 0
Case "维保执行表"
    frmWBjl.lblZNAME.Caption = ""
    frmWBjl.comLx.Text = "项目名称"
    frmWBjl.txtZ.Text = ""
    frmWBjl.cmdSave.Enabled = False
    frmWBjl.frmMod.Visible = False
    tt = "select * from wbjl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "' and yy=" & Year(mod1.DQda) & " order by jid"
    frmWBjl.adoNr.Recordset.Close
    If mod1.Zuf = True Then
        frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        frmWBjl.cmdMod.Enabled = True
    Else
        frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmWBjl.cmdMod.Enabled = False
    End If
    Set frmWBjl.dtpNr.DataSource = frmWBjl.adoNr
    If mod1.Zuf = False Or mod1.DName = "张寅" Then
        frmWBjl.frmRen.Visible = True
        frmWBjl.cmdMod.Enabled = False
    Else
        frmWBjl.frmRen.Visible = False
        frmWBjl.lblZNAME = mod1.DName
        frmWBjl.cmdMod.Enabled = True
    End If
    frmWBjl.Show
    frmWBjl.ZOrder 0
Case "财务报表"
    If mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1" Then
        ii = MsgBox("是否要转到2008新财务报表？", vbYesNo + vbQuestion, "请选择")
        If ii = vbYes Then
            frmCWBB.Show
            frmCWBB.ZOrder 0
            Me.Enabled = False
            Exit Sub
        End If
    End If
    fyBB.Show
    fyBB.ZOrder 0
    Me.Enabled = False
Case "员工档案"
    frmRen.Visible = True
    If mod1.Qy = "上海" Then
        tt = "select userid as 工号,username as 姓名,qy as 区域,bm as 部门,userzw as 职务,nx as 工作年限 from worker where zzf=1 order by userid"
    Else
        tt = "select userid as 工号,username as 姓名,qy as 区域,bm as 部门,userzw as 职务,nx as 工作年限 from worker where zzf=1 and qy='" & mod1.Qy & "' order by userid"
    End If
    frmRen.adoRen.Close
    frmRen.adoRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmRen.dtgRen.DataSource = frmRen.adoRen
    Call frmRen.RenQing
    frmRen.frmMod.Enabled = False
    frmRen.txtTang.Visible = True
Case "应收帐款"
        frmCWBB.cmdVnew.Visible = True
        frmCWBB.cmdV.Visible = False
            frmCWBB.comLx = "应收帐款"
        If mod1.KhK = 1 Then
            frmCWBB.comLx.Enabled = False
            frmCWBB.comLx = "应收帐款"
        Else
            frmCWBB.comLx.Enabled = True
        End If
            frmCWBB.Show: frmCWBB.ZOrder 0
Case Else
    MsgBox "正在建设中!3"
    frmZu.Enabled = True
    Call frmZu.OLOn
End Select

End Sub



















Private Sub cmdBi_Click()
Dim tt As String
Dim oo As Integer
Dim Mon As Integer
Dim Bxid As Long
On Error Resume Next

'tt = "select * from worker where bird=" & Month(mod1.DQda)
'mod1.HTT.Close
'mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Mon = Month(mod1.DQda)
'
'mod1.HTT.MoveFirst
'Do While Not mod1.HTT.EOF
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "FydAddBird"
    mod1.cmd.CommandType = adCmdStoredProc
    'mod1.CMD.Parameters("@qy") = mod1.HTT.Fields("qy").Value
    'mod1.CMD.Parameters("@bm") = mod1.HTT.Fields("bm").Value
    mod1.cmd.Parameters("@qy") = "南京"
    mod1.cmd.Parameters("@bm") = "南京办"
    mod1.cmd.Parameters("@trq") = mod1.DQda
    mod1.cmd.Parameters("@ywy") = "李俊"
    mod1.cmd.Parameters("@uid") = "HM011"
    mod1.cmd.Parameters("@Lcou") = 3 '流程总数
    mod1.cmd.Parameters("@Lc") = 1 '当前流程
    mod1.cmd.Parameters("@frq") = DateSerial(Year(mod1.DQda), 1, 1)
    mod1.cmd.Parameters("@lrq") = DateSerial(Year(mod1.DQda), 12, 31)
'    mod1.CMD.Parameters("@lcRen") = mod1.HTT.Fields("username").Value
'    mod1.CMD.Parameters("@lcUid") = mod1.HTT.Fields("userid").Value
    mod1.cmd.Parameters("@lcRen") = "李俊"
    mod1.cmd.Parameters("@lcUid") = "HM011"
    mod1.cmd.Parameters("@nlb") = 66

    mod1.cmd.Parameters("@fbt") = "生日报销单"
    mod1.cmd.Parameters("@sj") = 100
    mod1.cmd.Parameters("@hg") = 100
    mod1.cmd.Parameters("@hGd") = mod1.ChangBi(100)
    mod1.cmd.Parameters("@Lb") = 66
    mod1.cmd.Execute
    Bxid = mod1.cmd.Parameters("@bxid").Value
    Set cmd = Nothing
    
    '添加事务
    Call mod1.EnventAdd("报销单", 100, mod1.HTT.Fields("username").Value, mod1.HTT.Fields("userid").Value, Str(Bxid), "报销人", "", "", mod1.HTT.Fields("username").Value, mod1.HTT.Fields("userid").Value, 0, Str(Bxid))
    
    '设置流程按钮
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "QMRZAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@NLb") = 66
    mod1.cmd.Parameters("@btz") = 23
    mod1.cmd.Parameters("@QDBH") = Bxid '编号
    mod1.cmd.Execute
    Set cmd = Nothing
'    mod1.HTT.MoveNext
'Loop
End Sub

Private Sub cmdBu_Click(Index As Integer)
Dim tt As String
Dim oo As Integer
Dim jj As Integer
Dim YY As Integer
Dim Ra: Dim La
On Error Resume Next
If cmdBu(Index).ToolTipText = "没有权限" Then
    MsgBox "没有权限!"
    cmdFocus.SetFocus
    Exit Sub
End If


'MsgBox cmdBu(Index).Caption
cmdBack.SetFocus
Call frmZu.OLOff
Select Case Trim(cmdBu(Index).Caption)
Case "执行报表"
    fmxcZfile.Show
    fmxcZfile.Visible = True
    fmxcZfile.ZOrder 0
    Call fmxcZfile.Bound
Case "出工统计"
    frmWork.Show
    frmWork.Visible = True
    frmWork.ZOrder 0
Case "项目资料" '项目资料
    mod1.BTZ = 1
    frmKhBr.Visible = False
    frmKhBr.Left = 0
    frmKhBr.Top = 0
    frmKhBr.ZOrder 0
    frmKhBr.Enabled = True
    frmKhBr.ZOrder 0
    Set frmKhBr.adoKhBr = CreateObject("adodb.recordset")
    Set frmKhBr.adoRenBr = CreateObject("adodb.recordset")
    tt = "vXmNew('" & mod1.DName & "','" & mod1.DHid & "')"
    frmKhBr.adoKhBr.Close
    frmKhBr.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set frmKhBr.dtgKh.DataSource = frmKhBr.adoKhBr
    frmKhBr.tabCx.Tab = 0
    frmKhBr.Visible = True
'    tt = "VkhrNew('" & mod1.DName & "')"
'    frmKhBr.adoRenBr.Close
'    frmKhBr.adoRenBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'    Set frmKhBr.dtgLx.DataSource = frmKhBr.adoRenBr
    '取得新建客户联系人流程参数
    tt = "khBut('" & mod1.DName & "','" & mod1.DHid & "')"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    frmKhBr.cmdNew.Tag = mod1.HTP.Fields("nlb").Value
    frmKhBr.cmdNew.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
Case "项目跟踪"
    mod1.BTZ = 4
    frmGzbN.Show
    frmGzbN.dtgB.Row = 1: frmGzbN.dtgB.Col = 0
    If frmGzbN.dtgB.Text = "" Then
        frmGzbN.dtgB.Visible = False
        frmGzbN.lblRen.ToolTipText = mod1.DHid   '业务员打开工作报告,默认为本人
        frmGzbN.lblRen.Caption = mod1.DName
        Call frmGzbN.WeekDate(mod1.DQda, mod1.DHid)
      
        Call frmGzbN.QV(frmGzbN.FS)
    
        frmGzbN.dtgB.Visible = True
        frmGzbN.dtgB.Row = 1
        frmGzbN.cmdXZ.Visible = False
    
    End If

'''''''''    mod1.BTZ = 4
'''''''''    frmGzBG.Visible = False
'''''''''
'''''''''
'''''''''    Dim FR As Date '一周的起始日期（星期一）
'''''''''    Dim LR As Date '一周的截至日期（星期二）
'''''''''
'''''''''
'''''''''    '设置默认日期和星期
'''''''''    frmGzBG.dtpRq.Value = mod1.HMDa
'''''''''    frmGzBG.lblWeek.Caption = modXmGz.dayWeek(frmGzBG.dtpRq.DayOfWeek)
'''''''''
'''''''''
'''''''''    '打开工作报告表(先取得当前周日期)
'''''''''
'''''''''    Select Case frmGzBG.dtpRq.DayOfWeek
'''''''''    Case 1 '星期日
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 6)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa))
'''''''''
'''''''''    Case 2 '星期一
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa))
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 6)
'''''''''    Case 3
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 1)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 5)
'''''''''    Case 4
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 2)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 4)
'''''''''    Case 5
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 3)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 3)
'''''''''    Case 6
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 4)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 2)
'''''''''    Case 7
'''''''''    FR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) - 5)
'''''''''    LR = DateSerial(Year(mod1.HMDa), Month(mod1.HMDa), Day(mod1.HMDa) + 1)
'''''''''    End Select
'''''''''
'''''''''    frmGzBG.lblFr.Caption = FR
'''''''''    frmGzBG.lblLR.Caption = LR
'''''''''    modXmGz.FR = FR
'''''''''    modXmGz.LR = LR
'''''''''
'''''''''    tt = "Select atime,khqc,xmfy,NewF,gid from xmgz where ywy='" & mod1.DName & "' and aTime>='" & FR & _
'''''''''    "' and aTime <='" & LR & "' and lb=1 order by aTime"
'''''''''    frmGzBG.adoXm.Close
'''''''''    frmGzBG.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''
'''''''''    Set frmGzBG.dtgXmgz.DataSource = frmGzBG.adoXm
'''''''''
'''''''''    tt = "Select atime,xmmc,newF,gid from xmgz where ywy='" & mod1.DName & "' and aTime>='" & FR & _
'''''''''    "' and aTime <='" & LR & "' and lb=0 order by aTime"
'''''''''    frmGzBG.adoJi.Close
'''''''''    frmGzBG.adoJi.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''    Set frmGzBG.dtgJi.DataSource = frmGzBG.adoJi
'''''''''    frmGzBG.dtpXr.Value = DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda))
'''''''''
''''''''''     tt = "select xmmc as 项目名称,khjb as 项目平台,xid as 编号,xmfy as 费用,xid from xmzl where ywy='" & mod1.DName & "'"
''''''''''    Set frmGzBG.AdoKh = CreateObject("adodb.recordset")
''''''''''    frmGzBG.AdoKh.Close
''''''''''    frmGzBG.AdoKh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''    Set frmGzBG.dtgKH.DataSource = frmGzBG.AdoKh
'''''''''    frmGzBG.lblFw.Caption = mod1.DName
'''''''''    frmGzBG.lblYwy.Caption = mod1.DName
'''''''''
'''''''''    '取得销售日记流程参数
'''''''''    tt = "ZBut('" & mod1.DName & "','" & mod1.DHid & "','销售日记')"
'''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'''''''''    frmGzBG.cmdOK.Tag = mod1.HTP.Fields("nlb").Value
'''''''''    frmGzBG.cmdOK.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
'''''''''    frmGzBG.Show
'''''''''    frmGzBG.comKhmc.Text = ""
'''''''''    Set frmGzBG.dtgHt.DataSource = Nothing
'''''''''    frmGzBG.lblHtbh.Caption = ""
'''''''''    frmGzBG.cmdFw.Visible = False
'''''''''    frmGzBG.lblFw.Visible = False
Case "合同评审"
    'frmGxNew.Show
    
    htBrow.Visible = False
    htBrow.timWait.Enabled = False
    htBrow.timQuit.Enabled = False
    htBrow.DT = "Select top 30 * from htView where (业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
    "') order by 合同日期 desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open htBrow.DT, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2)
    Call htBrow.dtgGG
    htBrow.dtgBr.Rows = La + 30
    htBrow.dtgN.Rows = La + 30: htBrow.dtgN.Cols = htBrow.dtgBr.Cols
    For oo = 1 To La + 1
        htBrow.dtgBr.Row = oo: htBrow.dtgN.Row = oo
        For ii = 1 To htBrow.dtgBr.Cols
            htBrow.dtgBr.Col = ii: htBrow.dtgN.Col = ii
            If ii = 3 Then '日期格式化
                htBrow.dtgBr.Text = Format(Ra(ii - 1, oo - 1), "YYYY-MM-DD")
                htBrow.dtgN.Text = Format(Ra(ii - 1, oo - 1), "YYYY-MM-DD")
            Else
                htBrow.dtgBr.Text = Ra(ii - 1, oo - 1)
                htBrow.dtgN.Text = Ra(ii - 1, oo - 1)
            End If
            If ii = 17 And Val(htBrow.dtgBr.Text) > 0 Then
                For YY = 1 To htBrow.dtgBr.Cols
                    htBrow.dtgBr.Col = YY
                    htBrow.dtgBr.CellForeColor = &H8000000D
                Next
            End If
            htBrow.dtgBr.Col = ii
        Next
    Next
'''''''    htBrow.adoBr.Close
'''''''    htBrow.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''    Set htBrow.dtgBr.DataSource = htBrow.adoBr
'''''''    htBrow.dtgBr.FixedRows = 0
'''''''    htBrow.dtgBr.MergeCol(1) = True
'''''''    htBrow.dtgBr.MergeCol(2) = True
'''''''    htBrow.dtgBr.MergeCol(3) = True
'''''''    htBrow.dtgBr.MergeCol(4) = True
'''''''    htBrow.dtgBr.MergeCol(7) = True
'''''''    htBrow.dtgBr.MergeCells = 3
'''''''    htBrow.dtgBr.FixedRows = 1
    If mod1.ZT = "HMData" Then '上海的帐套有新建新版合同的功能
        'htBrow.NF.Visible = True
        
    Else
        'htBrow.NF.Visible = False
    End If
    htBrow.Show
    htBrow.optY(0).Value = True
    htBrow.ZOrder 0
'    frmWbNew.Show
'    frmWbNew.frmWb.Visible = True
Case "询价"
Exit Sub
Me.Enabled = True
    mod1.BTZ = 36
    frmGxBiao.Visible = False
    tt = "select * from xunjiaView where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "' order by bid desc"
    frmGxBiao.adoXj.Close
    frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
    If frmGxBiao.adoXj.RecordCount > 1 Then
        frmGxBiao.dtgXj.FixedRows = 0
        frmGxBiao.dtgXj.MergeCol(1) = True
        frmGxBiao.dtgXj.MergeCol(3) = True
        frmGxBiao.dtgXj.MergeCol(4) = True
        frmGxBiao.dtgXj.MergeCells = 3
        frmGxBiao.dtgXj.FixedRows = 1
    End If
    frmGxBiao.Visible = True
    '取得新建维保询价单及购销询价单的流程参数
    tt = "xunJiaBut('" & mod1.DName & "','" & mod1.DHid & "','维保询价')"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    frmGxBiao.cmdNew.Tag = mod1.HTP.Fields("nlb").Value
    frmGxBiao.cmdNew.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
    frmGxBiao.cmdDx.Tag = mod1.HTP.Fields("nlb").Value                          '大修的流程同维保
    frmGxBiao.cmdDx.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
    tt = "xunJiaBut('" & mod1.DName & "','" & mod1.DHid & "','购销')"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    frmGxBiao.cmdCreat.Tag = mod1.HTP.Fields("nlb").Value
    frmGxBiao.cmdCreat.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
    frmGxBiao.cmdCP.Tag = mod1.HTP.Fields("nlb").Value
    frmGxBiao.cmdCP.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
    frmGxBiao.cmdFb.Tag = mod1.HTP.Fields("nlb").Value
    frmGxBiao.cmdFb.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
    frmGxBiao.frmNew.Visible = True
    frmGxBiao.frmC.Visible = False

    
''''''''''    '显示工程部询价表
''''''''''    tt = "select * from xunjiagcv where 业务员='" & mod1.DName & "' and yuid='" & mod1.DHid & "' order by 询价日期 desc"
''''''''''    frmGxBiao.adoGc.Close
''''''''''    frmGxBiao.adoGc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''    Set frmGxBiao.dtgGc.DataSource = frmGxBiao.adoGc
''''''''''    If frmGxBiao.adoGc.RecordCount > 1 Then
''''''''''        frmGxBiao.dtgGc.FixedRows = 0
''''''''''        frmGxBiao.dtgGc.MergeCol(1) = True
''''''''''        frmGxBiao.dtgGc.MergeCol(3) = True
''''''''''        frmGxBiao.dtgGc.MergeCol(4) = True
''''''''''        frmGxBiao.dtgGc.MergeCells = 3
''''''''''        frmGxBiao.dtgGc.FixedRows = 1
''''''''''    End If
Case "满意度调查"
    Exit Sub
    frmZu.Enabled = True
    Me.MousePointer = 11
    Set mod1.report = mod1.crapp.OpenReport(App.Path & "\myd.rpt")
     'Set mod1.report = mod1.crapp.OpenReport(App.Path & "\tt.rpt")
    Set mod1.table = mod1.report.Database.Tables
    Set mod1.cProp = mod1.table.Item(1).ConnectionProperties
    mod1.cProp.Item("Password") = "guyonghui"
    mod1.report.SQLQueryString = "Select xmmc,myd,jxrq from gzbN  "
    mod1.report.ReadRecords
    frmReport.Show
    frmReport.cR1.ReportSource = mod1.report
    frmReport.cR1.ViewReport
    'frmReport.cR1.EnableExportButton = True

    Me.MousePointer = 0
    frmReport.cR1.EnableExportButton = False
    frmReport.cR1.EnableExportButton = True
Case "项目执行"
    'frmGxNew.Show
    frmHtZX.Visible = False
'''''    'tt = "Select * from htView where (业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
'''''    "') order by 合同日期 desc"
'''''    tt = "select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where ((业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & _
'''''        mod1.DName & "' and Xuid='" & mod1.DHid & "')) and 状态='执行' order by 合同日期 desc"
'''''    frmHtZX.adoBr.Close
'''''    frmHtZX.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''    Set frmHtZX.dtgBr.DataSource = frmHtZX.adoBr
'''''    If frmHtZX.adoBr.RecordCount > 0 Then
'''''        frmHtZX.dtgBr.FixedRows = 0
'''''        frmHtZX.dtgBr.MergeCol(1) = True
'''''        frmHtZX.dtgBr.MergeCol(2) = True
'''''        frmHtZX.dtgBr.MergeCol(3) = True
'''''        frmHtZX.dtgBr.MergeCells = 3
'''''        frmHtZX.dtgBr.FixedRows = 1
'''''    End If
'''''
'''''    frmHtZX.lblHtbh.Caption = ""
'''''    tt = "select * from PldView where htbh='" & frmHtZX.lblHtbh.Caption & "' order by 编号"
'''''    frmHtZX.adoPld.Close
'''''    frmHtZX.adoPld.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''    Set frmHtZX.dtgPld.DataSource = frmHtZX.adoPld
    frmHtZX.Show
    'frmHtZX.optY(0).Value = True
    frmHtZX.ZOrder 0
Case "工作单"
    frmGZDBR.Show
    frmGZDBR.ZOrder 0
    tt = "Select 检验日期,编号,工作单类型,业务员,gid,uid,qy,trq,项目名称,日期,fl,合格否  from gzdView where trq>='" & DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda)) & " ' and 业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "' order by gid desc"
    frmGZDBR.adoY.Close
    frmGZDBR.adoY.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGZDBR.dtgY.DataSource = frmGZDBR.adoY
    tt = "select 检验日期,编号,工作单类型,gid,fl FROM gzdView where trq is null and 业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "' order by 检验日期"
    frmGZDBR.adoW.Close
    frmGZDBR.adoW.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGZDBR.dtgW.DataSource = frmGZDBR.adoW

    frmGZDBR.frmFw.Visible = False
    frmGZDBR.lblFw.Caption = mod1.DName
    frmGZDBR.lblFw.ToolTipText = mod1.DHid
Case "报价单"
    Set frmBJD.adoBJD = CreateObject("adodb.recordset")
    tt = "select * from bjdV where 业务员='" & mod1.DName & "'"
    frmBJD.adoBJD.Close
    frmBJD.adoBJD.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmBJD.dtgBJD.DataSource = frmBJD.adoBJD
    frmBJD.Show
    frmBJD.ZOrder 0
    frmBJD.cmdFw.Visible = False
    frmBJD.lblFw.Visible = False
    frmBJD.cmdAll.Visible = False
Case "施工计划"

    tt = "select * from GjwV where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"

    Set frmGjwV.adoGJW = CreateObject("adodb.recordset")
    frmGjwV.adoGJW.Close
    frmGjwV.adoGJW.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGjwV.dtgBr.DataSource = frmGjwV.adoGJW
    If frmGjwV.adoGJW.RecordCount > 0 Then
        frmGjwV.dtgBr.FixedRows = 0
        frmGjwV.dtgBr.MergeCol(2) = True
        frmGjwV.dtgBr.MergeCol(3) = True
        frmGjwV.dtgBr.MergeCol(4) = True
        frmGjwV.dtgBr.MergeCells = 3
        frmGjwV.dtgBr.FixedRows = 1
    End If
    frmGjwV.Visible = True
    frmGjwV.Enabled = True
    frmGjwV.ZOrder 0
Case Else
    MsgBox "正在建设中1"
    frmZu.Enabled = True
    Call frmZu.OLOn
End Select

End Sub






















Private Sub cmdCancel_Click()
frmTX.Visible = False
End Sub

Private Sub cmdFBGR_Click()
'''''''''Dim tt As String
'''''''''Dim oo As Integer
'''''''''Dim Mon As Integer
'''''''''Dim Bxid As Long
'''''''''Dim DName As String
'''''''''Dim Uid As String
'''''''''Dim Qy As String
'''''''''Dim BM As String
'''''''''On Error Resume Next
'''''''''DName = InputBox("请输入员工姓名!")
'''''''''tt = "select bm,qy,userid from worker where username='" & DName & "'"
'''''''''mod1.HTT.Close
'''''''''mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''BM = mod1.HTT.Fields("bm").Value
'''''''''Qy = mod1.HTT.Fields("qy").Value
'''''''''Uid = mod1.HTT.Fields("userid").Value
'''''''''Mon = InputBox("请输入月份!")
'''''''''
'''''''''
'''''''''    Set mod1.cmd = createobject("adodb.command")
'''''''''    mod1.cmd.ActiveConnection = mod1.CC
'''''''''    mod1.cmd.CommandText = "FydAddFwbt"
'''''''''    mod1.cmd.CommandType = adCmdStoredProc
'''''''''    mod1.cmd.Parameters("@qy") = Qy
'''''''''    mod1.cmd.Parameters("@bm") = BM
'''''''''    mod1.cmd.Parameters("@trq") = mod1.DQda
'''''''''    mod1.cmd.Parameters("@ywy") = DName
'''''''''    mod1.cmd.Parameters("@uid") = Uid
'''''''''    mod1.cmd.Parameters("@Lcou") = 3 '流程总数
'''''''''    mod1.cmd.Parameters("@Lc") = 1 '当前流程
'''''''''    mod1.cmd.Parameters("@frq") = DateSerial(Year(mod1.DQda), Month(mod1.DQda), 1)
'''''''''    mod1.cmd.Parameters("@lrq") = DateSerial(Year(mod1.DQda), Month(mod1.DQda), 28)
'''''''''    mod1.cmd.Parameters("@lcRen") = DName
'''''''''    mod1.cmd.Parameters("@lcUid") = Uid
'''''''''    mod1.cmd.Parameters("@nlb") = 67
'''''''''    mod1.cmd.Parameters("@mon") = Str(Mon)
'''''''''    mod1.cmd.Parameters("@fbt") = "房屋补贴报销单"
'''''''''    mod1.cmd.Parameters("@fwbt") = 100
'''''''''    mod1.cmd.Parameters("@hg") = 100
'''''''''    mod1.cmd.Parameters("@hGd") = mod1.ChangBi(100)
'''''''''    mod1.cmd.Parameters("@Lb") = 67
'''''''''    mod1.cmd.Execute
'''''''''    Bxid = mod1.cmd.Parameters("@bxid").Value
'''''''''    Set cmd = Nothing
'''''''''
'''''''''    '添加事务
'''''''''    Call mod1.EnventAdd("报销单", 100, DName, Uid, Str(Bxid), "报销人", "", "", DName, Uid, 0, Str(Bxid))
'''''''''
'''''''''    '设置流程按钮
'''''''''    Set mod1.cmd = createobject("adodb.command")
'''''''''    mod1.cmd.ActiveConnection = mod1.CC
'''''''''    mod1.cmd.CommandText = "QMRZAdd"
'''''''''    mod1.cmd.CommandType = adCmdStoredProc
'''''''''    mod1.cmd.Parameters("@NLb") = 67
'''''''''    mod1.cmd.Parameters("@btz") = 23
'''''''''    mod1.cmd.Parameters("@QDBH") = Bxid '编号
'''''''''    mod1.cmd.Execute
'''''''''    Set cmd = Nothing
Shell "explorer.exe http://10.128.123.10", vbHide
End Sub

Private Sub cmdFwbt_Click()
Dim tt As String
Dim oo As Integer
Dim Mon As Integer
Dim Bxid As Long
Dim Ddate As Date
On Error Resume Next
Ddate = "2007-9-1"
tt = "select * from fwbtView"
mod1.HTT.Close
mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Mon = Month(Ddate) - 1
mod1.HTT.MoveFirst
Do While Not mod1.HTT.EOF

    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "FydAddFwbt"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@qy") = mod1.HTT.Fields("qy").Value
    mod1.cmd.Parameters("@bm") = mod1.HTT.Fields("bm").Value
    mod1.cmd.Parameters("@trq") = Ddate
    mod1.cmd.Parameters("@ywy") = mod1.HTT.Fields("ywy").Value
    mod1.cmd.Parameters("@uid") = mod1.HTT.Fields("uid").Value
    mod1.cmd.Parameters("@Lcou") = 3 '流程总数
    mod1.cmd.Parameters("@Lc") = 1 '当前流程
    mod1.cmd.Parameters("@frq") = DateSerial(Year(Ddate), Month(Ddate), 1)
    mod1.cmd.Parameters("@lrq") = DateSerial(Year(Ddate), Month(Ddate), 28)
    mod1.cmd.Parameters("@lcRen") = mod1.HTT.Fields("ywy").Value
    mod1.cmd.Parameters("@lcUid") = mod1.HTT.Fields("uid").Value
    'mod1.CMD.Parameters("@nlb") = 67
    mod1.cmd.Parameters("@mon") = Str(Mon)
    mod1.cmd.Parameters("@fbt") = "房屋补贴报销单"
    mod1.cmd.Parameters("@fwbt") = mod1.HTT.Fields("fwbt").Value
    mod1.cmd.Parameters("@hg") = mod1.HTT.Fields("fwbt").Value
    mod1.cmd.Parameters("@hGd") = mod1.ChangBi(mod1.HTT.Fields("fwbt").Value)
    mod1.cmd.Parameters("@Lb") = 67
    mod1.cmd.Execute
    Bxid = mod1.cmd.Parameters("@bxid").Value
    Set cmd = Nothing
    
    '添加事务
    Call mod1.EnventAdd("报销单", mod1.HTT.Fields("fwbt").Value, mod1.HTT.Fields("ywy").Value, mod1.HTT.Fields("uid").Value, Str(Bxid), "报销人", "", "", mod1.HTT.Fields("ywy").Value, mod1.HTT.Fields("uid").Value, 0, Str(Bxid))
    
    '设置流程按钮
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "QMRZAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@NLb") = 67
    mod1.cmd.Parameters("@btz") = 23
    mod1.cmd.Parameters("@QDBH") = Bxid '编号
    mod1.cmd.Execute
    Set cmd = Nothing
    mod1.HTT.MoveNext
    'MsgBox ""
Loop
End Sub

Private Sub cmdLyf_Click()
Dim tt As String
Dim oo As Integer
Dim HYEAR As Integer
Dim Bxid As Long
Dim hg As Long
Dim GRen As String
Dim Guid As String
Dim Khmc As String
On Error Resume Next
Dim HH As Double
'tt = "select * from lyf"
'mod1.HTT.Close
'mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
HYEAR = Year(mod1.DQda)
'mod1.HTT.MoveFirst
'Do While Not mod1.HTT.EOF
tt = "select sum(lyf) as lyf from lyf"
mod1.HTT.Close
mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
hg = mod1.HTT.Fields("lyf").Value
HH = hg
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "fydAddLYF"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@qy") = "上海"
    mod1.cmd.Parameters("@bm") = "行政人事"
    mod1.cmd.Parameters("@trq") = mod1.DQda
    mod1.cmd.Parameters("@ywy") = "吴之禺"
    mod1.cmd.Parameters("@uid") = "HM025"
    mod1.cmd.Parameters("@Lcou") = 3 '流程总数
    mod1.cmd.Parameters("@Lc") = 1 '当前流程
    mod1.cmd.Parameters("@frq") = DateSerial(Year(mod1.DQda), 1, 1)
    mod1.cmd.Parameters("@lrq") = DateSerial(Year(mod1.DQda), 12, 31)
    mod1.cmd.Parameters("@lcRen") = "吴之禺"
    mod1.cmd.Parameters("@lcUid") = "HM025"
    mod1.cmd.Parameters("@nlb") = 72
    mod1.cmd.Parameters("@Hyear") = Str(HYEAR)
    mod1.cmd.Parameters("@fbt") = "旅游费报销单"
    'mod1.CMD.Parameters("@lyf") = mod1.HTT.Fields("lyf").Value
    mod1.cmd.Parameters("@hg") = hg
    mod1.cmd.Parameters("@hGd") = mod1.ChangBi(HH)
    mod1.cmd.Parameters("@Lb") = 72
    mod1.cmd.Execute
    Bxid = mod1.cmd.Parameters("@bxid").Value
    Set cmd = Nothing
    
tt = "select * from lyf"
mod1.HTT.Close
mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTT.MoveFirst
Do While Not mod1.HTT.EOF
    '检验归属人
    If mod1.HTT.Fields("bm") = "工程部" Then
        If mod1.HTT.Fields("username").Value = "" Or mod1.HTT.Fields("username").Value = "张寅" Then
            GRen = "张寅"
            Guid = "HM110"
        Else
            tt = "select gzu from worker where username='" & mod1.HTT.Fields("username").Value & "'"
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            Select Case mod1.HTP.Fields("gzu").Value
            Case 1
                GRen = "宋晓丹"
                Guid = "HM010"
            Case 2
                GRen = "滕伟"
                Guid = "HM022"
            Case 4
                GRen = "张寅"
                Guid = "HM110"
            Case 5
                GRen = "吴胜明"
                Guid = "HM031"
            End Select
        End If
    Else
        GRen = mod1.HTT.Fields("username").Value
        Guid = mod1.HTT.Fields("userId").Value
    End If
    'tt = "insert into fybx (bm,qy,ywy,bxid,xg,ywyuid,lyf,khmc) values (@bm,@qy,@ywy,@bxid,@lyf,@uid,@lyf,@hyear+'旅游费')"
    Khmc = mod1.HTT.Fields("username").Value & " 年限" & mod1.HTT.Fields("nx").Value
    tt = "insert into fybx (bm,qy,ywy,bxid,xg,ywyuid,lyf,khmc) values ('" & mod1.HTT.Fields("bm").Value & "','" & mod1.HTT.Fields("qy").Value & _
          "','" & GRen & "'," & Bxid & "," & mod1.HTT.Fields("lyf").Value & ",'" & mod1.HTT.Fields("userid").Value & "'," & mod1.HTT.Fields("lyf").Value & _
          ",'" & Khmc & "')"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    mod1.HTT.MoveNext
          
Loop
    
    
    '添加事务
    Call mod1.EnventAdd("报销单", Str(HH), "吴之禺", "HM025", Str(Bxid), "报销人", "", "", "吴之禺", "HM025", 0, Str(Bxid))
    
    '设置流程按钮
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "QMRZAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@NLb") = 72
    mod1.cmd.Parameters("@btz") = 23
    mod1.cmd.Parameters("@QDBH") = Bxid '编号
    mod1.cmd.Execute
    Set cmd = Nothing
    mod1.HTT.MoveNext
    'MsgBox ""

End Sub

Private Sub cmdOK_Click()
Dim tt As String
On Error GoTo ter
tt = "update worker set imgId=" & imgTX.Tag & " where userid='" & mod1.DHid & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = Nothing
If mod1.ZT = "HBData" Then
    tt = "update worker set imgId=" & imgTX.Tag & " where userid='" & mod1.DHid & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workHM, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
End If
'mod1.HTP.Close

Call TXBound
frmTX.Visible = False
Exit Sub
ter:
MsgBox "网络故障！再试一次！"
End Sub

Private Sub cmdRef_Click()
Call TXBound
End Sub

Private Sub cmdXZ_Click()
'Set frmZu.XForm = New frmZu
'''If mod1.Mname = "马晓聪" Then
'''   frmRenNew.Show
'''Else
    Call mod1.RenXz("frmZu", Me, 0)
'''End If
End Sub


Private Sub Command1_Click()
Dim tt As String
On Error Resume Next
Set Bjxt = CreateObject("adodb.recordset")
tt = "SELECT jzPb, jzXh, XT, wbX, wNr, kxF, gT, dGt, Pbid, xhId, xtid, dw, fjL,BZ , BF, Wid, dsl,xzF From Bjxt where jzpb<>'通用' ORDER BY Pbid, xhId, xtid, wId"
Bjxt.Close
Bjxt.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Bjxt.Save "c:\work\bjxt.xml"
tt = "SELECT jzPb, Pbid, jzXh, xhId From bjxt_jzxh order by pbid,xhid "
Bjxt.Close
Bjxt.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Bjxt.Save "c:\work\bjxtXh.xml"
tt = "select * from NewFuWu"
Bjxt.Close
Bjxt.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Bjxt.Save "c:\work\FW.xml"
Beep
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()



End Sub

Private Sub dtgOline_DblClick()
On Error Resume Next
Dim Ren As String
dtgOline.Col = 1
Ren = dtgOline.Text
frmOL.Show
frmOL.Left = frmZu.Left
frmOL.Top = 0
frmOL.ZOrder
frmOL.Caption = "您正在和" & Ren & "交谈"
End Sub

Private Sub Form_Activate()
frmZu.WindowState = 0
Call frmZu.OLOn
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If mod1.DName = "马晓聪" And KeyCode = 72 And Shift = 2 Then

        htBrowG.lblFw.Caption = "业务部"


    htBrowG.Visible = True
    htBrowG.ZOrder 0
End If
End Sub

Private Sub Form_Load()
Dim oo As Integer
Dim Ra
Dim La
Dim tt As String
'''''''''''NF.LoadSkin 3


    frmZu.Width = 14010
''''''If mod1.Mname <> "马晓聪" Then
''''''    tabZu.Visible = True
''''''    toolNew.Visible = False
''''''    frmZu.Height = 5910
''''''Else
    frmZu.Height = 8070

    toolNew.Visible = True
''''''End If

    NR(0).Style = 天蓝光泽
'加载QQ头像控件
For oo = 1 To 46
    Load NR(oo)
    NR(oo).Left = NR(oo - 1).Left + NR(oo - 1).Width + 20
    If oo Mod 5 = 0 Then
        NR(oo).Left = NR(0).Left
    End If
    NR(oo).Top = NR(oo - 1).Top
    If oo Mod 5 = 0 Then
        NR(oo).Top = NR(oo - 1).Top + NR(oo - 1).Height + 20
    End If
    NR(oo).Style = 天蓝光泽
    NR(oo).Visible = True
Next

''''''cmdXZ.Style = 蓝色经典2
''''''Set cmdXZ.PictureNormal = ImageList2.ListImages(77).Picture
''''''Set cmdXZ.PictureNormal = ImageList2.ListImages(78).Picture

cmdBack.Top = Me.Height - 900


'Call ResizeInit(Me) '在程序装入时必须加入


frmTX.Visible = False

If mod1.DName = "马晓聪" Then
'''''''''    cmdBi.Visible = True
'''''''''    cmdFwbt.Visible = True
'''''''''    cmdLyf.Visible = True
'''''''''    cmdFBGR.Visible = True
Else
'''''''''    cmdBi.Visible = False
'''''''''    cmdFwbt.Visible = False
'''''''''    cmdLyf.Visible = False
'''''''''    cmdFBGR.Visible = False
End If

On Error Resume Next
'''''''''''    tt = "select username,userid,imgid from oline order by bmid"
'''''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''''''''    RA = mod1.HTP.GetRows
'''''''''''    mod1.HTP.Close
'''''''''''    Set mod1.HTP = Nothing
'''''''''''    La = UBound(RA, 2) + 1
''''''''''''''''    dtgOline.Rows = 30
''''''''''''''''    dtgOline.Col = 1
''''''''''''''''    dtgOline.Row = 1
''''''''''''''''    For oo = 1 To La + 1
''''''''''''''''        dtgOline.Row = oo
''''''''''''''''        dtgOline.Text = Ra(0, oo - 1)
''''''''''''''''    Next
'''''''''''For oo = 1 To La
'''''''''''    tb1.Buttons.Add oo
'''''''''''    tb1.Buttons(oo).Caption = RA(0, oo - 1)
'''''''''''    tb1.Buttons(oo).Image = RA(2, oo - 1)
'''''''''''    tb1.Buttons(oo).Key = RA(1, oo - 1)
'''''''''''Next
'''''''''''For oo = La + 1 To 15
'''''''''''    tb1.Buttons.Add oo
'''''''''''Next
Call TXBound
MaxTop = (480 - 10200 + 5000)
MinTop = 480

'头像列表
tb2.Buttons.Clear
For oo = 1 To 76
    tb2.Buttons.Add oo
    tb2.Buttons(oo).Image = oo
Next
imgTX.Picture = ImageList2.ListImages(64).Picture
End Sub



Private Sub Form_Resize()

'Call mod1.ResizeForm(Me) '确保窗体改变时控件随之改变

frmZu.WindowState = 0

End Sub






Private Sub NB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'''''Dim oo As Integer
'''''For oo = 0 To 5
'''''    NB(oo).ForeColor = &H0&
'''''Next
'''''NB(Index).ForeColor = &HFF&
End Sub


Private Sub NR_Click(Index As Integer)
On Error Resume Next
Dim Ren As String
If NR(Index).Caption = mod1.DName Then
    frmTX.Visible = True
    meIndex = Button.Index
    imgTX.Picture = NR(Index).PictureNormal
ElseIf NR(Index).Caption <> "" Then
    Call frmOL.Tbound(NR(Index).Tag)
    MDI.timFl.Enabled = False
    Ren = NR(Index).ToolTipText
    frmOL.Show
    frmOL.Left = frmZu.Left
    frmOL.Top = 0
    frmOL.Caption = "您正在和" & Ren & "交谈"
    frmOL.img1.Picture = NR(Index).PictureNormal
    frmOL.lbl1.Caption = NR(Index).ToolTipText
    NR(Index).Caption = NR(Index).ToolTipText
    frmOL.lbl1.ToolTipText = NR(Index).Tag
    frmOL.img2.Picture = NR(meIndex).PictureNormal
    frmOL.lbl2.Caption = mod1.DName
    frmOL.txt1.SelStart = Len(txt1.Text)
    frmOL.txt1.SelLength = 0
    frmOL.txt2.Text = ""
    frmOL.txt2.SetFocus
    frmOL.ZOrder 0
    MDI.ztT.Panels(3).Text = ""
    MDI.ztT.Panels(3).Key = ""
    MDI.ztT.Panels(3).Tag = 0
End If
End Sub

Private Sub Slider1_Scroll()
tb1.Top = MaxTop * Slider1.Value / Slider1.Max + 120
End Sub

Private Sub Slider2_Scroll()
tb2.Top = MaxTop * Slider2.Value / Slider2.Max + 270
End Sub


Private Sub tabZu_Click(PreviousTab As Integer)
cmdBf(0).Visible = False
cmdBG(0).Visible = False
cmdBe(0).Visible = False
cmdBc(0).Visible = False
cmdBB(0).Visible = False
If tabZu.Tab = 5 Then
    cmdBf(0).Visible = True
End If
If tabZu.Tab = 6 Then
    cmdBG(0).Visible = True
    cmdBu(0).Visible = False
End If
If tabZu.Tab = 0 Then
    cmdBu(0).Visible = True
End If
If tabZu.Tab = 4 Then
    cmdBe(0).Visible = True
End If
If tabZu.Tab = 2 Then
    cmdBc(0).Visible = True
End If
If tabZu.Tab = 1 Then
    cmdBB(0).Visible = True
End If
End Sub

Private Sub tabZu_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
Dim oo As Integer
Dim hg As Double
On Error Resume Next

If mod1.DName = "马晓聪" And KeyCode = 72 And Shift = 2 Then
'''''    htBrowG.Visible = False
'''''    tt = "Select * from htView1 where 项目归属人='马晓聪' order by 部门,合同日期 desc"
'''''    htBrowG.adoBr.Close
'''''    htBrowG.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''    Set htBrowG.dtgBr.DataSource = htBrowG.adoBr
'''''    If htBrowG.adoBr.RecordCount > 0 Then
'''''        htBrowG.dtgBr.FixedRows = 0
'''''        htBrowG.dtgBr.MergeCol(1) = True
'''''        htBrowG.dtgBr.MergeCol(2) = True
'''''        htBrowG.dtgBr.MergeCol(3) = True
'''''        htBrowG.dtgBr.MergeCol(4) = True
'''''        htBrowG.dtgBr.MergeCol(7) = True
'''''        htBrowG.dtgBr.MergeCol(13) = True
'''''        htBrowG.dtgBr.MergeCells = 3
'''''        htBrowG.dtgBr.FixedRows = 1
'''''    End If

        htBrowG.lblFw.Caption = "业务部"


    htBrowG.Visible = True
    htBrowG.ZOrder 0
End If

End Sub


Private Sub tb1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Dim Ren As String
If Button.Caption = mod1.DName Then
    frmTX.Visible = True
    meIndex = Button.Index
    imgTX.Picture = ImageList2.ListImages(Button.Image).Picture
ElseIf Button.Caption <> "" Then
    Call frmOL.Tbound(Button.Key)
    MDI.timFl.Enabled = False
    Ren = Button.ToolTipText
    frmOL.Show
    frmOL.Left = frmZu.Left
    frmOL.Top = 0
    frmOL.Caption = "您正在和" & Ren & "交谈"
    frmOL.img1.Picture = ImageList2.ListImages(Button.Image).Picture
    frmOL.lbl1.Caption = Button.ToolTipText
    Button.Caption = Button.ToolTipText
    frmOL.lbl1.ToolTipText = Button.Key
    frmOL.img2.Picture = ImageList2.ListImages(tb1.Buttons(meIndex).Image).Picture
    frmOL.lbl2.Caption = mod1.DName
    frmOL.txt1.SelStart = Len(txt1.Text)
    frmOL.txt1.SelLength = 0
    frmOL.txt2.Text = ""
    frmOL.txt2.SetFocus
    frmOL.ZOrder 0
    MDI.ztT.Panels(3).Text = ""
    MDI.ztT.Panels(3).Key = ""
    MDI.ztT.Panels(3).Tag = 0
End If
End Sub

Private Sub tb2_ButtonClick(ByVal Button As MSComctlLib.Button)
imgTX.Picture = ImageList2.ListImages(Button.Index).Picture
imgTX.Tag = Button.Index
End Sub

Private Sub TBa_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tt As String
Dim oo As Integer
Dim ii As Integer
Dim Ra: Dim La
Dim pk As String
On Error Resume Next
'MsgBox Button.Index
Call frmZu.OLOff
Select Case Button.Index
    Case 2 '公告栏
        Call modGGL.CHZT
        
        tt = "Select top 1 gid from ggl where (" & mod1.DName & "=0 or lb='胡萝卜' and " & mod1.DName & " is null  order by gid desc"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
        'Oid = mod1.HTP.Fields("gid").Value

        If mod1.HTP.RecordCount > 0 Then
            Set frmGGL.adoGGl = CreateObject("adodb.recordset")
            tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid from ggl where " & mod1.DName & "=0 or lb='胡萝卜' and " & mod1.DName & " is null order by gid desc"
            frmGGL.adoGGl.Close
            frmGGL.adoGGl.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
            modGGL.Oid = frmGGL.adoGGl.Fields("gid").Value
            frmGGL.rihNr.Text = frmGGL.adoGGl.Fields("Gnr").Value
            If Left(frmGGL.adoGGl.Fields("zz").Value, 1) = "n" Then
                frmGGL.lblZZ.Caption = "匿名者"
            Else
                frmGGL.lblZZ.Caption = frmGGL.adoGGl.Fields("zz").Value
            End If
            frmGGL.lblDate.Caption = frmGGL.adoGGl.Fields("rq").Value
            If IsNull(frmGGL.adoGGl.Fields("lb").Value) = True Then
                frmGGL.comLb.Visible = False
                frmGGL.lblLb.Visible = False
            Else
                frmGGL.comLb.Text = frmGGL.adoGGl.Fields("lb").Value
                frmGGL.comLb.Visible = True
                frmGGL.lblLb.Visible = True
                frmGGL.comLb.Locked = True

            End If
            frmGGL.Show

            frmGGL.ZOrder 0
            'frmZu.Enabled = False
            
            '判断字颜色
            frmGGL.rihNr.SelStart = 0
            frmGGL.rihNr.SelLength = Len(frmGGL.rihNr.Text)
        
                frmGGL.rihNr.SelColor = &HFF0000
        
            frmGGL.rihNr.SelFontSize = frmGGL.adoGGl.Fields("Fdx").Value
            frmGGL.rihNr.SelStart = 0
            frmGGL.rihNr.SelLength = 0
        End If
        frmGGL.Show
        frmGGL.cmdSave.Enabled = False
        frmGGL.cmdReply.Enabled = True
        frmGGL.ZOrder 0
        frmGGL.cmdZx.Enabled = True
        frmGGL.comLb.Locked = True

    Case 3 '报销
        mod1.BTZ = 23
        'frmBxBrow.WindowState = 0
        frmBxBrow.Show
        'frmBxBrow.WindowState = 2
        Set frmBxBrow.AdoBxBro = CreateObject("adodb.recordset")
        
        tt = "FydV('" & mod1.DHid & "','" & mod1.DName & "')"
        If mod1.DName = "李莉娜" Then
            tt = "FydV('HM025','吴之禺')"
        End If
        frmBxBrow.AdoBxBro.Close
        frmBxBrow.AdoBxBro.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adcmdstoreproc
'        tt = "Select convert (nvarchar(25),frq,1) as 日期范围 ,convert (nvarchar(25),lrq,1) as 日期范围 ,HG as 金额,bxid as 报销单编号,convert (nvarchar(25),qrq,1) as 签收日期 from FyDBrow " & _
'       "where Uid='" & mod1.DHid & "' and ywy='" & mod1.DName & "' and not (hg is null)"
'        frmBxBrow.AdoBxBro.Close
'        frmBxBrow.AdoBxBro.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmBxBrow.mga.DataSource = frmBxBrow.AdoBxBro
        'PK = "<起 始 期  |<截 至 期  |>  金 额 |^ 报 销 单 编 号|> 签收日期 "
        'PK = "^  日期范围|^  日期范围|>  金 额 |^ 报 销 单 编 号|> 签收日期 "
         frmBxBrow.mga.ColWidth(0) = 500
        'frmBxBrow.mga.FormatString = PK
        frmBxBrow.mga.MergeRow(0) = True
        frmBxBrow.mga.MergeCells = flexMergeRestrictAll
        frmBxBrow.optMe.Value = True
        'frmBxBrow.BorderStyle = 1
        'frmBxBrow.MGa.FixedCols = 1
'        frmBxBrow.MaxButton = False
'        frmBxBrow.MinButton = False

        '生成报销单按钮
        For oo = 15 To 1 Step -1
            Unload frmBxBrow.cmdFyd(oo)
        Next
        frmBxBrow.cmdFyd(0).Caption = Trim(mod1.fydA(0, 0)) & "报销单"
        frmBxBrow.cmdFyd(0).Tag = mod1.fydA(1, 0)
        frmBxBrow.cmdFyd(0).ToolTipText = Trim(mod1.fydA(3, 0)) & ",流程的总数为:" & mod1.fydA(2, 0)
         frmBxBrow.cmdFyd(0).Visible = True
        For oo = 1 To mod1.fyuA
            '如果有重复的名称(主要为<=500)的钱),则不显示按钮
            If Not (mod1.fydA(1, oo) = 47 Or _
            mod1.fydA(1, oo) = 68 Or mod1.fydA(1, oo) = 81 Or mod1.fydA(1, oo) = 282 Or _
            mod1.fydA(1, oo) = 141 Or mod1.fydA(1, oo) = 187 Or mod1.fydA(1, oo) = 278) Then
            Load frmBxBrow.cmdFyd(oo)
            frmBxBrow.cmdFyd(oo).Caption = Trim(mod1.fydA(0, oo)) & "报销单"
            frmBxBrow.cmdFyd(oo).Tag = mod1.fydA(1, oo)
            frmBxBrow.cmdFyd(oo).ToolTipText = Trim(mod1.fydA(3, oo)) & ",流程的总数为:" & mod1.fydA(2, oo)
                If oo = 4 Then
                    frmBxBrow.cmdFyd(oo).Left = frmBxBrow.cmdFyd(oo).Left + 2000
                    frmBxBrow.cmdFyd(oo).Top = frmBxBrow.cmdFyd(0).Top
                ElseIf oo = 5 Then
                    frmBxBrow.cmdFyd(oo).Left = frmBxBrow.cmdFyd(oo).Left + 2000
                    frmBxBrow.cmdFyd(oo).Top = frmBxBrow.cmdFyd(4).Top + 1500
                ElseIf oo = 6 Then
                    frmBxBrow.cmdFyd(oo).Left = frmBxBrow.cmdFyd(oo).Left + 2000
                    frmBxBrow.cmdFyd(oo).Top = frmBxBrow.cmdFyd(5).Top + 1500
                ElseIf oo = 7 Then
                    frmBxBrow.cmdFyd(oo).Left = frmBxBrow.cmdFyd(oo).Left + 2000
                    frmBxBrow.cmdFyd(oo).Top = frmBxBrow.cmdFyd(6).Top + 1500
                Else
                   frmBxBrow.cmdFyd(oo).Top = frmBxBrow.cmdFyd(oo - 1).Top + 1500
                   frmBxBrow.cmdFyd(oo).Left = frmBxBrow.cmdFyd(0).Left
                End If
                frmBxBrow.cmdFyd(oo).Visible = True

            End If

        Next
        If IsNull(mod1.fyuA) = True Then
            frmBxBrow.cmdFyd(0).Visible = False
        End If
        If frmBxBrow.cmdFyd(0).Caption = "" Then
            frmBxBrow.cmdFyd(0).Visible = False
        End If
        frmBxBrow.frmYj.Visible = False
        frmBxBrow.Enabled = True
        frmBxBrow.ZOrder 0
    Case 4 '当前事务
        Dim RL
        Dim ul
        
        Call mod1.refEnt1
        tt = mod1.ETT & " and delf=1 order by rq desc"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        RL = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        ul = UBound(RL, 2)
        Call mod1.refEnt2(RL, ul)
        Dialog.Show: Dialog.ZOrder 0
        frmZu.Enabled = False
    Case 5 '信息
        'Exit Sub
'        XinXi.Show
'        frmZu.Enabled = False
        'Exit Sub
        'frmTip.Show
    Case 6 '企业文化
        '
        Exit Sub
        frmYG.Show
    Case 7 '绩效考核
'''''        Call b1.KPIQing
'''''        Call b1.KPIBound(mod1.DName, mod1.DHid, mod1.DQda)
''''''''        Dim OBm As String
       tt = "select bm,username,phoD,userPho,phoX,userid from renyuan1 where username<>'匿名者' and zzf=1 order by bmid,userid"
''''''''''''''''''''''''        Set mod1.HTP = CreateObject("adodb.recordset")
''''''''''''''''''''''''        mod1.HTP.Open TT, mod1.workFF, adOpenForwardOnly, adLockReadOnly, adCmdText
''''''''''''''''''''''''        Ra = mod1.HTP.GetRows
''''''''''''''''''''''''        mod1.HTP.Close
''''''''''''''''''''''''        Set mod1.HTP = Nothing
''''''''''''''''''''''''        La = UBound(Ra, 2)
''''''''''''''''''''''''        DHB.dtgDHB.Rows = La + 30
''''''''''''''''''''''''        DHB.dtgLL.Rows = DHB.dtgDHB.Rows
''''''''''''''''''''''''        DHB.dtgLL.Cols = DHB.dtgDHB.Cols
''''''''''''''''''''''''
''''''''''''''''''''''''        For oo = 1 To La + 1
''''''''''''''''''''''''            DHB.dtgDHB.Row = oo: DHB.dtgLL.Row = oo
''''''''''''''''''''''''            For ii = 1 To 10
''''''''''''''''''''''''                DHB.dtgDHB.Col = ii: DHB.dtgLL.Col = ii
''''''''''''''''''''''''                DHB.dtgDHB.Text = Ra(ii - 1, oo - 1)
''''''''''''''''''''''''                DHB.dtgLL.Text = DHB.dtgDHB.Text
''''''''''''''''''''''''                If ii = 1 Then
''''''''''''''''''''''''                    If OBm <> DHB.dtgDHB.Text Then
''''''''''''''''''''''''                        OBm = DHB.dtgDHB.Text
''''''''''''''''''''''''                    Else
''''''''''''''''''''''''                        DHB.dtgDHB.Text = ""
''''''''''''''''''''''''                    End If
''''''''''''''''''''''''                End If
''''''''''''''''''''''''            Next
''''''''''''''''''''''''        Next

'''''''        DHB.dtgDHB.MergeCol(1) = True
'''''''        DHB.dtgDHB.MergeCells = 3

        
'''''        DHB.Show
        DHB.TC = tt
        Call DHB.Bound(tt)
        DHB.Visible = False
        frmRenNew.lblTitle.Visible = False
        frmRenNew.Show
    Case 8
        frmMeet.Show
        frmMeet.ZOrder 0
        
    Case 9 '在线更新
        ii = MsgBox("在线更新，将关闭豪曼信息！", vbYesNo + vbInformation, "豪曼科技")
        If ii = vbNo Then Exit Sub
        
        Dim bt() As Byte
        'Dim tt As String
        On Error Resume Next
        tt = "select Nup,Nz from upfile "
        frmGGL.adoFile.Recordset.Close
        frmGGL.adoFile.Recordset.Open tt, mod1.wzcc, adOpenKeyset, adLockReadOnly, adCmdText
        ReDim bt(frmGGL.adoFile.Recordset.Fields("Nz").Value) As Byte
        bt() = frmGGL.adoFile.Recordset.Fields("Nup").GetChunk(frmGGL.adoFile.Recordset.Fields("Nz").Value + 1)


        Open ("c:\work\demo\hmxp9000\" & "update.exe") For Binary As #3
        Put #3, , bt()
        Close #3
        
        '启动更新应用程序
        Shell (App.Path & "\update.exe"), vbNormalFocus
        End
        
End Select

End Sub






















Private Sub VScroll1_Change()

End Sub


Private Sub VScroll1_Scroll()

End Sub


Private Sub Va_Change()
 Dim i, j, k
        i = Va.Value
        If Location < Va.Value Then           '向下移动
              For k = 1 To i - Location
                  For j = 0 To Va.Max
                          tb1.Top = tb1.Top - tb1.Height
                  Next j
              Next k

        Else
              For k = 1 To Location - i
                  For j = 0 To Va.Max
                          tb1.Top = tb1.Top + tb1.Height

                  Next j
              Next k
                      

        End If
        Location = Va.Value

End Sub





Private Sub Va_Scroll()
Dim i, j, k
        i = Va.Value
        If Location < Va.Value Then           '向下移动
              For k = 1 To i - Location
                  For j = 0 To Va.Max
                          
                        tb1.Top = tb1.Top - tb1.Height
                  Next j
              Next k

        Else
              For k = 1 To Location - i
                  For j = 0 To Va.Max
                          
                        tb1.Top = tb1.Top + tb1.Height

                        
                  Next j
              Next k
                      

        End If
        Location = Va.Value


End Sub


Public Sub TXBound()
Dim tt As String
Dim oo As Integer
Dim ii As Integer
Dim MM As Integer
Dim mi As Integer
Dim Ra
Dim La
Dim Lb
Dim Rb
Dim REF As Boolean
Dim cc As Integer

On Error GoTo TERR
    tt = "select username,userid,imgid from oline order by bmid,userid;" & _
        "select zz,zuid,imgid,count(gid) from hmtext.dbo.oline where  tuid='" & mod1.DHid & "' group by zuid,zz,imgid,bmid order by bmid"
    Set mod1.OLT = CreateObject("adodb.recordset")
    mod1.OLT.Open tt, mod1.workHM, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.OLT.GetRows
    Set mod1.OLT = mod1.OLT.NextRecordset
    If mod1.OLT.BOF = False Then
        Rb = mod1.OLT.GetRows
        Lb = UBound(Rb, 2)
    Else
        Lb = -1
    End If
    mod1.OLT.Close
    Set mod1.OLT = Nothing
    La = UBound(Ra, 2) + 1
    
On Error Resume Next
    For oo = tb1.Buttons.Count - 1 To 100
        ORa(oo, 0) = ""
        ORa(oo, 1) = ""
    Next

REF = False
For oo = 0 To La
    If ORa(oo, 0) = "" And oo > 0 Then
        Exit For
    End If
    If (ORa(oo, 0) <> Ra(1, oo) Or ORa(oo, 1) <> Ra(2, oo)) Then
    'Set ORa = Nothing
        'For ii = 0 To La - 1

        'Next
        REF = True
        Exit For
    End If
Next
'''If ORa(oo, 0) <> Ra(1, oo) Then
'''    REF = True
'''End If
'''''''''If ORa(La, 0) <> tb1.Buttons(La + 1).Key Or ORa(La, 1) <> tb1.Buttons(La).Image Then
'''''''''    REF = True
'''''''''End If
If REF = True Then
tb1.Visible = False
tb1.Buttons.Clear

For oo = 0 To La
'    tb1.Buttons.Add oo
'    tb1.Buttons(oo).Caption = Ra(0, oo - 1)
'    tb1.Buttons(oo).ToolTipText = Ra(0, oo - 1)
'    tb1.Buttons(oo).Image = Ra(2, oo - 1)
'    tb1.Buttons(oo).Key = Ra(1, oo - 1)
'    If tb1.Buttons(oo).Key = mod1.DHid Then
'        meIndex = oo
'    End If
'            ORa(oo - 1, 0) = Ra(1, oo - 1)
'            ORa(oo - 1, 1) = Ra(2, oo - 1)
'    If tb1.Buttons(oo).Caption = "" Then
'        tb1.Buttons(oo).Caption = tb1.Buttons(oo).ToolTipText
'    End If
    NR(oo).Caption = Ra(0, oo)
    NR(oo).ToolTipText = Ra(0, oo)
    Set NR(oo).PictureNormal = ImageList2.ListImages(Ra(2, oo)).Picture
    NR(oo).Tag = Ra(1, oo)
    If NR(oo).Tag = mod1.DHid Then
        meIndex = oo
    End If
    NR(oo).Style = XP时代
    NR(oo).Visible = True
    ORa(oo, 0) = Ra(1, oo)
    ORa(oo, 1) = Ra(2, oo)
    DoEvents
    
Next
cc = La
'For oo = La + 1 To 50
'    tb1.Buttons.Add oo
'Next

'tb1.Visible = True
'tb1.Refresh
End If
For oo = cc To 47
    NR(oo).Visible = False
Next

'查看有无未看短信
If Lb = -1 Then Exit Sub
REF = False
For ii = 0 To Lb
    For oo = 0 To 100
        If ORa(oo, 0) = Rb(1, ii) Then
            'REF = True
            Exit For
        End If
    Next
    If oo > 99 Then
        REF = True
    End If
    If REF = True Then
        cc = cc + 1
'        tb1.Buttons.Add ii
'        tb1.Buttons(cc).Caption = Rb(0, ii)
'        tb1.Buttons(cc).ToolTipText = Rb(0, ii)
'        tb1.Buttons(cc).Image = Rb(2, ii)
'        tb1.Buttons(cc).Key = Rb(1, ii)
'        tb1.Buttons(cc).MixedState = True
    NR(cc - 1).Caption = Rb(0, ii) & "*" & Rb(3, ii)
    NR(cc - 1).ToolTipText = Rb(0, ii)
    Set NR(cc - 1).PictureNormal = ImageList2.ListImages(Rb(2, ii)).Picture
    NR(cc - 1).Style = 条纹之美
    NR(cc - 1).Tag = Rb(1, ii)
    NR(cc - 1).Visible = True
        ORa(cc - 1, 0) = Rb(1, ii)
        ORa(cc - 1, 1) = Rb(2, ii)
    End If
Next
For oo = cc + 1 To 47
    NR(oo).Visible = False
Next
Exit Sub
TERR:
Exit Sub
End Sub

Private Sub timOline_Timer()
Call TXBound
End Sub


Private Sub timRev_Timer()
Dim ZTB As Boolean
Dim ZTR As String
Dim tt As String
Dim oo As Integer
Dim ii As Integer
Dim cc As Integer
Dim Ra
Dim La
Dim Rb
Dim TRev As Object
ZTB = False
On Error Resume Next
Dim ZTid As String
tt = "select zz,zuid,imgid,count(gid) from hmtext.dbo.oline where  tuid='" & mod1.DHid & "' group by zuid,zz,imgid,bmid order by bmid"
Set TRev = CreateObject("adodb.recordset")
TRev.Open tt, mod1.workHM, adOpenForwardOnly, adLockReadOnly, adCmdText
If TRev.BOF = False Then
    Ra = TRev.GetRows
    TRev.Close
    Set TRev = Nothing
    La = UBound(Ra, 2) + 1
Else
    Exit Sub
End If
'''''cc = tb1.Buttons.Count
'''''For ii = 1 To La
'''''    ZTB = False
'''''    For oo = 1 To cc
'''''        If tb1.Buttons(oo).Key = Ra(1, ii - 1) Then
'''''            tb1.Buttons(oo).Caption = tb1.Buttons(oo).ToolTipText & " *" & Ra(3, ii - 1)
'''''            ZTB = True
'''''                    ZTR = ZTR & " " & tb1.Buttons(oo).ToolTipText & " " & Ra(3, ii - 1)
'''''                    MDI.ztT.Panels(3).Key = tb1.Buttons(oo).Key
'''''                    MDI.ztT.Panels(3).ToolTipText = tb1.Buttons(oo).ToolTipText
'''''                    MDI.ztT.Panels(3).Tag = tb1.Buttons(oo).Image
'''''            Exit For
'''''        End If
'''''    Next
'''''Next
On Error Resume Next
cc = 50
For ii = 1 To La
    ZTB = False
    For oo = 0 To cc
        If NR(oo).Tag = Ra(1, ii - 1) Then
            NR(oo).Caption = NR(oo).ToolTipText & " *" & Ra(3, ii - 1)
            ZTB = True
                    ZTR = ZTR & " " & NR(oo).ToolTipText & " " & Ra(3, ii - 1)
                    MDI.ztT.Panels(3).Key = NR(oo).Tag
                    MDI.ztT.Panels(3).ToolTipText = NR(oo).ToolTipText
                    MDI.ztT.Panels(3).Tag = Ra(2, ii - 1)
            Exit For
        End If
    Next
Next


    If ZTB = True Then
        MDI.ztT.Panels(3).Text = "您有未看信息：" & ZTR
        'If MDI.WindowState = 1 Then
            If frmOL.WindowState <> 0 And frmOL.Visible = True Or frmOL.Visible = False Then
                MDI.timFl.Enabled = True
            Else
                MDI.timFl.Enabled = False
            End If
        'End If
    Else
        If Left(MDI.ztT.Panels(3).Text, 6) = "您有未看信息" Then
            MDI.ztT.Panels(3).Text = ""
        End If
    End If
End Sub



Public Sub OLOff()
'timOline.Enabled = False
timOline.Interval = 60000
timRev.Interval = 20000
End Sub

Public Sub OLOn()
'timOline.Enabled = True
timOline.Interval = 20000
timRev.Interval = 5000
End Sub

Public Sub AN(Pi As Integer)
Dim oo As Integer
On Error Resume Next
For oo = 0 To 20
    cmdBu(oo).Visible = False
Next
For oo = 0 To 20
    cmdBc(oo).Visible = False
Next
For oo = 0 To 20
    cmdBB(oo).Visible = False
Next
For oo = 0 To 20
    cmdBe(oo).Visible = False
Next
For oo = 0 To 20
    cmdBf(oo).Visible = False
Next
For oo = 0 To 20
    cmdBG(oo).Visible = False
Next

Select Case Pi
Case 1
    For oo = 0 To 20
        cmdBu(oo).Visible = True
    Next
    NC.Caption = "维保业务"
Case 2
    For oo = 0 To 20
        cmdBc(oo).Visible = True
    Next
    NC.Caption = "采购服务"
Case 3
    For oo = 0 To 20
        cmdBB(oo).Visible = True
    Next
    NC.Caption = "行政人事"
Case 4
    For oo = 0 To 20
        cmdBe(oo).Visible = True
    Next
    NC.Caption = "产品事务"
Case 5
    For oo = 0 To 20
        cmdBf(oo).Visible = True
    Next
    NC.Caption = "财务统计"
Case 6
    For oo = 0 To 20
        cmdBG(oo).Visible = True
    Next
    NC.Caption = "管理信息"
End Select
End Sub

Private Sub toolNew_ButtonClick(ByVal Button As MSComctlLib.Button)
Call AN(Button.Index)
End Sub


