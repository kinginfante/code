VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmComputer 
   BackColor       =   &H00C0FFC0&
   Caption         =   "固定资产"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15180
   FillColor       =   &H00FFC0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15180
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "IT设备"
      Height          =   615
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   8460
      Width           =   675
   End
   Begin VB.TextBox txtId 
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   8820
      TabIndex        =   50
      Top             =   8670
      Width           =   435
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00305AED&
      Caption         =   "关机"
      Height          =   345
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8640
      Width           =   705
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   8520
      Top             =   4350
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   9270
      Top             =   7410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   465
      Index           =   31
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7470
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   30
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7020
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   29
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6180
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   28
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4890
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   27
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7020
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   26
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6180
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   25
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1830
      Width           =   315
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   24
      Left            =   10110
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1830
      Width           =   315
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   23
      Left            =   9780
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1830
      Width           =   315
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   22
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3300
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   21
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2970
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   20
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2130
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   19
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2970
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   18
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2130
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   17
      Left            =   14430
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2970
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   16
      Left            =   14430
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2130
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   15
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1440
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   14
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2310
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   13
      Left            =   5565
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2310
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   12
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2310
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   11
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3690
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   10
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3690
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   9
      Left            =   5565
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3690
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   8
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3690
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   7
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4290
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   6
      Left            =   5565
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4290
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   5
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4290
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   4
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4290
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "3"
      Height          =   405
      Index           =   3
      Left            =   5565
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5580
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "2"
      Height          =   405
      Index           =   2
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5580
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   1
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8460
      Width           =   435
   End
   Begin VB.CommandButton cmdComputer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      Height          =   405
      Index           =   0
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   585
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8460
      Width           =   675
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "控制"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   8370
      TabIndex        =   49
      Top             =   8700
      Width           =   435
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   47
      Top             =   240
      Width           =   195
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "关机,休眠"
      Height          =   225
      Left            =   570
      TabIndex        =   46
      Top             =   270
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   45
      Top             =   480
      Width           =   195
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "运行"
      Height          =   165
      Left            =   600
      TabIndex        =   44
      Top             =   480
      Width           =   465
   End
   Begin VB.Label Label5 
      BackColor       =   &H00255EF3&
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   43
      Top             =   750
      Width           =   195
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "硬件异常"
      Height          =   195
      Left            =   570
      TabIndex        =   42
      Top             =   750
      Width           =   885
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   945
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "开票"
      Height          =   165
      Left            =   10290
      TabIndex        =   35
      Top             =   2430
      Width           =   435
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "前台"
      Height          =   225
      Index           =   2
      Left            =   6540
      TabIndex        =   8
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "杰升"
      Height          =   225
      Index           =   1
      Left            =   1740
      TabIndex        =   7
      Top             =   5490
      Width           =   1005
   End
   Begin VB.Line Line5 
      Index           =   1
      X1              =   1800
      X2              =   1800
      Y1              =   5820
      Y2              =   7710
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   930
      X2              =   2700
      Y1              =   6750
      Y2              =   6750
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1905
      Index           =   11
      Left            =   930
      Shape           =   4  'Rounded Rectangle
      Top             =   5820
      Width           =   1785
   End
   Begin VB.Line Line6 
      Index           =   1
      X1              =   150
      X2              =   690
      Y1              =   6750
      Y2              =   6750
   End
   Begin VB.Shape Shape4 
      Height          =   1935
      Index           =   3
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5820
      Width           =   585
   End
   Begin VB.Line Line7 
      X1              =   780
      X2              =   780
      Y1              =   7980
      Y2              =   5280
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   2685
      Index           =   10
      Left            =   0
      Top             =   5280
      Width           =   2985
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "财务部"
      Height          =   225
      Index           =   0
      Left            =   10530
      TabIndex        =   6
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Line Line6 
      Index           =   0
      X1              =   14040
      X2              =   15090
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Shape Shape4 
      Height          =   1935
      Index           =   2
      Left            =   14070
      Shape           =   4  'Rounded Rectangle
      Top             =   1770
      Width           =   1065
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   12240
      X2              =   12240
      Y1              =   1800
      Y2              =   3690
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   11370
      X2              =   13140
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1905
      Index           =   9
      Left            =   11370
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   1785
   End
   Begin VB.Shape Shape4 
      Height          =   1425
      Index           =   1
      Left            =   9720
      Shape           =   5  'Rounded Square
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Shape Shape4 
      Height          =   1185
      Index           =   0
      Left            =   9720
      Shape           =   5  'Rounded Square
      Top             =   2910
      Width           =   1065
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   2775
      Index           =   8
      Left            =   9720
      Top             =   1260
      Width           =   5415
   End
   Begin VB.Line Line3 
      Index           =   11
      X1              =   4530
      X2              =   4530
      Y1              =   1950
      Y2              =   1260
   End
   Begin VB.Line Line3 
      Index           =   10
      X1              =   5820
      X2              =   5820
      Y1              =   1920
      Y2              =   1260
   End
   Begin VB.Line Line3 
      Index           =   9
      X1              =   7050
      X2              =   7050
      Y1              =   1920
      Y2              =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "总经理"
      Height          =   315
      Index           =   3
      Left            =   10500
      TabIndex        =   5
      Top             =   420
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "销售总监"
      Height          =   315
      Index           =   2
      Left            =   8490
      TabIndex        =   4
      Top             =   420
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "销售经理"
      Height          =   315
      Index           =   1
      Left            =   6990
      TabIndex        =   3
      Top             =   420
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "大会议室"
      Height          =   315
      Index           =   0
      Left            =   4170
      TabIndex        =   2
      Top             =   420
      Width           =   1395
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   7
      Left            =   9750
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   5385
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   6
      Left            =   8280
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   1485
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   5
      Left            =   6750
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   1545
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   4
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   3525
   End
   Begin VB.Line Line3 
      Index           =   8
      X1              =   5280
      X2              =   5280
      Y1              =   3360
      Y2              =   2130
   End
   Begin VB.Line Line3 
      Index           =   7
      X1              =   6330
      X2              =   6330
      Y1              =   3360
      Y2              =   2130
   End
   Begin VB.Line Line3 
      Index           =   6
      X1              =   7350
      X2              =   7350
      Y1              =   3360
      Y2              =   2130
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   4260
      X2              =   4260
      Y1              =   3360
      Y2              =   2130
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   4260
      X2              =   8325
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1245
      Index           =   2
      Left            =   3300
      Shape           =   4  'Rounded Rectangle
      Top             =   2130
      Width           =   5025
   End
   Begin VB.Line Line3 
      Index           =   5
      X1              =   5280
      X2              =   5280
      Y1              =   4740
      Y2              =   3540
   End
   Begin VB.Line Line3 
      Index           =   4
      X1              =   6300
      X2              =   6300
      Y1              =   4740
      Y2              =   3540
   End
   Begin VB.Line Line3 
      Index           =   3
      X1              =   7350
      X2              =   7350
      Y1              =   4740
      Y2              =   3540
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   4260
      X2              =   4260
      Y1              =   4740
      Y2              =   3540
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   4260
      X2              =   8325
      Y1              =   4230
      Y2              =   4230
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Index           =   1
      Left            =   3300
      Shape           =   4  'Rounded Rectangle
      Top             =   3540
      Width           =   5025
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   5280
      X2              =   5280
      Y1              =   6120
      Y2              =   4920
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   6300
      X2              =   6300
      Y1              =   6120
      Y2              =   4920
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   7350
      X2              =   7350
      Y1              =   6120
      Y2              =   4920
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   4260
      X2              =   4260
      Y1              =   6120
      Y2              =   4920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "商务支持部"
      Height          =   1035
      Left            =   5130
      TabIndex        =   1
      Top             =   7170
      Width           =   315
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFC0C0&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   2625
      Left            =   3360
      Shape           =   5  'Rounded Square
      Top             =   6450
      Width           =   2445
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4260
      X2              =   8325
      Y1              =   5490
      Y2              =   5490
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   3
      Left            =   3300
      Shape           =   4  'Rounded Rectangle
      Top             =   1260
      Width           =   5025
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Index           =   0
      Left            =   3300
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   5025
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   6300
      Shape           =   4  'Rounded Rectangle
      Top             =   8310
      Width           =   2025
   End
End
Attribute VB_Name = "frmComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim LF As Boolean '闪灯
Private Sub cmdBack_Click()
Timer1.Enabled = False
Me.Visible = False
End Sub

Private Sub cmdClose_Click()
Dim tt As String
On Error GoTo comERR1

If mod1.DName = "马晓聪" Or mod1.DName = "郑刚" Or mod1.DName = "宋晓炯" Then
    Set mod1.HTP = New ADODB.Recordset
    tt = "update computer set offC=1 where cid=" & Val(txtId.Text)
    mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
    MsgBox "系统将在十秒钟内响应关机!"
End If

Exit Sub
comERR1:
MsgBox "出错!"
End Sub

Private Sub cmdComputer_Click(Index As Integer)
Dim XH As Integer
Dim tt As String
Dim ii As Integer
Dim GetChar As String
Dim Ra
Dim Rb
XH = Index

tt = "select cpu,memory,hd,motherboard,monitor,CDrom from computer where cid=" & XH & ";" & _
    "select cip from computerIP where cid=" & XH
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
GetChar = "  芯片: " & Ra(0, 0) & Chr(13) & Chr(10) & "  内存: " & Ra(1, 0) & Chr(13) & Chr(10) & "  硬盘: " & Trim(Ra(2, 0)) & Chr(13) & Chr(10) & _
            "  主板: " & Ra(3, 0) & Chr(13) & Chr(10) & "显示器: " & Ra(4, 0) & Chr(13) & Chr(10) & "  光驱: " & Ra(5, 0) & Chr(13) & Chr(10) & _
            "IP地址: " & Rb(0, 0) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "是否查看详细配置?"
ii = MsgBox(GetChar, vbYesNo + vbInformation + vbDefaultButton2, "硬件配置")
If ii = vbYes Then
    frmComDetail.Show
    frmComDetail.ZOrder 0
    Call frmComDetail.Initialize
    Call frmComDetail.Bound(Index)
    
End If
txtId.Text = Index
End Sub

Private Sub Command1_Click()
frmComJian.Show
Call frmComJian.Bound
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Dim oo As Integer
On Error Resume Next
For oo = 0 To 31
    cmdComputer(oo).Caption = oo
Next
If mod1.DName = "马晓聪" Then
    cmdComputer(0).Visible = True
End If
End Sub

Public Sub OlineF()
Dim oo As Integer
Dim ii As Integer
Dim tt As String
Dim Ra
Dim Rb
Dim La
tt = "select datediff(second,uptime,getdate()),cpu,cpudetail,memory,memorydetail,hd,hddetail,motherboard,mbdetail,monitor,monitordetail,cdrom,cdromdetail from computer order by cid;" & _
    "select uptime,cpu,cpudetail,memory,memorydetail,hd,hddetail,motherboard,mbdetail,monitor,monitordetail,cdrom,cdromdetail from computerO order by cid"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
ReDim Preserve mod1.Light(La) As Boolean
For oo = 1 To La
    If IsNull(Ra(0, oo - 1)) = True Then
        Ra(0, oo - 1) = "1900-1-1"
    End If
    If Val(Ra(0, oo - 1)) < 200 Then

        cmdComputer(oo - 1).BackColor = &HC0FFFF
        cmdComputer(oo - 1).Tag = &HC0FFFF

    Else

        cmdComputer(oo - 1).BackColor = &HFFFFC0
        cmdComputer(oo - 1).Tag = &HFFFFC0
    End If
    mod1.Light(oo - 1) = False
    For ii = 1 To 12
        If Ra(ii, oo - 1) <> Rb(ii, oo - 1) Then '检测硬件不同
            mod1.Light(oo - 1) = True
            If (oo = 24 Or oo = 25 Or oo = 26) And (ii = 9 Or ii = 10) Then
                mod1.Light(oo - 1) = False
            Else
                Exit For
            End If
        End If

    Next
Next
LF = True
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Timer1.Enabled = False
Cancel = True
End Sub


Private Sub Timer1_Timer()

Dim oo As Integer

For oo = 0 To 31
    If mod1.Light(oo) = True Then

        If LF = True Then
            cmdComputer(oo).BackColor = &H255EF3
        Else


            cmdComputer(oo).BackColor = cmdComputer(oo).Tag
        End If
    End If
Next
If LF = False Then
    LF = True
Else
    LF = False
End If
End Sub


