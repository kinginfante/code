VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmRen 
   Caption         =   "Ա������"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdXQ 
      BackColor       =   &H00FFC0C0&
      Caption         =   "������ϸ����"
      Height          =   285
      Left            =   13860
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   7620
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRN 
      Height          =   585
      Left            =   5220
      TabIndex        =   53
      Top             =   7920
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1032
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtLyf 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8460
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   7890
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   8190
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4410
      Top             =   8190
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   585
      Left            =   14490
      Picture         =   "frmRen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8580
      Width           =   675
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "ȫ��"
      Height          =   285
      Left            =   6840
      TabIndex        =   33
      Top             =   8850
      Width           =   945
   End
   Begin VB.Frame frmTj 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   90
      TabIndex        =   26
      Top             =   8730
      Width           =   6555
      Begin VB.ComboBox comBj 
         Height          =   300
         ItemData        =   "frmRen.frx":0102
         Left            =   3120
         List            =   "frmRen.frx":010C
         TabIndex        =   29
         Text            =   "="
         Top             =   150
         Width           =   825
      End
      Begin VB.ComboBox comLx 
         Height          =   300
         ItemData        =   "frmRen.frx":0119
         Left            =   990
         List            =   "frmRen.frx":0126
         TabIndex        =   28
         Top             =   150
         Width           =   1485
      End
      Begin VB.ComboBox txtZ 
         Height          =   300
         ItemData        =   "frmRen.frx":013C
         Left            =   4500
         List            =   "frmRen.frx":013E
         TabIndex        =   27
         Top             =   120
         Width           =   1965
      End
      Begin VB.Label Label4 
         Caption         =   "�Ƚ�:"
         Height          =   225
         Left            =   2610
         TabIndex        =   32
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "ֵ:"
         Height          =   255
         Left            =   4110
         TabIndex        =   31
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label5 
         Caption         =   "��ѯ���:"
         Height          =   255
         Left            =   30
         TabIndex        =   30
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "��   ѯ"
      Height          =   315
      Left            =   7920
      TabIndex        =   24
      Top             =   8850
      Width           =   1095
   End
   Begin VB.Frame frmMod 
      Caption         =   "��������"
      Height          =   7905
      Left            =   10230
      TabIndex        =   6
      Top             =   0
      Width           =   4995
      Begin VB.TextBox txtLT 
         Height          =   270
         Left            =   1620
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   6090
         Width           =   1425
      End
      Begin VB.TextBox txtTT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3600
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdXZ 
         Caption         =   "ѡ  ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         TabIndex        =   48
         Top             =   4590
         Width           =   1065
      End
      Begin VB.CheckBox optZZF 
         Caption         =   "��ְ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   45
         Top             =   5580
         Width           =   885
      End
      Begin VB.CheckBox chkHGF 
         Caption         =   "ת֤"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   43
         Top             =   6030
         Width           =   825
      End
      Begin VB.CheckBox chkFyF 
         Caption         =   "���ý���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2850
         TabIndex        =   41
         ToolTipText     =   "���ž���ҵ��Ա�����̲��鳤�蹴ѡ���������������������"
         Top             =   6570
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox txtTang 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -1050
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   6900
         Width           =   3075
      End
      Begin MSComCtl2.DTPicker txtOld 
         Height          =   375
         Left            =   1620
         TabIndex        =   39
         Top             =   1890
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   8454016
         CalendarTitleBackColor=   16744576
         CalendarTrailingForeColor=   12583104
         Format          =   135790592
         CurrentDate     =   29295
      End
      Begin VB.ComboBox comQy 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1620
         TabIndex        =   36
         Top             =   3060
         Width           =   3075
      End
      Begin VB.ComboBox comXb 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRen.frx":0140
         Left            =   1620
         List            =   "frmRen.frx":014A
         TabIndex        =   22
         Text            =   "comXb"
         Top             =   1380
         Width           =   3135
      End
      Begin VB.TextBox txtZh 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1620
         TabIndex        =   21
         Top             =   2370
         Width           =   3075
      End
      Begin VB.TextBox txtZw 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1620
         TabIndex        =   20
         Top             =   4050
         Width           =   3075
      End
      Begin VB.TextBox txtYwy 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1620
         TabIndex        =   19
         Top             =   840
         Width           =   3075
      End
      Begin VB.TextBox txtNx 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   5040
         Width           =   915
      End
      Begin VB.TextBox txtUid 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   330
         Width           =   3075
      End
      Begin VB.OptionButton optZZFM 
         Caption         =   "��ְ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   16
         Top             =   7200
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSDataListLib.DataCombo comBm 
         Height          =   390
         Left            =   1620
         TabIndex        =   35
         Top             =   3570
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   688
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtFT 
         Height          =   270
         Left            =   1620
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   5670
         Width           =   1425
      End
      Begin MSComCtl2.DTPicker dtpFT 
         Height          =   255
         Left            =   1620
         TabIndex        =   60
         Top             =   5670
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   123207681
         CurrentDate     =   40807
      End
      Begin MSComCtl2.DTPicker dtpLT 
         Height          =   255
         Left            =   1620
         TabIndex        =   62
         Top             =   6090
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   123207681
         CurrentDate     =   40807
      End
      Begin VB.Label Label16 
         Caption         =   "��ְ����"
         Height          =   255
         Left            =   450
         TabIndex        =   58
         Top             =   6120
         Width           =   765
      End
      Begin VB.Label Label15 
         Caption         =   "��ְ����"
         Height          =   315
         Left            =   420
         TabIndex        =   57
         Top             =   5670
         Width           =   765
      End
      Begin VB.Label Label14 
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         TabIndex        =   55
         Top             =   5100
         Width           =   705
      End
      Begin VB.Label lblGGL 
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   1620
         TabIndex        =   47
         Top             =   4590
         Width           =   1875
      End
      Begin VB.Label Label13 
         Caption         =   "�ϼ�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   180
         TabIndex        =   46
         Top             =   4620
         Width           =   1215
      End
      Begin VB.Label Label9 
         Height          =   405
         Left            =   1860
         TabIndex        =   42
         Top             =   6570
         Width           =   2475
      End
      Begin VB.Label lblGzu 
         Caption         =   "lblGzu"
         Height          =   225
         Left            =   570
         TabIndex        =   34
         Top             =   7500
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblBmid 
         Caption         =   "lblBmid"
         Height          =   285
         Left            =   2370
         TabIndex        =   25
         Top             =   7020
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label10 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   15
         Top             =   5130
         Width           =   1005
      End
      Begin VB.Label Label8 
         Caption         =   "ְ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   14
         Top             =   4140
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   13
         Top             =   3630
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   3105
         Width           =   1035
      End
      Begin VB.Label lblzh 
         Caption         =   "���֤"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   11
         Top             =   2445
         Width           =   855
      End
      Begin VB.Label lblaaa 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   1935
         Width           =   1155
      End
      Begin VB.Label txtXb 
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   1425
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   390
         Width           =   975
      End
   End
   Begin VB.Frame frmAn 
      Height          =   795
      Left            =   12210
      TabIndex        =   3
      Top             =   8460
      Width           =   2175
      Begin VB.CommandButton cmdAdd 
         Caption         =   "���"
         Height          =   555
         Left            =   810
         Picture         =   "frmRen.frx":0156
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton cmdMod 
         Caption         =   "�޸�"
         Height          =   555
         Left            =   90
         Picture         =   "frmRen.frx":0598
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�ύ"
         Height          =   585
         Left            =   1470
         Picture         =   "frmRen.frx":08A2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   705
      End
   End
   Begin VB.Frame frmZzf 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   30
      TabIndex        =   1
      Top             =   7980
      Width           =   2895
      Begin VB.CheckBox chkHH 
         Caption         =   "ת֤"
         Height          =   255
         Left            =   1110
         TabIndex        =   44
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton opt3 
         Caption         =   "��ְ"
         Height          =   195
         Left            =   1980
         TabIndex        =   23
         Top             =   150
         Width           =   855
      End
      Begin VB.OptionButton opt1 
         Caption         =   "��ְ"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   150
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRen 
      Height          =   7875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   13891
      _Version        =   393216
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSDataListLib.DataCombo comGzu 
      Height          =   390
      Left            =   8460
      TabIndex        =   49
      Top             =   8310
      Visible         =   0   'False
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   688
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      Caption         =   "���η�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7050
      TabIndex        =   52
      Top             =   7995
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label12 
      Caption         =   "���̲����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7050
      TabIndex        =   51
      Top             =   8385
      Visible         =   0   'False
      Width           =   1425
   End
End
Attribute VB_Name = "frmRen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoRen As Object
Dim adoXR As Object
Dim adoGz As Object
Dim adoBm As Object
Public ZZF As Boolean
Public Bm As String
Dim timZm As Integer  '1��Ŀ�ƽ�

Private Sub chkAll_Click()
If chkAll.Value = 1 Then
    frmTj.Enabled = False
    opt1.Value = False
    opt3.Value = False
    frmZzf.Enabled = False
Else
    frmTj.Enabled = True
    comLx.Text = ""
    txtZ.Text = ""
    opt1.Value = True
    frmZzf.Enabled = True
End If
End Sub

Private Sub cmdAdd_Click()
Call RenQing
frmMod.Enabled = True
cmdAdd.Enabled = False
cmdMod.Enabled = False
cmdSave.Enabled = True
'optZZF.Value = True
chkHGF.Value = 0 'ת֤
txtNx.Text = 0
txtLyf.Text = 0
txtYwy.SetFocus
cmdXZ.Visible = True
chkFyF.Value = 1
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
frmZu.ZOrder 0
comLx.Text = ""
txtZ.Text = ""
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Private Sub cmdMod_Click()
frmMod.Enabled = True
cmdSave.Enabled = True
cmdXZ.Visible = True
End Sub

Private Sub cmdRef_Click()
Dim oo As Integer
Dim tt As String
Dim Qy As String
Dim ZZF As Integer
Dim hgF As Integer
On Error Resume Next
'�����Ա,ֻ�ܿ����Լ��������
If mod1.Qy <> "�Ϻ�" Then
    If comLx.Text = "����" And mod1.Qy <> txtZ.Text Then
        Exit Sub
    End If
    Qy = mod1.Qy
End If
If chkAll.Value = 1 Then
    If mod1.Qy = "�Ϻ�" Then
        tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where username<>'������' order by userid"
    Else
        tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where username<>'������' and qy='" & mod1.Qy & "' order by userid"
    End If
Else
    If opt1.Value = True Then
        ZZF = 1
        If chkHH.Value = 1 Then
            hgF = 1
        Else
            hgF = 0
        End If
    Else
        ZZF = 0
    End If
    Select Case comLx.Text
    Case "����"
        tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=" & ZZF & " and qy='" & txtZ.Text & "' order by userid"
    Case "����"
        If mod1.Qy = "�Ϻ�" Then
            tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=" & ZZF & " and bm='" & txtZ.Text & "' order by userid"
'''''            If txtZ.Text = "��������" Then
'''''                tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=" & ZZF & " and bm='" & txtZ.Text & "' or bm='����' order by bm,userid"
'''''            End If
'''''            If txtZ.Text = "ά����" Then
'''''                tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=" & ZZF & " and bm='" & txtZ.Text & "' or bm='ά����1' or bm='ά����2' or bm='�Ͼ���' or bm='���ݰ�' or bm='������' order by bm,userid"
'''''            End If
        
        Else
            tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=" & ZZF & " and qy='" & mod1.Qy & "' and bm='" & txtZ.Text & "' order by userid"
        End If
    
    Case "����"
        If mod1.Qy = "�Ϻ�" Then
            tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=" & ZZF & " and username like '%" & txtZ.Text & "%' order by userid"
        Else
            tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=" & ZZF & " and qy='" & mod1.Qy & "' and username like '%" & txtZ.Text & "%'  order by userid"
        End If
    Case Else
        If mod1.Qy = "�Ϻ�" Then
            tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=" & ZZF & " and hgf=" & hgF & " order by userid"
        Else
            tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=" & ZZF & " and qy='" & mod1.Qy & "' and hgf=" & hgF & " order by userid"
        End If
    End Select
End If
''''''''''''''frmRen.adoRen.Close
''''''''''''''frmRen.adoRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''''''''Set frmRen.dtgRen.DataSource = frmRen.adoRen
''''''''''''''If adoRen.RecordCount > 0 Then
''''''''''''''    Set frmRen.dtgRen.DataSource = frmRen.adoRen
'''''''''''''''    dtgCGj.FixedRows = 1
'''''''''''''''    dtgCGj.Row = adoGCJ.RecordCount - 1
''''''''''''''Else
''''''''''''''
''''''''''''''    dtgRen.Rows = 2
''''''''''''''    dtgRen.FixedRows = 1
''''''''''''''    dtgRen.Row = 1
''''''''''''''    For oo = 0 To 10
''''''''''''''        dtgRen.Col = oo
''''''''''''''        dtgRen.Text = ""
''''''''''''''    Next
''''''''''''''End If
''''''''''''''Call renQing
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
La = UBound(Ra, 2) + 1
mod1.HTP.Close
Set mod1.HTP = Nothing
Call frmRen.RenQing
frmRen.dtgRen.Clear: frmRen.dtgRN.Clear
frmRen.dtgRen.Visible = False
frmRen.dtgRen.Rows = La + 30: frmRen.dtgRen.Cols = 10
frmRen.dtgRen.Row = 0: frmRen.dtgRen.Col = 1: frmRen.dtgRen.Text = "����": frmRen.dtgRen.Col = 2: frmRen.dtgRen.Text = "����"
frmRen.dtgRen.Col = 3: frmRen.dtgRen.Text = "����": frmRen.dtgRen.Col = 4: frmRen.dtgRen.Text = "����":
frmRen.dtgRen.Col = 5: frmRen.dtgRen.Text = "ְ��": frmRen.dtgRen.Col = 6: frmRen.dtgRen.Text = "��������"
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
End Sub

Private Sub cmdSave_Click()
Dim KZF As Boolean '��ְת����Ա�Ƿ��ܱ���
Dim cmd As Object
Dim tt As String
Dim ERRch As String
On Error Resume Next
KZF = True
'��ְת�ڼ��
'''''''''If (BM <> comBm.Text Or optZZF.Value = 0) And txtUid.Text <> "" Then
'''''''''    tt = "select count(uid) from newfuwu where uid='" & txtUid.Text & "' and cf=0"
'''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''    If mod1.HTP.Fields(0).Value > 0 Then
'''''''''
'''''''''        MsgBox "������" & mod1.HTP.Fields(0).Value & "��������Ҫ����"
'''''''''        KZF = False
'''''''''
'''''''''    End If
'''''''''    tt = "select count(uid) from xmzl where uid='" & txtUid.Text & "'"
'''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''    If mod1.HTP.Fields(0).Value > 0 Then
    timZm = 1 '��Ŀ�ƽ�
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "���µ���"
        mod1.cmd.Parameters("@NBLX") = "��Ŀ�ƽ�"
        'mod1.cmd.Parameters("@NBLX") = "������"
        mod1.cmd.Parameters("@bh") = "��ǹ"
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = txtYwy.Text
        mod1.cmd.Parameters("@mt2") = txtUid.Text
        mod1.cmd.Parameters("@mt3") = comBm.Text '����
        mod1.cmd.Parameters("@mt4") = txtUid.Text
        mod1.cmd.Parameters("@mt5") = txtYwy.Text
        mod1.cmd.Parameters("@mt6") = comXb.Text
        mod1.cmd.Parameters("@mt7") = txtZh.Text
        mod1.cmd.Parameters("@mt8") = comQy.Text
        mod1.cmd.Parameters("@mt9") = comBm.Text
        mod1.cmd.Parameters("@mt10") = txtZw.Text
        mod1.cmd.Parameters("@mt11") = lblGGL.ToolTipText
        mod1.cmd.Parameters("@mt13") = txtTT.Text '������
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = 0
        mod1.cmd.Parameters("@mb1") = chkFyF.Value
        mod1.cmd.Parameters("@mb2") = chkHGF.Value
        mod1.cmd.Parameters("@mb3") = optZZF.Value
        mod1.cmd.Parameters("@mb4") = 0
        mod1.cmd.Parameters("@mb5") = 0
        mod1.cmd.Parameters("@md1") = txtTang.Text
        mod1.cmd.Parameters("@md2") = txtFT.Text
        mod1.cmd.Parameters("@md3") = txtLT.Text
   Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute



''''''''''''''cmd.Parameters("@hgf") = chkHGF.Value 'ת֤
''''''''''''''cmd.Parameters("@ggl") = lblGGL.ToolTipText '�ϼ��˵Ĺ���
''''''''''''''If optZZF.Value = 1 Then
''''''''''''''    cmd.Parameters("@zzf") = True
''''''''''''''Else
''''''''''''''    cmd.Parameters("@zzf") = False
''''''''''''''End If
''''''''''''''cmd.Parameters("@date") = DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda))
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
            MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
            cmdDing.Enabled = False
            Exit Sub
        Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
    
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
        End If
    
        
'''''''''''    Set mod1.cmd = Nothing
'''''''''''        MsgBox "������" & mod1.HTP.Fields(0).Value & "��Ŀδ�ƽ���"
'''''''''''        KZF = False
'''''''''''    End If
'''''''''''    tt = "select count(uid) from fyd where uid='" & txtUid.Text & "' and qrq is null and hg>0"
'''''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''''    If mod1.HTP.Fields(0).Value > 0 Then
'''''''''''        MsgBox "������" & mod1.HTP.Fields(0).Value & "�ű�����δǩ�գ�"
'''''''''''        KZF = False
'''''''''''    End If
'''''''''''End If
'If KZF = False Then
'    MsgBox "�����ݲ��ܰ���ְ�������������Ա�����������Ϣ�е�δ������"
'    Exit Sub
'End If
''''''''''''''Set cmd = createobject("adodb.command")
''''''''''''''cmd.ActiveConnection = mod1.CC
''''''''''''''cmd.CommandText = "TXrenAdd"
''''''''''''''cmd.CommandType = adCmdStoredProc
''''''''''''''cmd.Parameters("@uid") = txtUid.Text
''''''''''''''cmd.Parameters("@username") = txtYwy.Text
''''''''''''''cmd.Parameters("@usersex") = comXb.Text
''''''''''''''cmd.Parameters("@userold") = txtOld.Value
''''''''''''''cmd.Parameters("@userbid") = txtZH.Text
''''''''''''''cmd.Parameters("@qy") = comQy.Text
''''''''''''''cmd.Parameters("@bm") = comBm.Text
''''''''''''''cmd.Parameters("@bmid") = lblBmid.Caption
''''''''''''''cmd.Parameters("@userzw") = txtZw.Text
''''''''''''''cmd.Parameters("@nx") = Val(txtNx.Text)
''''''''''''''cmd.Parameters("@lyf") = Val(txtLyf.Text)
''''''''''''''cmd.Parameters("@gzu") = Val(lblGzu.Caption)
''''''''''''''cmd.Parameters("@lren") = mod1.DName '¼����
''''''''''''''cmd.Parameters("@luid") = mod1.DHid
''''''''''''''cmd.Parameters("@fyf") = chkFyF.Value
''''''''''''''cmd.Parameters("@hgf") = chkHGF.Value 'ת֤
''''''''''''''cmd.Parameters("@ggl") = lblGGL.ToolTipText '�ϼ��˵Ĺ���
''''''''''''''If optZZF.Value = 1 Then
''''''''''''''    cmd.Parameters("@zzf") = True
''''''''''''''Else
''''''''''''''    cmd.Parameters("@zzf") = False
''''''''''''''End If
''''''''''''''cmd.Parameters("@date") = DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda))
''''''''''''''cmd.Parameters("@errch") = ""
''''''''''''''cmd.Execute
''''''''''''''ERRch = cmd.Parameters("@errch").Value


''''''''''''''If ERRch <> "�ɹ�" Then
''''''''''''''        MsgBox "������ֹ���,������һ��,��������ύ���ɹ�,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"


''''''''''''''''''        Exit Sub
''''''''''''''''''End If
''''''''''''''''''txtUid.Text = cmd.Parameters("@uid").Value
''''''''''''''''''adoRen.Requery
''''''''''''''''''If adoRen.RecordCount > 0 Then
''''''''''''''''''    Set frmRen.dtgRen.DataSource = frmRen.adoRen
''''''''''''''''''Else
''''''''''''''''''
''''''''''''''''''    dtgRen.Rows = 2
''''''''''''''''''    dtgRen.FixedRows = 1
''''''''''''''''''    dtgRen.Row = 1
''''''''''''''''''    For oo = 0 To 10
''''''''''''''''''        dtgRen.Col = oo
''''''''''''''''''        dtgRen.Text = ""
''''''''''''''''''    Next
''''''''''''''''''End If
Set cmd = Nothing
frmMod.Enabled = False
cmdXZ.Visible = False
cmdSave.Enabled = False
cmdAdd.Enabled = True
cmdMod.Enabled = True
End Sub

Private Sub cmdXQ_Click()
Me.Enabled = False
frmRL.Show
Call frmRL.Qing
Call frmRL.Bound(txtUid.Text)
frmRL.ZOrder 0
End Sub

Private Sub cmdXZ_Click()
Set Ren.XForm = New frmRen
Call mod1.RenXz("frmRen", Me, 0)
End Sub



Private Sub comGzu_Click(Area As Integer)
On Error Resume Next
If comGzu.Text = "������" Then
    comGzu.Text = ""
End If
lblGzu.Caption = comGzu.BoundText

End Sub


Private Sub comLx_Click()
Dim adoZ As Object
Dim oo As Integer
Dim tt As String
Set adoZ = CreateObject("adodb.recordset")
On Error Resume Next
Select Case comLx.Text
Case "����"
    For oo = 20 To 0 Step -1
        txtZ.RemoveItem oo
    Next
    tt = "select qy from yzqy order by qid"
    adoZ.Close
    adoZ.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
    adoZ.MoveFirst
    Do While Not adoZ.EOF
        txtZ.AddItem adoZ.Fields("qy").Value
        adoZ.MoveNext
    Loop
    adoZ.Close
    Set adoZ = Nothing
    comBj.Text = "="
Case "����"
    For oo = 20 To 0 Step -1
        txtZ.RemoveItem oo
    Next
    'tt = "select bm from worker where zzf=1 group by bm,bmid,zzf order by bmid"
    tt = "select bm from bm order by bmid"
    adoZ.Close
    adoZ.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
    adoZ.MoveFirst
    Do While Not adoZ.EOF
        txtZ.AddItem adoZ.Fields("bm").Value
        adoZ.MoveNext
    Loop
    adoZ.Close
    Set adoZ = Nothing
    comBj.Text = "="
Case "����"
    comBj.Text = "����"
End Select
txtZ.Text = ""
End Sub

Private Sub dtgRen_DblClick()
Call RenQing
dtgRen.Col = 1

Call RenBound(dtgRen.Text)
End Sub


Private Sub dtgRen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call RenQing
    dtgRen.Col = 1
    Call RenBound(dtgRen.Text)
End If
End Sub


Private Sub dtpFT_CloseUp()
txtFT.Text = dtpFT.Value
End Sub


Private Sub dtpLT_CloseUp()
txtLT.Text = dtpLT.Value
End Sub


Private Sub Form_Load()
Dim adoZ As Object

Dim oo As Integer
Dim tt As String
Set adoZ = CreateObject("adodb.recordset")
On Error Resume Next
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
Set adoRen = CreateObject("adodb.recordset")
Set adoXR = CreateObject("adodb.recordset")
dtgRen.ColWidth(0) = 300
dtgRen.ColWidth(4) = 2000
dtgRen.ColWidth(5) = 2000
txtTang.Left = 1620
txtTang.Top = 1890
    For oo = 20 To 0 Step -1
        comQy.RemoveItem oo
    Next
    tt = "select qy from yzqy order by qid"
    adoZ.Close
    adoZ.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
    adoZ.MoveFirst
    Do While Not adoZ.EOF
        comQy.AddItem adoZ.Fields("qy").Value
        adoZ.MoveNext
    Loop
    adoZ.Close
    Set adoZ = Nothing
    Set adoBm = CreateObject("adodb.recordset")
    tt = "select bm,bmid from bm where zzf=1  order by bmid"
    adoBm.Close
    adoBm.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
    Set comBm.RowSource = adoBm
    comBm.ListField = "bm"
    comBm.BoundColumn = "bmid"
    
frmAn.Visible = True
dtpFT.Value = mod1.DQda
dtpLT.Value = mod1.DQda
    
'���̲������Ͽ�
tt = "select username,gzu from worker where zzf=1 and zuf=1 order by gzu"
Set adoGz = CreateObject("adodb.recordset")
adoGz.Close
adoGz.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
Set comGzu.RowSource = adoGz
comGzu.ListField = "username"
comGzu.BoundColumn = "gzu"
If mod1.DName <> "������" And mod1.DName <> "��ɺɺ" And mod1.DName <> "����ƽ" And mod1.DName <> "����" Then
    frmAn.Visible = False
End If

If mod1.DName = "����" Or mod1.DName = "������" Or mod1.DName = "������" Or mod1.DName = "����ƽ" Or mod1.DName = "��ɺɺ" Then
    cmdXQ.Visible = True
Else
    cmdXQ.Visible = False
End If

End Sub

Public Sub RenQing()
txtUid.Text = ""
txtYwy.Text = ""
comXb.Text = ""
'txtOld.Value = ""
txtZh.Text = ""
comQy.Text = ""
comBm.Text = ""
txtZw.Text = ""
txtNx.Text = ""
txtLyf.Text = ""
comGzu.Text = ""
optZZF.Value = 0
txtTang.Visible = True
chkFyF.Value = 0
ZZF = False
Bm = ""
lblGGL.Caption = ""
lblGGL.ToolTipText = ""
txtTang.Text = ""
txtTang.Visible = True
txtFT.Text = ""
txtLT.Text = ""
txtTT.Text = ""
End Sub

Private Sub opt1_Click()
'Dim tt As String
'On Error Resume Next
'If opt1.Value = False Then Exit Sub
'tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=1 and username<>'������' order by userid"
'adoRen.Close
'adoRen.Open tt, mod1.workkk, adOpenKeyset, adLockReadOnly, adCmdText
'Set dtgRen.DataSource = adoRen
'Call renQing
chkHH.Value = 1
End Sub
Public Sub RenBound(Uid As String)
Dim tt As String
Dim Ra
Dim Rb
Dim RC
Dim RD
tt = "Declare @bm nvarchar(20),@GGL nvarchar(20);" & _
    "select @bm=bm,@ggl=ggl from worker where userid='" & Uid & "';" & _
    "select userid,username,usersex,userold,userbid,qy,bm,userzw,lyf,zzf,hgf,ggl,ft,lt,tt from worker where userid='" & Uid & "';" & _
    "select bmid from bm where bm=@bm;" & _
    "select username from worker where userid=@ggl;" & _
    "select Bdate,getdate() from rla where auid='" & Uid & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RD = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing

txtUid.Text = Ra(0, 0)
txtYwy.Text = Ra(1, 0)
comXb.Text = Ra(2, 0)
txtTang.Text = Ra(3, 0)
txtZh.Text = Ra(4, 0)
comQy.Text = Ra(5, 0)
comBm.Text = Ra(6, 0)
Bm = Ra(6, 0)
lblBmid.Caption = Rb(0, 0)
txtZw.Text = Ra(7, 0)
txtLyf.Text = Ra(8, 0)
If Ra(10, 0) = True Then
    chkHGF.Value = 1
Else
    chkHGF.Value = 0
End If
If Ra(9, 0) = True Then
    optZZF.Value = 1
Else
    optZZF.Value = 0
End If
lblGGL.ToolTipText = Ra(11, 0)
lblGGL.Caption = RC(0, 0)
ZZF = Ra(9, 0)
txtFT.Text = Ra(12, 0)
txtLT.Text = Ra(13, 0)
txtTT.Text = Ra(14, 0)

txtNx.Text = Round(DateDiff("yyyy", RD(0, 0), RD(1, 0)), 1)
cmdMod.Enabled = True
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdXZ.Visible = False
End Sub

Private Sub opt2_Click()
'Dim tt As String
'On Error Resume Next
'If opt2.Value = False Then Exit Sub
'tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where username<>'������' order by userid"
'adoRen.Close
'adoRen.Open tt, mod1.workkk, adOpenKeyset, adLockReadOnly, adCmdText
'Set dtgRen.DataSource = adoRen
'Call renQing
End Sub

Private Sub opt3_Click()
'Dim tt As String
'On Error Resume Next
'If opt3.Value = False Then Exit Sub
'tt = "select userid as ����,username as ����,qy as ����,bm as ����,userzw as ְ��,nx as �������� from worker where zzf=0 and username<>'������' order by userid"
'adoRen.Close
'adoRen.Open tt, mod1.workkk, adOpenKeyset, adLockReadOnly, adCmdText
'Set dtgRen.DataSource = adoRen
'Call renQing

End Sub

Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '�����༭
'''''MsgBox "�Ѿ��ɹ�֪ͨ���۾���ת�ƴ��˵���Ŀ��"

End If

timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '�ύ�ɹ�
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0

   If timZm = 1 Then '��Ŀ�ƽ�
                
    txtUid.Text = mod1.WP.Fields("mt2").Value

        

        

        
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then 'δ���

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '����ʧ��
    timWait.Enabled = False
    ii = MsgBox("���������ڴ�����������ʱ,�������´���:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0

    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("���������ڴ�����������ʱ,��ʱ!", vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0

    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub

Private Sub txtOld_CloseUp()
txtTang.Text = txtOld.Value
txtOld.Visible = False
txtTang.Visible = True

End Sub


Private Sub txtTang_Click()
txtTang.Visible = False
txtOld.Visible = True
End Sub


