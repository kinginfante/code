VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form FmxcLx 
   BackColor       =   &H00C0FFFF&
   Caption         =   "ҵ������"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   7605
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2910
      Top             =   4110
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   4050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLx 
      Height          =   4065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7170
      _Version        =   393216
      BackColor       =   12648447
      Rows            =   14
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648447
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin NiceFormControl.NiceButton cmdNew 
      Height          =   345
      Left            =   4170
      TabIndex        =   2
      Top             =   4200
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   609
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "FmxcLx.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "����ѯ�۵�"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�������˫����Ӧ����Ŀ"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   4260
      Width           =   2535
   End
End
Attribute VB_Name = "FmxcLx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LX As String
Private Sub cmdNew_Click()
If cmdNew.Caption = "����ѯ�۵�" Then
    MsgBox "��ѡ����Ӧ��ҵ������!(���б���˫��)"
    Exit Sub
End If
Call FMXCXmmc.Qing
FMXCXmmc.Show
FMXCXmmc.ZOrder 0
FMXCXmmc.Lb = "ѯ�۵�"
FMXCXmmc.NiceButton1.Caption = "�� �� �� �� (ѯ�۵�)"
End Sub

Private Sub dtgLx_DblClick()

If dtgLx.Row = 0 Then Exit Sub
dtgLx.Col = 1: LX = dtgLx.Text
cmdNew.Caption = "����" & LX & "ѯ�۵�"
cmdNew.ToolTipText = dtgLx.Row

End Sub

Private Sub Form_Load()
Me.Width = 7725
Me.Height = 5190

dtgLx.Row = 0
dtgLx.Col = 0: dtgLx.Text = "ҵ������"
dtgLx.Col = 1: dtgLx.Text = "ҵ������"
dtgLx.Col = 2: dtgLx.Text = "��׼�۸�"
dtgLx.Col = 3: dtgLx.Text = "�ٴ���"
dtgLx.Col = 4: dtgLx.Text = "ѯ�۵�"
dtgLx.Col = 5: dtgLx.Text = "��ͬ���"
dtgLx.Col = 6: dtgLx.Text = "˵��"
dtgLx.MergeCells = flexMergeFree
dtgLx.MergeRow(0) = True
dtgLx.Row = 1: dtgLx.Col = 0: dtgLx.Text = "�˹���"
dtgLx.Row = 1: dtgLx.Col = 1: dtgLx.Text = "ά��": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 1: dtgLx.Col = 6: dtgLx.Text = "����˾��Ա������ɵ��˹�"
dtgLx.Row = 2: dtgLx.Col = 0: dtgLx.Text = "�˹���"
dtgLx.Row = 2: dtgLx.Col = 1: dtgLx.Text = "����": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 2: dtgLx.Col = 6: dtgLx.Text = "����˾��Ա������ɵ��˹�"
dtgLx.Row = 3: dtgLx.Col = 0: dtgLx.Text = "�˹���"
dtgLx.Row = 3: dtgLx.Col = 1: dtgLx.Text = "�����˹�": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 3: dtgLx.Col = 6: dtgLx.Text = "����˾��Ա������ɵ��˹�"
dtgLx.Row = 4: dtgLx.Col = 0: dtgLx.Text = "ѹ����"
dtgLx.Row = 4: dtgLx.Col = 1: dtgLx.Text = "ѹ����ά�ޱ���": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 4: dtgLx.Col = 6: dtgLx.Text = "ѹ����������ά�޻���"
dtgLx.Row = 5: dtgLx.Col = 0: dtgLx.Text = "ѹ����"
dtgLx.Row = 5: dtgLx.Col = 1: dtgLx.Text = "ѹ����ó��": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 5: dtgLx.Col = 6: dtgLx.Text = "ѹ���������Ĳ�Ʒ����"
dtgLx.Row = 6: dtgLx.Col = 0: dtgLx.Text = "�н�"
dtgLx.Row = 6: dtgLx.Col = 1: dtgLx.Text = "�н�ҵ��"
dtgLx.Row = 6: dtgLx.Col = 6: dtgLx.Text = "�н飨�Ӽ䣩ҵ������"
dtgLx.Row = 7: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 7: dtgLx.Col = 1: dtgLx.Text = "����": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 7: dtgLx.Col = 6: dtgLx.Text = "�����豸��ó��"
dtgLx.Row = 8: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 8: dtgLx.Col = 1: dtgLx.Text = "����": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 8: dtgLx.Col = 6: dtgLx.Text = "���ݽ�ʨ���������豸��ó��"
dtgLx.Row = 9: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 9: dtgLx.Col = 1: dtgLx.Text = "�ڴ︻": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 9: dtgLx.Col = 6: dtgLx.Text = "�ڴ︻�豸��ó��"
dtgLx.Row = 10: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 10: dtgLx.Col = 1: dtgLx.Text = "��ͼ": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 10: dtgLx.Col = 6: dtgLx.Text = "��ͼ�豸��ó��"
dtgLx.Row = 11: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 11: dtgLx.Col = 1: dtgLx.Text = "�����": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 11: dtgLx.Col = 6: dtgLx.Text = "����������������׺ģ���ó��"
dtgLx.Row = 12: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 12: dtgLx.Col = 1: dtgLx.Text = "�ְ�"
dtgLx.Row = 12: dtgLx.Col = 6: dtgLx.Text = "�ְ���ͬ"
dtgLx.Row = 13: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 13: dtgLx.Col = 1: dtgLx.Text = "�Ǵ����Ʒ"
dtgLx.Row = 13: dtgLx.Col = 6: dtgLx.Text = "�Ǵ����Ʒ��ó��"
dtgLx.Col = 5
dtgLx.Row = 1: dtgLx.Text = "RG": dtgLx.Row = 2: dtgLx.Text = "RG": dtgLx.Row = 3: dtgLx.Text = "RG"
dtgLx.Row = 4: dtgLx.Text = "YS": dtgLx.Row = 5: dtgLx.Text = "YS"
dtgLx.Row = 6: dtgLx.Text = "ZJ"
dtgLx.Row = 7: dtgLx.Text = "TR": dtgLx.Row = 8: dtgLx.Text = "TR": dtgLx.Row = 8: dtgLx.Text = "TR": dtgLx.Row = 10: dtgLx.Text = "TR": dtgLx.Row = 11: dtgLx.Text = "TR"
dtgLx.Row = 12: dtgLx.Text = "TR": dtgLx.Row = 13: dtgLx.Text = "TR": dtgLx.Row = 9: dtgLx.Text = "TR"
dtgLx.MergeCol(5) = True
dtgLx.MergeCol(0) = True
dtgLx.ColWidth(1) = 1695
dtgLx.ColWidth(2) = 0
dtgLx.ColWidth(3) = 0
dtgLx.ColWidth(4) = 0
dtgLx.ColWidth(6) = 2925
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
End Sub


