VERSION 5.00
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form frmTip 
   BackColor       =   &H00C0FFC0&
   Caption         =   "2012年度奖项说明"
   ClientHeight    =   5475
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7695
   StartUpPosition =   2  '屏幕中心
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin NiceFormControl.NiceButton NiceButton1 
      Height          =   345
      Left            =   6210
      TabIndex        =   3
      Top             =   4710
      Width           =   1185
      _ExtentX        =   2090
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
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTip.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Style           =   7
      Caption         =   "确定"
   End
   Begin VB.TextBox txtNr 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmTip.frx":001C
      Top             =   870
      Width           =   7545
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      BackColor       =   &H00C0FFC0&
      Caption         =   "在启动时显示提示(&S)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   4830
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin NiceFormControl.NiceButton NiceButton2 
      Height          =   345
      Left            =   3900
      TabIndex        =   4
      Top             =   4710
      Width           =   1095
      _ExtentX        =   1931
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
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTip.frx":0022
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Style           =   7
      Caption         =   "<"
   End
   Begin NiceFormControl.NiceButton NiceButton3 
      Height          =   345
      Left            =   5010
      TabIndex        =   5
      Top             =   4710
      Width           =   1095
      _ExtentX        =   1931
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
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTip.frx":003E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Style           =   7
      Caption         =   ">"
   End
   Begin VB.Label lblBt 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "2012年度奖项说明"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   330
      Width           =   5805
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 内存中的提示数据库。
Dim Tips As New Collection

' 提示文件名称
Const TIP_FILE = "TIPOFDAY.TXT"

' 当前正在显示的提示集合的索引。
Dim CurrentTip As Long
Dim Nid As Long
Dim Ra

Dim La As Integer

Private Sub DoNextTip()

'    ' 随机选择一条提示。
'    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' 或者，您可以按顺序遍历提示

    CurrentTip = CurrentTip + 1
    If Tips.Count < CurrentTip Then
        CurrentTip = 1
    End If
    
    ' 显示它。
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' 从文件中读出的每条提示。
    Dim InFile As Integer   ' 文件的描述符。
    
    ' 包含下一个自由文件描述符。
    InFile = FreeFile
    
    ' 确定为指定文件。
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' 在打开前确保文件存在。
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' 从文本文件中读取集合。
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' 随机显示一条提示。
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()
    ' 保存在下次启动时是否显示此窗体
'    SaveSetting App.EXEName, "Options", "在启动时显示提示", chkLoadTipsAtStartup.Value
Dim tt As String
On Error Resume Next
tt = "update worker set txf =" & chkLoadTipsAtStartup.Value & " where userid='" & mod1.DHid & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim tt As String

    Dim ShowAtStartup As Long
'''NF.LoadSkin 10
'''NF.AutoSkinControl
    
    ' 察看在启动时是否将被显示
    ShowAtStartup = GetSetting(App.EXEName, "Options", "在启动时显示提示", 1)
    If mod1.ZT = "HMData" Then ShowAtStartup = 2
    If ShowAtStartup = 0 Then
        Unload Me
        Exit Sub
    End If
  
    ' 设置复选框，强行将值写回到注册表
    Me.chkLoadTipsAtStartup.Value = vbChecked

    
    
'''''''    ' 随机寻找
'''''''    Randomize
'''''''
'''''''     '读取提示文件并且随机显示一条提示?
'''''''    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
'''''''        lblTipText.Caption = "文件 " & TIP_FILE & " 没有被找到吗? " & vbCrLf & vbCrLf & _
'''''''           "创建文本文件名为 " & TIP_FILE & " 使用记事本每行写一条提示。 " & _
'''''''           "然后将它存放在应用程序所在的目录 "
'''''''    End If
    tt = "select bt,nr,nid from qywh order by nid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2) + 1
    lblBt.Caption = Ra(0, 0)
    txtNr.Text = Ra(1, 0)
    Nid = Ra(2, 0)
'''''    frmTip.Show
'''''    frmTip.ZOrder 0
End Sub

Public Sub DisplayCurrentTip()
'''''    If Tips.Count > 0 Then
'''''        lblTipText.Caption = Tips.Item(CurrentTip)
'''''    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmZu.TBa.Buttons(5).Value = tbrUnpressed
End Sub

Private Sub NiceButton1_Click()
    Unload Me
    frmZu.TBa.Buttons(5).Value = tbrUnpressed
End Sub


Private Sub NiceButton2_Click()

On Error Resume Next
Nid = Nid - 1
If Nid = 0 Then Nid = 1
    lblBt.Caption = Ra(0, Nid - 1)
    txtNr.Text = Ra(1, Nid - 1)
    Nid = Ra(2, Nid - 1)
End Sub

Private Sub NiceButton3_Click()
On Error Resume Next
Nid = Nid + 1
    lblBt.Caption = Ra(0, Nid - 1)
    txtNr.Text = Ra(1, Nid - 1)
    Nid = Ra(2, Nid - 1)
End Sub


