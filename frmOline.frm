VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form frmOL 
   Caption         =   "您正和         交谈"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   8055
   Begin NiceFormControl.NiceForm NF 
      Left            =   6750
      Top             =   1590
      _ExtentX        =   847
      _ExtentY        =   847
      MnuStyleIdx     =   4
   End
   Begin MSComDlg.CommonDialog cmdA 
      Left            =   7320
      Top             =   1860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "传送文件"
      Height          =   315
      Left            =   4230
      TabIndex        =   10
      Top             =   3690
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Timer timOL 
      Interval        =   5000
      Left            =   7650
      Top             =   2550
   End
   Begin VB.Frame Frame1 
      Height          =   3645
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6465
      Begin VB.CommandButton cmdOld 
         BackColor       =   &H00C0FFC0&
         Caption         =   "聊天记录"
         Height          =   285
         Left            =   5490
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2250
         Width           =   915
      End
      Begin VB.ComboBox comCC 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         ItemData        =   "frmOline.frx":0000
         Left            =   1050
         List            =   "frmOline.frx":0025
         TabIndex        =   6
         Top             =   3270
         Width           =   3705
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "发送消息"
         Height          =   315
         Left            =   5490
         TabIndex        =   4
         Top             =   3270
         Width           =   915
      End
      Begin VB.TextBox txt2 
         BackColor       =   &H00C0FFFF&
         Height          =   675
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2550
         Width           =   6465
      End
      Begin VB.TextBox txt1 
         BackColor       =   &H00C0C0FF&
         Height          =   2175
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   90
         Width           =   6465
      End
      Begin VB.Label Label1 
         Caption         =   "常用语"
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   3330
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdBack 
      Height          =   435
      Left            =   7500
      Picture         =   "frmOline.frx":00A7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   6630
      TabIndex        =   8
      Top             =   3000
      Width           =   675
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   870
      Width           =   735
   End
   Begin VB.Image img2 
      Height          =   480
      Left            =   6660
      Picture         =   "frmOline.frx":01A9
      Top             =   2340
      Width           =   480
   End
   Begin VB.Image img1 
      Height          =   480
      Left            =   6660
      Picture         =   "frmOline.frx":05EB
      Top             =   270
      Width           =   480
   End
End
Attribute VB_Name = "frmOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MaxGid As Long
Dim Ol As Object
Private Sub cmdBack_Click()
Me.Visible = False
Unload Me
End Sub

Private Sub cmdFile_Click()
Dim Fsize As Long
Dim oo As Long
Dim ii As Long
Dim Yd As Long
    Dim bytData() As Byte
    On Error GoTo PKK
    cmdA.ShowOpen
    
 Open cmdA.FileName For Binary As #1
Fsize = LOF(1) - 1
Close #1
ii = UpInt(Fsize / 1000)
Yd = Fsize Mod 1000
For oo = 1 To ii
    bytData = ReadFile(cmdA.FileName, 1, 30000)
    Call WriteFile(App.Path & "\1.mp3", bytData)
    bytData = ReadFile(App.Path & "\music.mp3", 30001)
    Call WriteFile(App.Path & "\2.mp3", bytData)
Next

    Exit Sub
PKK:
    
End Sub

Private Sub cmdOld_Click()
Dim tt As String
Dim Ra
Dim La
Dim oo As Integer
On Error Resume Next

tt = "select frq,zz,nr,zuid from hmtext.dbo.qq where zuid='" & mod1.DHid & "' and tuid='" & lbl1.ToolTipText & "' or zuid='" & lbl1.ToolTipText & "' and tuid='" & mod1.DHid & _
    "'  order by gid "
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workHM, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2)

frmOR.dtgOld.Clear
frmOR.dtgOld.Rows = La + 40
For oo = 0 To La
    frmOR.dtgOld.Row = oo
    frmOR.dtgOld.Col = 3
    If Ra(3, oo) = mod1.DHid Then
        frmOR.dtgOld.Col = 0: frmOR.dtgOld.Text = Ra(0, oo): frmOR.dtgOld.CellForeColor = &HFF0000
        frmOR.dtgOld.Col = 1: frmOR.dtgOld.Text = Ra(1, oo): frmOR.dtgOld.CellForeColor = &HFF0000
        frmOR.dtgOld.Col = 2: frmOR.dtgOld.Text = Ra(2, oo): frmOR.dtgOld.CellForeColor = &HFF0000
        'frmOR.dtgOld.RowHeight(oo) = frmOR.dtgOld.CellHeight
    Else
        frmOR.dtgOld.Col = 0: frmOR.dtgOld.Text = Ra(0, oo): frmOR.dtgOld.CellForeColor = &H8080FF
        frmOR.dtgOld.Col = 1: frmOR.dtgOld.Text = Ra(1, oo): frmOR.dtgOld.CellForeColor = &H8080FF
        frmOR.dtgOld.Col = 2: frmOR.dtgOld.Text = Ra(2, oo): frmOR.dtgOld.CellForeColor = &H8080FF
        'frmOR.dtgOld.RowHeight(oo) = frmOR.dtgOld.CellHeight
    End If
Next
frmOR.Show
frmOR.dtgOld.Col = 0
frmOR.Caption = "您与" & lbl1.Caption & "的私聊记录:"

End Sub

Private Sub cmdSend_Click()
Dim tt As String
Dim Ra
On Error GoTo FERR
If txt2.Text = "" Then Exit Sub
tt = "insert into hmtext.dbo.qq (zz,zuid,tz,tuid,frq,nr) values ('" & _
    mod1.DName & "','" & mod1.DHid & "','" & lbl1.Caption & "','" & lbl1.ToolTipText & "',getdate(),'" & _
    txt2.Text & "');" & _
    "select @@identity"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workHM, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = mod1.HTP.NextRecordset
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing


txt1.Text = txt1.Text & mod1.DName & " (" & Hour(mod1.DQda) & ":" & Minute(mod1.DQda) & ":" & Second(mod1.DQda) & "):" & Chr(13) & Chr(10) & _
 "                  " & txt2.Text & Chr(13) & Chr(10)

txt1.SelStart = Len(txt1.Text)
txt1.SelLength = 0
MaxGid = Ra(0, 0)





txt2.Text = ""
    Exit Sub
FERR:
MsgBox "网络故障！再试一次！"
End Sub

Public Sub Tbound(Tuid As String)
Dim tt As String
Dim Ra
Dim La
Dim oo As Integer
On Error Resume Next
txt1.Text = ""
tt = "select zz,frq,nr,gid from hmtext.dbo.qq where (zuid='" & mod1.DHid & "' and tuid='" & Tuid & "' or zuid='" & Tuid & "' and tuid='" & mod1.DHid & _
    "') and frq>'" & DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda)) & "' order by gid ;" & _
        "update hmtext.dbo.qq set cf=1  where  zuid='" & Tuid & "' and tuid='" & mod1.DHid & "'"
Set mod1.TBD = CreateObject("adodb.recordset")
mod1.TBD.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.TBD.GetRows
mod1.TBD.Close
Set mod1.TBD = Nothing
La = UBound(Ra, 2) + 1
For oo = 1 To La
txt1.Text = txt1.Text & Ra(0, oo - 1) & " (" & Hour(Ra(1, oo - 1)) & ":" & Minute(Ra(1, oo - 1)) & ":" & Second(Ra(1, oo - 1)) & "):" & Chr(13) & Chr(10) & _
 "                  " & Ra(2, oo - 1) & Chr(13) & Chr(10)
Next
txt1.SelStart = Len(txt1.Text)
txt1.SelLength = 0
MaxGid = Ra(3, oo - 2)

End Sub

Private Sub comCC_Click()
txt2.Text = comCC.Text
End Sub

Private Sub cmdCut_Click()
    Dim bytData() As Byte
    bytData = ReadFile(App.Path & "\music.mp3", 1, 30000)
    Call WriteFile(App.Path & "\1.mp3", bytData)
    bytData = ReadFile(App.Path & "\music.mp3", 30001)
    Call WriteFile(App.Path & "\2.mp3", bytData)
End Sub

Private Function ReadFile(ByVal strFileName As String, Optional ByVal lngStartPos As Long = 1, Optional ByVal lngFileSize As Long = -1) As Byte()
    Dim FilNum As Integer
    FilNum = FreeFile
    Open strFileName For Binary As #FilNum
        If lngFileSize = -1 Then
            ReDim ReadFile(LOF(FilNum) - lngStartPos)
        Else
            ReDim ReadFile(lngFileSize - 1)
        End If
        Get #FilNum, lngStartPos, ReadFile
    Close #FilNum
End Function

Private Function WriteFile(ByVal strFileName As String, bytData() As Byte, Optional ByVal lngStartPos As Long = -1, Optional ByVal OverWrite As Boolean = True)
    Dim FilNum As Integer
    FilNum = FreeFile
    If OverWrite = True And Dir(strFileName) <> "" Then
        Kill strFileName
    End If
    Open strFileName For Binary As #FilNum
        If lngStartPos = -1 Then
            Put #FilNum, LOF(FilNum) + 1, bytData
        Else
            Put #FilNum, lngStartPos, bytData
        End If
    Close #FilNum
End Function

Private Sub Form_Load()
Me.Height = 4125
Me.Width = 8175
NF.LoadSkin 4
NF.AutoSkinControl
End Sub

Private Sub Form_Resize()
On Error Resume Next
NF.DoRndForm
Frame1.Width = Me.Width - 1700
txt1.Width = Frame1.Width
txt2.Width = Frame1.Width
End Sub

Private Sub timOL_Timer()
Dim tt As String
Dim Ra
Dim La As Integer
Dim oo As Integer
On Error Resume Next
La = 0
Set Ol = CreateObject("adodb.recordset")

tt = "select zz,frq,nr,gid from hmtext.dbo.qq where (zuid='" & mod1.DHid & "' and tuid='" & lbl1.ToolTipText & "' or zuid='" & lbl1.ToolTipText & "' and tuid='" & mod1.DHid & _
    "') and gid>" & MaxGid & " order by gid ;" & _
    "update hmtext.dbo.qq set cf=1  where  zuid='" & lbl1.ToolTipText & "' and tuid='" & mod1.DHid & "'"
Ol.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = Ol.GetRows
Ol.Close
Set Ol = Nothing
La = UBound(Ra, 2) + 1
If La = 0 Then Exit Sub
If La > 0 Then
For oo = 1 To La
txt1.Text = txt1.Text & Ra(0, oo - 1) & " (" & Hour(Ra(1, oo - 1)) & ":" & Minute(Ra(1, oo - 1)) & ":" & Second(Ra(1, oo - 1)) & "):" & Chr(13) & Chr(10) & _
"                  " & Ra(2, oo - 1) & Chr(13) & Chr(10)
DoEvents
Next
txt1.SelStart = Len(txt1.Text)
txt1.SelLength = 0
MaxGid = Ra(3, oo - 2)
If Me.Visible = True And La > 0 And frmGGL.WindowState <> 0 Then
    Me.Show
    Me.ZOrder 0
    txt2.SetFocus
End If
End If

End Sub

Private Sub txt2_Change()
If txt2.Text = Chr(13) & Chr(10) Then
    txt2.Text = ""
End If
End Sub

Private Sub txt2_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If Shift = 2 And KeyCode = 13 Then
    KeyCode = 0
    Call cmdSend_Click


End If
End Sub

'-- 合并文件
'-- 将之前分割出来的 1.mp3 和 2.mp3 合并为 music_new.mp3
Private Sub cmdAddFile_Click()
    Dim bytData() As Byte
    bytData = ReadFile(App.Path & "\1.mp3")
    Call WriteFile(App.Path & "\music_new.mp3", bytData)
    bytData = ReadFile(App.Path & "\2.mp3")
    Call WriteFile(App.Path & "\music_new.mp3", bytData, , False)
End Sub


