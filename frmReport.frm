VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmReport 
   Caption         =   "豪曼报表"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   LinkTopic       =   "Form3"
   ScaleHeight     =   7635
   ScaleWidth      =   8670
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdExcel 
      Caption         =   "导出Excel"
      Height          =   345
      Left            =   7320
      TabIndex        =   2
      Top             =   0
      Width           =   1365
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   375
      Left            =   7740
      Picture         =   "frmReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6990
      Visible         =   0   'False
      Width           =   675
   End
   Begin CRVIEWER9LibCtl.CRViewer9 cR1 
      Height          =   7635
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8625
      lastProp        =   500
      _cx             =   15214
      _cy             =   13467
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
cR1.PrintReport
End Sub

Private Sub Form_Resize()
cR1.Top = 0
cR1.Left = 0
cR1.Height = ScaleHeight
cR1.Width = ScaleWidth
cmdPrint.Left = ScaleWidth - cmdPrint.Width
cmdPrint.Top = ScaleHeight - cmdPrint.Height
cmdExcel.Left = ScaleWidth - cmdExcel.Width
 
End Sub
