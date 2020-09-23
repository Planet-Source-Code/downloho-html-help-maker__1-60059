VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTemp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HTML Template"
   ClientHeight    =   5895
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cd 
      Left            =   2280
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "HTML Files (*.html)|*.html"
      Flags           =   7
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmTemp.frx":0000
      Top             =   0
      Width           =   9015
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "Open File"
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = -1
Me.Hide
End Sub

Private Sub mnuOpen_Click()
On Error GoTo 1
cd.ShowOpen

Open cd.Filename For Input As #1
 txtData.Text = Input(LOF(1), #1)
Close #1
1
End Sub

Private Sub txtData_Change()
frmMain.mChanged = True
End Sub
