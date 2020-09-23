VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "HTML Help Maker"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11190
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   746
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTOC 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   2400
      TabIndex        =   5
      ToolTipText     =   "Add item"
      Top             =   6555
      Width           =   1095
   End
   Begin VB.CommandButton cmdTOC 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   5
      Left            =   1920
      TabIndex        =   6
      ToolTipText     =   "Delete Selection"
      Top             =   6555
      Width           =   495
   End
   Begin VB.CommandButton cmdTOC 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "Move Selection Out"
      Top             =   6555
      Width           =   495
   End
   Begin VB.CommandButton cmdTOC 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "Move Selection In"
      Top             =   6555
      Width           =   495
   End
   Begin VB.CommandButton cmdTOC 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Move Selection Down"
      Top             =   6555
      Width           =   495
   End
   Begin VB.CommandButton cmdTOC 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Move Selection Up"
      Top             =   6555
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   720
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Template (*.hhm)|*.hhm|HTML File (*.html)|*.html|"
      Flags           =   7
   End
   Begin VB.ListBox lstToc 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   3600
      ScaleHeight     =   465
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   7
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton cmdHTML 
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdHTML 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1440
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdHTML 
         Caption         =   ">"
         Height          =   375
         Index           =   4
         Left            =   1080
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdHTML 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   11
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdHTML 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Text            =   "frmMain.frx":10C2
         Top             =   360
         Width           =   7575
      End
      Begin VB.CommandButton cmdHTML 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSavePrj 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSavePrjAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuLine356 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save HTML"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save HTML As..."
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Options"
      Begin VB.Menu mnuOptTemp 
         Caption         =   "Edit HTML Template"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuLine32211 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptAuto 
         Caption         =   "Auto Line Breaks"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHTMLHelp 
      Caption         =   "Html Help"
      Begin VB.Menu mnuHtml 
         Caption         =   "Bold"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuHtml 
         Caption         =   "Italics"
         Index           =   1
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuHtml 
         Caption         =   "Underlined"
         Index           =   2
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuHtml 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHtml 
         Caption         =   "Ordered List"
         Index           =   4
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuHtml 
         Caption         =   "Unordered List"
         Index           =   5
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuHtml 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuHtml 
         Caption         =   "Bullet"
         Index           =   7
         Shortcut        =   ^K
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mChanged As Boolean

Dim mArr() As String, mFileName As String, mPrjFileName As String

Private Function ExtractInfo(ByVal Text As String, ByVal Tag As String) As String
Dim a As Long, b As Long
a = InStr(Text, "<" & Tag & ">")
 If a = 0 Then Exit Function
b = InStr(a + 1, Text, "</" & Tag & ">")
 If b = 0 Then Exit Function

ExtractInfo = Mid(Text, a + (Len(Tag) + 2), b - a - (Len(Tag) + 2))
End Function

Private Sub SaveHHM(ByVal Filename As String)
Dim i As Integer, l As Long, Data As String, Final As String, Help As String
Data = "<htmlfilename>" & mFileName & "</htmlfilename>"
Data = Data & "<prjfilename>" & mPrjFileName & "</prjfilename>"
Data = Data & "<template>" & frmTemp.txtData.Text & "</template>"
Data = Data & "<autoline>" & mnuOptAuto.Checked & "</autoline>"

For i = 0 To lstToc.ListCount - 1
 Final = Final & lstToc.List(i) & vbCrLf
 Help = Help & mArr(i) & "<*code*>"
Next i
Help = Left(Help, Len(Help) - 8)
If FileExist(Filename) = True Then Call Kill(Filename)
l = FreeFile()
Open Filename For Binary Access Write As #l
 Put #l, , CStr(Data & "<*split*>" & Final & "<*split*>" & Help)
Close #l
End Sub

Private Function CountIndent(ByVal Index As Integer) As Integer
Dim s As String, k As Integer, t As Integer
 s = lstToc.List(Index)
 k = 0
 For t = 1 To Len(s)
  If Mid(s, t, 1) = "." Then k = k + 1 Else Exit For
 Next t
 CountIndent = Int(k / 3)
End Function

Private Sub SaveHTML(ByVal Filename As String)
Dim i As Integer, s As String
Dim t As Integer, k As Integer
Dim Final As String, F As Integer
Dim j As Integer, y As Integer
Dim Help As String, l As Long, max As Integer

F = 0
Final = "<ol>" & vbCrLf
For i = 0 To lstToc.ListCount - 1
 s = lstToc.List(i)
 k = 0
 For t = 1 To Len(s)
  If Mid(s, t, 1) = "." Then k = k + 1 Else Exit For
 Next t
 If k <> 0 Then s = Mid(s, k + 1)
 k = Int(k / 3)
 
 If k <> 0 Then
 
 If k > max Then
  y = y + 1
  Final = Final & "<ul>" & vbCrLf
  Help = Help & "<ul>"
  Final = Final & "<li><a href=""#ch" & j & "sub" & y & """>" & s & "</a>" & vbCrLf
  max = k
 ElseIf k < max Then
  y = y + 1
  Final = Final & "</ul>" & vbCrLf
  Help = Help & "</ul>"
  Final = Final & "<li><a href=""#ch" & j & "sub" & y & """>" & s & "</a>" & vbCrLf
  max = max - 1
 Else
  y = y + 1
  Final = Final & "<li><a href=""#ch" & j & "sub" & y & """>" & s & "</a>" & vbCrLf
  
 End If
 Else
   j = j + 1
  If max <> 0 Then
   For t = 1 To max
    Help = Help & "</ul>"
    Final = Final & "</ul>" & vbCrLf
   Next t
  End If
  y = 0
  Final = Final & "<li><a href=""#ch" & j & "sub" & 0 & """>" & s & "</a>" & vbCrLf

  max = 0
 End If
 
 '  Final = Final & String(k, Chr(9)) & "<li><a href=""#ch" & j & "sub" & y & """>" & s & "</a>" & vbCrLf
   
   Help = Help & "<a name=""ch" & j & "sub" & y & """>" & "</a>" & j & " " & IIf(y <> 0, "(" & y & ")", "") & " <b>" & s & "</b>" & vbCrLf & "<br/>" & IIf(mnuOptAuto.Checked = True, Replace(mArr(i), vbCrLf, "<br/>" & vbCrLf), mArr(i)) & "<br/>" & vbCrLf & "<a href=""#"">top ^</a><br/><br/>" & vbCrLf
Next i

  If max <> 0 Then
   For t = 1 To max
    Final = Final & "</ul>" & vbCrLf
   Next t
  End If

Final = Final & "</ol>" & vbCrLf
If FileExist(cd.Filename) = True Then Call Kill(cd.Filename)
l = FreeFile()
Open cd.Filename For Binary Access Write As #l
 Put #l, , CStr(Replace(frmTemp.txtData.Text, "<%body%>", Final & vbCrLf & "<br/><hr><br/>" & Help, , , vbTextCompare))
Close #l
End Sub

Private Function FileExist(ByVal F As String) As Boolean
On Error GoTo 1
Call FileLen(F)
FileExist = True
1
End Function

Private Sub cmdHTML_Click(Index As Integer)
Call mnuHTML_Click(Index)
End Sub

Private Sub cmdTOC_Click(Index As Integer)
mChanged = True
Dim s As String, i As Integer
Dim j As Integer, h As String

  i = lstToc.ListIndex
  s = lstToc.Text
Select Case Index
 Case 0 'move up
  If i <= 0 Then Exit Sub
  lstToc.RemoveItem i
  lstToc.AddItem s, i - 1
  
  i = i - 1
  h = mArr(i)
  mArr(i) = mArr(i + 1)
  mArr(i + 1) = h
 Case 1 'move down
  If i = lstToc.ListCount - 1 Or i = -1 Then Exit Sub
  lstToc.RemoveItem i
  lstToc.AddItem s, i + 1
  i = i + 1

  h = mArr(i)
  mArr(i) = mArr(i - 1)
  mArr(i - 1) = h

 Case 2 'move left
  If i = -1 Then Exit Sub
  If Left(s, 3) <> "..." Then Exit Sub
  lstToc.RemoveItem i
  lstToc.AddItem Mid(s, 4), i
 Case 3 'move right
  'If Left(s, 3) = "..." Then Exit Sub
  If i = -1 Then Exit Sub
  'If CountIndent(i) >= CountIndent(i - 1) + 1 Then Exit Sub
  lstToc.RemoveItem i
  lstToc.AddItem "..." & s, i
 Case 4 'add
  s = InputBox("Enter Chapter or Subsection Name", "Add")
  If s = "" Then Exit Sub
   lstToc.AddItem s
   i = lstToc.ListCount - 1
 Case 5 'delete
  If i = -1 Then Exit Sub
  lstToc.RemoveItem i
  For j = i To UBound(mArr()) - 1
   mArr(j) = mArr(j + 1)
  Next j
  i = i - 1
End Select

   If i > lstToc.ListCount - 1 Then i = lstToc.ListCount - 1
   If i >= 0 Then lstToc.ListIndex = i Else If lstToc.ListCount <> 0 Then lstToc.ListIndex = 0 Else Exit Sub
   ReDim Preserve mArr(lstToc.ListCount - 1)
End Sub

Private Sub Form_Load()
Call Load(frmTemp)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If mChanged = True Then
 If MsgBox("The File has changed, do you wish to save first?", vbQuestion + vbYesNo, "File Changed") = vbYes Then
  Call mnuFileSavePrj_Click
 End If
End If
End
End Sub

Private Sub Form_Resize()
On Error GoTo 1
lstToc.Height = ScaleHeight - cmdTOC(0).Height
Dim i As Integer
For i = 0 To cmdTOC.Count - 1
 cmdTOC(i).Top = lstToc.Top + lstToc.Height + 2
Next i

picBack.Height = ScaleHeight
picBack.Width = ScaleWidth - lstToc.Width - 8

txtEdit.Width = picBack.Width
txtEdit.Height = picBack.Height - cmdHTML(0).Height
1
End Sub

Private Sub lstToc_Click()
On Error Resume Next
txtEdit.Text = mArr(lstToc.ListIndex)
txtEdit.Locked = False
txtEdit.SetFocus
End Sub

Private Sub lstToc_DblClick()
Dim s As String, i As Integer
i = lstToc.ListIndex
If i = -1 Then Exit Sub

s = lstToc.Text
s = InputBox("Edit Title", "Edit", s)
If s = "" Then Exit Sub

lstToc.RemoveItem i
lstToc.AddItem s, i
End Sub

Private Sub mnuFile_Click()
mnuFileSave.Enabled = (mFileName <> "")
mnuFileSavePrj.Enabled = (mPrjFileName <> "")
End Sub

Private Sub mnuFileExit_Click()
Call Unload(Me)
End Sub

Private Sub mnuFileNew_Click()
If mChanged = True Then
 If MsgBox("The File has changed, do you wish to save first?", vbQuestion + vbYesNo, "File Changed") = vbYes Then
  Call mnuFileSavePrj_Click
 End If
End If

mChanged = False
txtEdit.Text = ""
mFileName = ""
lstToc.Clear
ReDim mArr(0)
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo 1

If mChanged = True Then
 If MsgBox("The File has changed, do you wish to save first?", vbQuestion + vbYesNo, "File Changed") = vbYes Then
  Call mnuFileSavePrj_Click
 End If
End If

mChanged = False
cd.Filter = "Template (*.hhm)|*.hhm"
cd.Filename = mPrjFileName
cd.ShowOpen

Dim s As String, l As Long

l = FreeFile()
Open cd.Filename For Input As #l
 s = Input(LOF(l), #l)
Close #l

lstToc.Clear

Dim arr() As String, v As Variant
Dim d As String, e As String, i As Integer
Dim arrx() As String
arrx() = Split(s, "<*split*>")

mPrjFileName = ExtractInfo(arrx(0), "prjfilename")
mFileName = ExtractInfo(arrx(0), "htmlfilename")
frmTemp.txtData.Text = ExtractInfo(arrx(0), "template")

If ExtractInfo(arrx(0), "autoline") = "" Then
mnuOptAuto.Checked = True
Else
mnuOptAuto.Checked = CBool(ExtractInfo(arrx(0), "autoline"))
End If

e = arrx(2) 'Mid(s, InStr(s, "<*split*>") + 9)
arr = Split(e, "<*code*>")

For i = 0 To UBound(arr())
 ReDim Preserve mArr(i)
 mArr(i) = arr(i)
Next i

d = arrx(1) ' Left(s, InStr(s, "<*split*>") - 3)
arr = Split(d, vbCrLf)

For Each v In arr()
 If v <> "" Then lstToc.AddItem v
Next v
mPrjFileName = cd.Filename
1
End Sub

Private Sub mnuFileSave_Click()
Call SaveHTML(mFileName)
End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo 1
cd.Filter = "HTML File (*.html)|*.html"
cd.Filename = mFileName
cd.ShowSave

mFileName = cd.Filename
Call SaveHTML(cd.Filename)
1
End Sub

Private Sub mnuFileSavePrj_Click()
If mPrjFileName = "" Then Call mnuFileSavePrjAs_Click: Exit Sub
mChanged = False
Call SaveHHM(mPrjFileName)
End Sub

Private Sub mnuFileSavePrjAs_Click()
On Error GoTo 1
cd.Filter = "Template (*.hhm)|*.hhm"
cd.Filename = mPrjFileName
cd.ShowSave
mChanged = False
mPrjFileName = cd.Filename
Call SaveHHM(cd.Filename)
1

End Sub

Private Sub mnuHTML_Click(Index As Integer)
If txtEdit.SelLength <> 0 Then
    Dim i As Long
    i = txtEdit.SelStart + Len("<b>" & txtEdit.SelText & "</b>")
    Select Case Index
     Case 0: txtEdit.SelText = "<b>" & txtEdit.SelText & "</b>"
     Case 1: txtEdit.SelText = "<i>" & txtEdit.SelText & "</i>"
     Case 2: txtEdit.SelText = "<u>" & txtEdit.SelText & "</u>"
     Case 4: txtEdit.SelText = "<ol>" & vbCrLf & txtEdit.SelText & vbCrLf & "</ol>"
     Case 5: txtEdit.SelText = "<ul>" & vbCrLf & txtEdit.SelText & vbCrLf & "</ul>"
     Case 7: txtEdit.SelText = "<li>" & txtEdit.SelText
    End Select
    txtEdit.SelStart = i
End If
txtEdit.SetFocus
End Sub

Private Sub mnuOptAuto_Click()
mnuOptAuto.Checked = IIf(mnuOptAuto.Checked = False, True, False)
End Sub

Private Sub mnuOptTemp_Click()
frmTemp.Show vbModal
End Sub

Private Sub txtEdit_Change()
If lstToc.ListIndex <> -1 Then mArr(lstToc.ListIndex) = txtEdit.Text
End Sub
