VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4755
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdd 
      Left            =   1800
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.txt"
      Flags           =   1
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu ff 
      Caption         =   "File"
      Begin VB.Menu new 
         Caption         =   "New"
      End
      Begin VB.Menu open 
         Caption         =   "open..."
      End
      Begin VB.Menu save 
         Caption         =   "save"
      End
      Begin VB.Menu saveas 
         Caption         =   "saveas..."
      End
      Begin VB.Menu printt 
         Caption         =   "Print"
      End
   End
   Begin VB.Menu editttt 
      Caption         =   "Edit"
      Begin VB.Menu undooo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu cpyy 
         Caption         =   "copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu cutt 
         Caption         =   "cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu selectt 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu pastee 
         Caption         =   "paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu dell 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu findd 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu findnextt 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu dt 
      Caption         =   "Date/Time"
   End
   Begin VB.Menu faff 
      Caption         =   "Font"
      Begin VB.Menu fontt 
         Caption         =   "Font"
      End
      Begin VB.Menu colr 
         Caption         =   "Font's color"
      End
   End
   Begin VB.Menu gooto 
      Caption         =   "Goto"
   End
   Begin VB.Menu aboot 
      Caption         =   "About"
   End
   Begin VB.Menu eexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub aboot_Click()
MsgBox "This program is designed by Sherif Rofael in egypt on sep,2002 , mailto:ya3amo@hotmail.com  website:http://vbsherif.members.easyspace.com/", , "Notepad by Sherif"
End Sub

Private Sub colr_Click()
cdd.ShowColor
Text1.ForeColor = cdd.Color
End Sub


Private Sub cpyy_Click()
Clipboard.SetText Text1.SelText
'toclipboard = Text1.SelText
End Sub

Private Sub cutt_Click()
PREV = Text1.Text
Clipboard.SetText Text1.SelText
'toclipboard = Text1.SelText
Text1.SelText = ""
End Sub

Private Sub dell_Click()
PREV = Text1.Text
If Text1.SelLength = 0 Then Text1.SelLength = 1
Text1.SelText = ""
End Sub

Private Sub dt_Click()
PREV = Text1.Text
Text1.SelText = Time & "   " & Date
End Sub

Private Sub eexit_Click()
Dim ss
If Text1.Text = "" Then End
If Text1.Text <> "" Then
ss = MsgBox(" Do you want to save the changes? ", vbYesNoCancel + vbExclamation, "save changes?")
If ss = vbYes Then Call save_Click
If ss = vbNo Then End
'If ss = vbCancel Then Exit Sub
End If
End Sub

Private Sub findd_Click()
On Error GoTo exitt:
findwhat = InputBox("Find what?", "Sherif's Notepad")
starting = InStr(starting + 1, Text1.Text, findwhat)
Text1.SelStart = starting - 1
Text1.SelLength = Len(findwhat)
'findwhat = InputBox("Find what?", "Sherif's Notepad")
't = "'" & findwhat & "'"
'If starting = "" Then MsgBox "cannot find  " & t, , "Sherif's Notepad"
exitt:
End Sub

Private Sub findnextt_Click()
On Error GoTo exitt:
'findwhat = InputBox("Find what?", "Sherif's Notepad", findwhat)
starting = InStr(starting + 1, Text1.Text, findwhat)
Text1.SelStart = starting - 1
Text1.SelLength = Len(findwhat)
't = "'" & findwhat & "'"
'If starting = "" Then MsgBox "cannot find  " & t, , "Sherif's Notepad"
exitt:
End Sub

Private Sub fontt_Click()
On Error GoTo exitt:
cdd.FontName = "Comic Sans MS"
cdd.ShowFont
'cdd.Flags = cddCFBoth
Text1.FontName = cdd.FontName
Text1.FontSize = cdd.FontSize
Text1.FontBold = cdd.FontBold
Text1.FontItalic = cdd.FontItalic
Text1.FontUnderline = cdd.FontUnderline
Text1.FontStrikethru = cdd.FontStrikethru
exitt:
End Sub

Private Sub Form_Load()
findwhat = ""
starting = 0
'Text1.Width = Form1.Width
'Text1.Height = Form1.Height - 100
foreverlabel = " - NOTEPAD BY SHERIF"
Form1.Caption = "New Text File " & foreverlabel
End Sub

Private Sub Form_Resize()
Text1.Width = Form1.Width - 100
Text1.Height = Form1.Height - 800

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim ss
If Text1.Text <> "" Then
ss = MsgBox(" Do you want to save the changes? ", vbYesNoCancel + vbExclamation, "save changes?")
If ss = vbYes Then Call save_Click
If ss = vbNo Then End
If ss = vbCancel Then Cancel = 1
End If
'If Text1.Text <> "" Then Call save_Click
End Sub

Private Sub gooto_Click()
'sum = 0
On Error GoTo exitt:
linenumber = InputBox("Enter the line's Number you wants to go to ! " & vbCrLf & " Note: Blank Lines are Calculated.", "Goto !")
Dim lenn, getfromtext, eachline, sum, i

lenn = Len(Text1.Text)
'getfromtext = Input(lenn, Text1.Text)
eachline = Split(Text1.Text, vbCrLf)
For i = 0 To (linenumber - 2)
sum = sum + Len(eachline(i))
Next
Text1.SelStart = sum + 2 * i
Text1.SelLength = Len(eachline(i))
'Text1.Text = eachline(0)
exitt:
End Sub

Private Sub new_Click()
Text1.Text = ""
Form1.Caption = "New text file" & foreverlabel

End Sub

Private Sub open_Click()
Dim ss As String
Dim lenn, allwords

If Text1.Text <> "" Then
ss = MsgBox(" Do you want to save the changes? ", vbYesNoCancel + vbExclamation, "save changes?")
If ss = vbYes Then Call save_Click
If ss = vbNo Then GoTo openn:
If ss = vbCancel Then Exit Sub
End If

openn:
On Error GoTo exitt:
cdd.Filter = "Text Files (*.txt)|*.txt|All  Files (*.*)|*.*"
cdd.ShowOpen
filename = cdd.filename
lenn = FileLen(cdd.filename)
Open cdd.filename For Input As #1
allwords = Input(lenn, #1)
Close #1
Form1.Caption = cdd.FileTitle & foreverlabel
Text1.Text = allwords
exitt:
End Sub

Private Sub pastee_Click()
PREV = Text1.Text
'Text1.SelText = toclipboard
'Clipboard.GetText Text1.SelText
Text1.SelText = Clipboard.GetText
End Sub

Private Sub printt_Click()
'cdd.Flags = cddpddisableprinttofile
cdd.Copies = 2
cdd.PrinterDefault = True
cdd.ShowPrinter

End Sub

Private Sub save_Click()
On Error GoTo exitt:
If filename = "" Then Call saveas_Click
If filename <> "" Then
Open filename For Output As #1
Print #1, Text1.Text
Close #1
End If
exitt:
End Sub

Private Sub saveas_Click()
cdd.Filter = "Text Files (*.txt)|*.txt|All  Files (*.*)|*.*"
cdd.ShowSave
filename = cdd.filename
Open cdd.filename For Output As #1
Print #1, Text1.Text
Close #1
Form1.Caption = cdd.FileTitle & foreverlabel
End Sub


Private Sub selectt_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_Change()
'PREV = Text1.Text
End Sub

Private Sub undooo_Click()
prevtonext = Text1.Text
Text1.Text = PREV
PREV = prevtonext
End Sub






