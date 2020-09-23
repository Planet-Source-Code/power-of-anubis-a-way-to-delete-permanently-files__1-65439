VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1692
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8364
   LinkTopic       =   "Form1"
   ScaleHeight     =   1692
   ScaleWidth      =   8364
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Delete file !"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   8052
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   600
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BROWSE"
      Height          =   372
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   288
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   156
      Width           =   6372
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3 Steps
'=======
'1.) Opening the file and deleting all the DATA
'2.) Changing the extension
'3.) Deleting

Private Sub Command1_Click()
   On Error GoTo ErrHandler
   CommonDialog1.Filter = "All Files"
   CommonDialog1.FilterIndex = 2
   CommonDialog1.DialogTitle = "Open file to delete..."
   CommonDialog1.ShowOpen

Text1.Text = CommonDialog1.FileName

ErrHandler:
   Exit Sub
End Sub

Private Sub Command2_Click()
'Opens the file and erases all the data
 Open CommonDialog1.FileName For Output As #1
  Print #1, ""
 Close #1
 'Changed the extension
 changeEX = ChangeFileExt(CommonDialog1.FileName, "tmp")
 
MsgBox "The file was permanently deleted !"
End Sub

Public Function ChangeFileExt(ByVal theflenme As String, ByVal newext As String) As Boolean
'Changing the original extension
Dim x As Long
Dim xy As Long
Dim newflenme As String

On Error Resume Next
ChangeFileExt = False
If theflenme = "" Then Exit Function
x = 0
Do
xy = x
x = InStr(x + 1, theflenme, ".", vbBinaryCompare)

Loop Until x = 0

If xy > 0 Then
  newflenme = Left(theflenme, xy - 1)
Else
  newflenme = theflenme
End If

newflenme = newflenme & "." & newext

Err.Clear

Name theflenme As newflenme
Kill newflenme
If Err.Number = 0 Then ChangeFileExt = True

End Function


Private Sub Form_Load()

End Sub
