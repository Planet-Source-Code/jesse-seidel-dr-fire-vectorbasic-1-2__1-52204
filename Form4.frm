VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save Code (.VZBC)"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   3120
      Width           =   5055
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   5775
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   3000
      Pattern         =   "*.VZBC"
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin Project1.chameleonButton chameleonButton2 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form4.frx":038A
      PICN            =   "Form4.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form4.frx":0740
      PICN            =   "Form4.frx":075C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub chameleonButton1_Click()
Command1_Click
End Sub

Private Sub chameleonButton2_Click()
Me.Hide
Form1.Enabled = True
Form1.SetFocus
End Sub

Private Sub Command1_Click()
On Error GoTo err:
WriteFile Dir1.Path & "\" & Text2.text & ".vzbc", Form1.rtb1.text
Me.Hide
Form1.Enabled = True
Form1.SetFocus
Exit Sub
err:
MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
Text2.text = File1.FileName
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
Text2.text = File1.FileName
End Sub

Private Sub File1_KeyUp(KeyCode As Integer, Shift As Integer)
Text2.text = File1.FileName
End Sub

Private Sub Form_Load()
Dir1.Path = App.Path
Form1.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim ab As String
Dim ac As String
ab = UCase(".VZBC")
ac = LCase(".vzbc")
Text1.text = Dir1.Path & "\" & Text2.text
Text2.text = Replace(Text2.text, ab, "")
Text2.text = Replace(Text2.text, ac, "")
End Sub


