VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VectorBASIC"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   4425
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "This option lets you start off with a completely blank code for you to edit in the code editor."
         Height          =   1095
         Left            =   2400
         TabIndex        =   10
         Top             =   120
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   2280
         X2              =   2280
         Y1              =   1200
         Y2              =   -120
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   1080
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Existing Code"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   315
         Left            =   1440
         Picture         =   "Form2.frx":038A
         Top             =   120
         Width           =   330
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "  New Code"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   360
         Picture         =   "Form2.frx":0960
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   80
         TabIndex        =   7
         Top             =   480
         Width           =   800
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1000
      End
   End
   Begin Project1.chameleonButton chameleonButton2 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "Form2.frx":0D92
      PICN            =   "Form2.frx":0DAE
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
      Left            =   1080
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "Form2.frx":1148
      PICN            =   "Form2.frx":1164
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Select an option below:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
Me.Hide
Form1.Enabled = True
Form1.SetFocus
End Sub

Private Sub chameleonButton2_Click()
If Label6.Visible = True Then
Me.Hide
Form1.Enabled = True
Form1.SetFocus
Else
Form3.Show
Form1.Enabled = False
Me.Hide
End If
End Sub

Private Sub Form_Load()
Form1.Enabled = False
Form1.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Enabled = True
End Sub

Private Sub Label4_Click()
Label7.Visible = False
Label6.Visible = True
Label8.caption = "This option lets you start off with a completely blank code for you to edit in the code editor."
End Sub

Private Sub Label5_Click()
Label7.Visible = True
Label6.Visible = False
Label8.caption = "This option lets you open a VectorBASIC code file that you have already made, and edit it in the code editor."
End Sub

Private Sub Picture1_Click()
Label7.Visible = False
Label6.Visible = False
Label8.caption = "Select an option"
End Sub
