VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About VectorBASIC"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks to Brian Bender for the color coding routine"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   0
      Top             =   3180
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "All other programming by Jesse Seidel"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks to Akhil P Jayaraj for the Win API class module"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks to gonchuki for the Chameleon Button"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   3180
      Left            =   0
      Picture         =   "Form8.frx":038A
      Top             =   0
      Width           =   4680
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Enabled = True
Me.Hide
End Sub
