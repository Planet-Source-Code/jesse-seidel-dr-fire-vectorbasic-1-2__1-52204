VERSION 5.00
Begin VB.UserControl StatusBar 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "StatusBar.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin Project1.txt label2 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   30
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   0
      Picture         =   "StatusBar.ctx":0156
      Top             =   50
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   1320
      Picture         =   "StatusBar.ctx":04E0
      Top             =   0
      Width           =   30
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   0
      Picture         =   "StatusBar.ctx":05D2
      Top             =   0
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   0
      Picture         =   "StatusBar.ctx":0728
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
Image3.Left = UserControl.Width - Image3.Width
Label1.Left = Image3.Left - Label1.Width + 50
Image1.Width = UserControl.Width
UserControl.Height = Image1.Height
label2.Width = UserControl.Width - label2.Left - Label1.Width - 20
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = label2.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    label2.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)


    label2.Caption = PropBag.ReadProperty("Caption", "StatusBar")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", label2.Caption, "StatusBar")
End Sub

