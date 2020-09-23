VERSION 5.00
Begin VB.UserControl txt 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OLLabel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OLLabel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "OLLabel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "OLLabel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "OLLabel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "OLLabel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "txt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
Label1.Top = UserControl.Height / 2 - 100
label2.Top = Label1.Top + 15
Label4.Top = Label1.Top
label2.Left = 0 + 15
Label1.Width = UserControl.Width
label2.Width = UserControl.Width
Label3.Width = UserControl.Width
Label4.Width = UserControl.Width
Label5.Width = UserControl.Width
Label6.Width = UserControl.Width
Label4.Left = 0 + 15
Label5.Top = Label1.Top - 15
Label6.Top = Label1.Top + 15
Label3.Top = Label1.Top
Label1.Height = Screen.Height
label2.Height = Screen.Height
Label3.Height = Screen.Height
Label4.Height = Screen.Height
Label5.Height = Screen.Height
Label6.Height = Screen.Height
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
    label2.Caption = Label1.Caption
    Label3.Caption = Label1.Caption
    Label4.Caption = Label1.Caption
    Label5.Caption = Label1.Caption
    Label6.Caption = Label1.Caption
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Label1.Caption = PropBag.ReadProperty("Caption", "Label1")
'    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
'    Label1.FontSize = PropBag.ReadProperty("FontSize", 0)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", Label1.Caption, "Label1")
'    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
'    Call PropBag.WriteProperty("FontSize", Label1.FontSize, 0)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Label1,Label1,-1,Font
'Public Property Get Font() As Font
'    Set Font = Label1.Font
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'    Set Label1.Font = New_Font
'    PropertyChanged "Font"
'    Set Label2.Font = New_Font
'    Set Label3.Font = New_Font
'    Set Label4.Font = New_Font
'    Set Label5.Font = New_Font
'    Set Label6.Font = New_Font
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Label1,Label1,-1,FontSize
'Public Property Get FontSize() As Single
'    FontSize = Label1.FontSize
'End Property
'
'Public Property Let FontSize(ByVal New_FontSize As Single)
'    Label1.FontSize() = New_FontSize
'    PropertyChanged "FontSize"
'    Label2.Font() = New_FontSize
'    Label3.Font() = New_FontSize
'    Label4.Font() = New_FontSize
'    Label5.Font() = New_FontSize
'    Label6.Font() = New_FontSize
'End Property
'

