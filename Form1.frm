VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "VectorBASIC"
   ClientHeight    =   5145
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6870
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSScriptControlCtl.ScriptControl sc1 
      Left            =   5280
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin Project1.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "VectorBASIC Status"
      Top             =   4800
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   609
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6870
      TabIndex        =   1
      Top             =   0
      Width           =   6870
      Begin Project1.chameleonButton chameleonButton2 
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         ToolTipText     =   "Main Form Controls"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
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
         MICON           =   "Form1.frx":038A
         PICN            =   "Form1.frx":03A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton color1 
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         ToolTipText     =   "Color code"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":0740
         PICN            =   "Form1.frx":075C
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
         Left            =   2400
         TabIndex        =   9
         ToolTipText     =   "Compile to EXE"
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
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
         MICON           =   "Form1.frx":0AF6
         PICN            =   "Form1.frx":0B12
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton stop1 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         ToolTipText     =   "End"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":11F4
         PICN            =   "Form1.frx":1210
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton run1 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         ToolTipText     =   "Start"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":1532
         PICN            =   "Form1.frx":154E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton save1 
         Height          =   375
         Left            =   720
         TabIndex        =   4
         ToolTipText     =   "Save As"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":1834
         PICN            =   "Form1.frx":1850
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton open1 
         Height          =   375
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "Open"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16711935
         MPTR            =   1
         MICON           =   "Form1.frx":1B72
         PICN            =   "Form1.frx":1B8E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton new1 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "New"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   -2147483643
         MPTR            =   1
         MICON           =   "Form1.frx":1EB0
         PICN            =   "Form1.frx":1ECC
         PICH            =   "Form1.frx":2176
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   3
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.OLLabel OLLabel1 
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "VectorBASIC"
      End
      Begin VB.Image Image4 
         Height          =   285
         Left            =   3840
         Picture         =   "Form1.frx":2420
         Top             =   30
         Width           =   30
      End
      Begin VB.Image Image3 
         Height          =   285
         Left            =   3240
         Picture         =   "Form1.frx":24FA
         Top             =   30
         Width           =   30
      End
      Begin VB.Image Image2 
         Height          =   285
         Left            =   2280
         Picture         =   "Form1.frx":25D4
         Top             =   30
         Width           =   30
      End
      Begin VB.Image Image1 
         Height          =   285
         Left            =   1200
         Picture         =   "Form1.frx":26AE
         Top             =   30
         Width           =   30
      End
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5318
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":2788
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Image15 
      Height          =   150
      Left            =   2520
      Picture         =   "Form1.frx":280A
      Top             =   3480
      Width           =   150
   End
   Begin VB.Image Image14 
      Height          =   195
      Left            =   2280
      Picture         =   "Form1.frx":298C
      Top             =   3480
      Width           =   165
   End
   Begin VB.Image Image13 
      Height          =   195
      Left            =   2040
      Picture         =   "Form1.frx":2BA2
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image Image12 
      Height          =   195
      Left            =   1800
      Picture         =   "Form1.frx":2DB8
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image Image11 
      Height          =   195
      Left            =   1560
      Picture         =   "Form1.frx":2FCE
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image Image10 
      Height          =   195
      Left            =   1320
      Picture         =   "Form1.frx":31E4
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image Image9 
      Height          =   195
      Left            =   1080
      Picture         =   "Form1.frx":33FA
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image Image8 
      Height          =   165
      Left            =   840
      Picture         =   "Form1.frx":3610
      Top             =   3480
      Width           =   165
   End
   Begin VB.Image Image7 
      Height          =   195
      Left            =   600
      Picture         =   "Form1.frx":37DE
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image Image6 
      Height          =   195
      Left            =   360
      Picture         =   "Form1.frx":39F4
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   195
      Left            =   120
      Picture         =   "Form1.frx":3C0A
      Top             =   3480
      Width           =   195
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "New"
      End
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu saveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu compile 
      Caption         =   "Compile"
      Begin VB.Menu start 
         Caption         =   "Start"
      End
      Begin VB.Menu end 
         Caption         =   "End"
      End
      Begin VB.Menu ln2 
         Caption         =   "-"
      End
      Begin VB.Menu compexe 
         Caption         =   "Compile to EXE"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About VectorBASIC"
      End
      Begin VB.Menu ln3 
         Caption         =   "-"
      End
      Begin VB.Menu howtouse 
         Caption         =   "How to use VectorBASIC"
      End
      Begin VB.Menu onlinehelp 
         Caption         =   "Online Help"
      End
      Begin VB.Menu contact 
         Caption         =   "Contact VectorBASIC creators"
      End
      Begin VB.Menu ln4 
         Caption         =   "-"
      End
      Begin VB.Menu vbhome 
         Caption         =   "Visit VectorBASIC homepage"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim WAPI As CWinAPI
Private Declare Function fCreateShellLink Lib "VB6STKIT.DLL" (ByVal _
        lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
        lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, _
                        ByVal uFlags As Long, ByVal dwItem1 As Long, _
                        ByVal dwItem2 As Long)

Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

Private Sub about_Click()
Form7.Show
Me.Enabled = False
End Sub

Private Sub chameleonButton2_Click()
Form6.Show
End Sub

Private Sub color1_Click()
Form1.Text3.text = Form1.rtb1.text
Form1.rtb1.text = ""
Colorize Form1.rtb1, Form1.Text3.text
Form1.rtb1.SelStart = Len(Form1.rtb1.text)
End Sub

Private Sub end_Click()
Unload Form5
End Sub

Private Sub exit_Click()
Dim frm As Form
StatusBar1.caption = "Are you sure you wish to exit?"
If MsgBox("Are you sure you wish to exit?", vbYesNo, "Exit") = vbYes Then
For Each frm In Forms
Unload frm
Next
Else
Cancel = 1
StatusBar1.caption = "VectorBASIC"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Set WAPI = New CWinAPI
open1.Left = new1.width
save1.Left = open1.Left + open1.width
about.caption = "About VectorBASIC " & App.Major & "." & App.Minor
OLLabel1.caption = "VectorBASIC " & App.Major & "." & App.Minor
Me.caption = "VectorBASIC " & App.Major & "." & App.Minor
StatusBar1.caption = "Welcome to VectorBASIC!"
SetAsoc
sc1.AddObject "TForm", Me
sc1.AddObject "MForm", Form5
sc1.AddObject "Screen", Screen
sc1.AddObject "App", App

WAPI.MenuBitmaps Form1.hWnd, 0, 0, Image5.Picture
WAPI.MenuBitmaps Form1.hWnd, 0, 1, Image6.Picture
WAPI.MenuBitmaps Form1.hWnd, 0, 2, Image7.Picture
WAPI.MenuBitmaps Form1.hWnd, 0, 4, Image8.Picture
WAPI.MenuBitmaps Form1.hWnd, 1, 0, Image9.Picture
WAPI.MenuBitmaps Form1.hWnd, 1, 1, Image10.Picture
WAPI.MenuBitmaps Form1.hWnd, 1, 3, Image11.Picture
WAPI.MenuBitmaps Form1.hWnd, 2, 6, Image12.Picture
WAPI.MenuBitmaps Form1.hWnd, 2, 3, Image13.Picture
WAPI.MenuBitmaps Form1.hWnd, 2, 0, Image14.Picture
WAPI.MenuBitmaps Form1.hWnd, 2, 4, Image15.Picture
WAPI.MenuBitmaps Form1.hWnd, 2, 2, Image13.Picture
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim frm As Form
StatusBar1.caption = "Are you sure you wish to exit?"
If MsgBox("Are you sure you wish to exit?", vbYesNo, "Exit") = vbYes Then
For Each frm In Forms
Unload frm
Next
Else
Cancel = 1
StatusBar1.caption = "VectorBASIC"
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
rtb1.height = Me.height - 670 - Picture1.height - StatusBar1.height
rtb1.width = Me.width - 120
OLLabel1.Left = Me.width - OLLabel1.width
End Sub

Private Sub new_Click()
If MsgBox("Are you sure you wish to start a blank code?", vbYesNo, "New Code") = vbYes Then
rtb1.text = ""
End If
End Sub

Private Sub new1_Click()
If MsgBox("Are you sure you wish to start a blank code?", vbYesNo, "New Code") = vbYes Then
rtb1.text = ""
End If
End Sub

Private Sub open_Click()
Form3.Show
Form1.Enabled = False
End Sub

Private Sub open1_Click()
Form3.Show
Form1.Enabled = False
End Sub


Private Sub run1_Click()
On Error Resume Next
Unload Form5
sc1.AddCode rtb1.text
sc1.Run "Main"
NewElement = 1
End Sub

Private Sub save1_Click()
Form4.Show
End Sub

Private Sub saveas_Click()
Form4.Show
End Sub

Public Sub SetAsoc()
On Error Resume Next
    Dim strString As String
    Dim lngDword As Long
    Dim Record As String


    If Command$ <> "%1" And Command$ <> "" Then
        'Command$ is the file you need To open!

        'Load the file
rtb1.LoadFile (Command$)

        'Add your file to the Recent file folder:
        lReturn = fCreateShellLink("..\..\Recent", _
                Command$, Command$, "")

    End If


    'See if our file extension already exists:
    If GetString(HKEY_CLASSES_ROOT, ".vzbc", "Content Type") = "" Then
        'Nope - not added yet. Register the file type:
        
        'create an entry in the class key
        Call SaveString(HKEY_CLASSES_ROOT, ".vzbc", "", "vzbcfile")
        'content type
        Call SaveString(HKEY_CLASSES_ROOT, ".vzbc", "Content Type", "text/plain")
        'name
        Call SaveString(HKEY_CLASSES_ROOT, "vzbcfile", "", "VectorBASIC Code File")
        'edit flags
        Call SaveDWord(HKEY_CLASSES_ROOT, "vzbcfile", "EditFlags", "0000")
        'file's icon (can be an icon file, or an icon located within a dll file)
        'in this example, I am using a resource icon in this exe, 0 (app icon).
        Call SaveString(HKEY_CLASSES_ROOT, "vzbcfile\DefaultIcon", "", App.path & "\" & App.EXEName & ".exe,0")
        'Shell
        Call SaveString(HKEY_CLASSES_ROOT, "vzbcfile\Shell", "", "")
        'Shell Open
        Call SaveString(HKEY_CLASSES_ROOT, "vzbcfile\Shell\Open", "", "")
        'Shell open command
        Call SaveString(HKEY_CLASSES_ROOT, "vzbcfile\Shell\Open\Command", "", App.path & "\" & App.EXEName & ".exe %1")
        'Update the Windows Icon Cache to see our icon right away:
        SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0

    End If
Form1.Text3.text = Form1.rtb1.text
Form1.rtb1.text = ""
Colorize Form1.rtb1, Form1.Text3.text
Form1.rtb1.SelStart = Len(Form1.rtb1.text)
End Sub

Private Sub start_Click()
On Error Resume Next
sc1.AddCode rtb1.text
sc1.Run "Main"
sc1.AddObject "TForm", Me
sc1.AddObject "MForm", Form5
sc1.AddObject "Screen", Screen
sc1.AddObject "App", App
NewElement = 1
End Sub

Private Sub stop1_Click()
Unload Form5
End Sub

