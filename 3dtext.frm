VERSION 5.00
Begin VB.Form text3d 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D text example"
   ClientHeight    =   2880
   ClientLeft      =   1125
   ClientTop       =   1500
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2880
   ScaleWidth      =   6210
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2160
      TabIndex        =   11
      Text            =   "3d Text Example"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "3d Direction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
      Begin VB.CommandButton Command8 
         Caption         =   "6"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "5"
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "4"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "3"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "2"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "1"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click the direction buttons to see different 3d text"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   2280
      Width           =   3615
   End
End
Attribute VB_Name = "text3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Draw_3d(direction)
Dim X
Picture1.BackColor = &H8000000F ' clear the picture box
 Picture1.Refresh
 Picture1.Font.Size = 20

 For X = 0 To 200
 
 If direction = "tl" Then
 Draw3dTxt Picture1, 1500 - X, 240 - X, Text1, X, X, X
 ElseIf direction = "tm" Then
 Draw3dTxt Picture1, 1500, 240 - X, Text1, X, X, X
 ElseIf direction = "tr" Then
 Draw3dTxt Picture1, 1500 + X, 240 - X, Text1, X, X, X
 
 ElseIf direction = "bl" Then
 Draw3dTxt Picture1, 1500 - X, X, Text1, X, X, X
 ElseIf direction = "bm" Then
 Draw3dTxt Picture1, 1500, X, Text1, X, X, X
 ElseIf direction = "br" Then
 Draw3dTxt Picture1, 1500 + X, X, Text1, X, X, X
 End If
 Next X
End Sub
Private Sub Draw3dTxt(ByVal canvas As Object, ByVal start_x As Single, ByVal start_y As Single, ByVal txt As String, r, g, b)
    canvas.CurrentX = start_x
    canvas.CurrentY = start_y
   
        canvas.ForeColor = RGB(r + 25, g + 25, b + 25)
        canvas.Print txt

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Picture1.BackColor = &H8000000F ' clear the picture box
Text1.Text = "http://members.xoom.com/devsfort"
 Picture1.Refresh
 Picture1.FontSize = 10

 Picture1.CurrentX = 100
 Picture1.CurrentY = 100
 Picture1.ForeColor = RGB(0, 0, 0)
 Picture1.Print "This example was created by dev." & Chr(13) & Chr(10)
 Picture1.CurrentX = 100
 Picture1.CurrentY = 500
 Picture1.ForeColor = RGB(255, 0, 0)
 Picture1.Print "Check out "

 Picture1.Refresh
 Picture1.CurrentX = 1000
 Picture1.CurrentY = 500
 Picture1.ForeColor = RGB(0, 0, 255)
 Picture1.FontUnderline = True
 Picture1.Print "http://members.xoom.com/devsfort"

 Picture1.CurrentY = 500
 Picture1.CurrentX = 3900
 Picture1.FontUnderline = False
 Picture1.ForeColor = RGB(0, 0, 0)
 Picture1.Print "for more great examples!"
End Sub

Private Sub Command3_Click()
Draw_3d "tl"
End Sub

Private Sub Command4_Click()
Draw_3d "tm"
End Sub

Private Sub Command5_Click()
Draw_3d "tr"
End Sub

Private Sub Command6_Click()
Draw_3d "bl"
End Sub

Private Sub Command7_Click()
Draw_3d "bm"
End Sub

Private Sub Command8_Click()
Draw_3d "br"
End Sub

