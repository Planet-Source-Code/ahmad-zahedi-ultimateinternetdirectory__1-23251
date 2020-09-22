VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   LinkTopic       =   "Form2"
   ScaleHeight     =   2865
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   2800
      TabIndex        =   3
      Top             =   360
      Width           =   135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFC0&
      Height          =   1815
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   840
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      Height          =   2895
      Left            =   2550
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      Height          =   2895
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   165
      Left            =   2400
      TabIndex        =   1
      Top             =   30
      Width           =   120
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   1
      Left            =   -480
      Picture         =   "Form2.frx":010D
      Top             =   2640
      Width           =   4500
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000B&
      BorderStyle     =   3  'Dot
      Height          =   135
      Left            =   0
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   555
      Index           =   0
      Left            =   600
      Picture         =   "Form2.frx":35EF
      Top             =   240
      Width           =   1305
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   165
      Left            =   2430
      TabIndex        =   2
      Top             =   60
      Width           =   120
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   0
      Picture         =   "Form2.frx":5C3D
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormDrag(Me)
End Sub


Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
col = Label7.ForeColor
Label5.Visible = False
Label7.ForeColor = Label5.ForeColor
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Visible = True
Label7.ForeColor = col
Unload Me
End Sub

Private Sub Text1_GotFocus()
Command1.SetFocus
End Sub
