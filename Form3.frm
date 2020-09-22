VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   420
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   1440
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   1440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Exit"
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   210
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   1600
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Check For Updates"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then End
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = QBColor(1)
Label2.BackColor = vbBlack
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then Unload Me
If Button = vbLeftButton Then Call update
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = vbBlack
Label2.BackColor = QBColor(1)
End Sub

