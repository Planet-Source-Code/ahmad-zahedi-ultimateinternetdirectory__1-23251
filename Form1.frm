VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4725
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   4725
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   855
      Left            =   1680
      TabIndex        =   45
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      _Version        =   327680
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":030A
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   44
      Text            =   "Form1.frx":03D3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2160
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.ListBox List3 
      Height          =   840
      ItemData        =   "Form1.frx":03D9
      Left            =   3240
      List            =   "Form1.frx":03DB
      TabIndex        =   36
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List6 
      Height          =   450
      ItemData        =   "Form1.frx":03DD
      Left            =   3840
      List            =   "Form1.frx":03DF
      TabIndex        =   42
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   120
      TabIndex        =   41
      Text            =   "Type In The First Few Letters Of The Search Engine Your Looking For"
      Top             =   1020
      Width           =   4455
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3120
      Top             =   2160
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   3360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327680
      BorderStyle     =   1
      Appearance      =   1
      Max             =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   315
      Left            =   960
      TabIndex        =   34
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   32
      Text            =   "Search String"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   33
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   255
      ItemData        =   "Form1.frx":03E1
      Left            =   3000
      List            =   "Form1.frx":03E3
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   2010
      ItemData        =   "Form1.frx":03E5
      Left            =   120
      List            =   "Form1.frx":03E7
      TabIndex        =   0
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Image Image5 
      Height          =   4365
      Left            =   0
      Picture         =   "Form1.frx":03E9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Image6 
      Height          =   4365
      Left            =   4575
      Picture         =   "Form1.frx":266F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   150
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "warez86@earthlink.net"
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   2880
      TabIndex        =   43
      Top             =   4155
      Width           =   1635
   End
   Begin VB.Image Image7 
      Height          =   150
      Left            =   0
      Picture         =   "Form1.frx":48F5
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   4725
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   120
      Picture         =   "Form1.frx":6973
      Top             =   780
      Width           =   750
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7595
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":81C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":8DF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":9A2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":A65D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":B28F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":BEC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":CAF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":D725
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":E357
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   4200
      TabIndex        =   37
      Top             =   30
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   4215
      TabIndex        =   39
      Top             =   60
      Width           =   75
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
      Left            =   4320
      TabIndex        =   35
      Top             =   30
      Width           =   120
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Ultimate Internet Directory-[Search Engines]"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   165
      Left            =   840
      TabIndex        =   31
      Top             =   30
      Width           =   2970
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Z"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   26
      Left            =   4080
      TabIndex        =   30
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Y"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   25
      Left            =   3945
      TabIndex        =   29
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "X"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   24
      Left            =   3795
      TabIndex        =   28
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "W"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   23
      Left            =   3600
      TabIndex        =   27
      Top             =   840
      Width           =   165
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "V"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   22
      Left            =   3480
      TabIndex        =   26
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "U"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   21
      Left            =   3345
      TabIndex        =   25
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "T"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   20
      Left            =   3225
      TabIndex        =   24
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "S"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   19
      Left            =   3120
      TabIndex        =   23
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "R"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   18
      Left            =   3000
      TabIndex        =   22
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Q"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   17
      Left            =   2880
      TabIndex        =   21
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "P"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   16
      Left            =   2760
      TabIndex        =   20
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "O"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   15
      Left            =   2640
      TabIndex        =   19
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "N"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   14
      Left            =   2520
      TabIndex        =   18
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "M"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   13
      Left            =   2385
      TabIndex        =   17
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "L"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   12
      Left            =   2280
      TabIndex        =   16
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "K"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   11
      Left            =   2160
      TabIndex        =   15
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "J"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   10
      Left            =   2085
      TabIndex        =   14
      Top             =   840
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "I"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   9
      Left            =   2040
      TabIndex        =   13
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "H"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   8
      Left            =   1920
      TabIndex        =   12
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "G"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   7
      Left            =   1800
      TabIndex        =   11
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "F"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   6
      Left            =   1680
      TabIndex        =   10
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "E"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   9
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "D"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   4
      Left            =   1440
      TabIndex        =   8
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "C"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   1320
      TabIndex        =   7
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "B"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "A"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "#"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Viewing    Site"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0FF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   4485
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   3120
      Picture         =   "Form1.frx":F759
      Top             =   3600
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   960
      Picture         =   "Form1.frx":11DA7
      Top             =   360
      Width           =   2925
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
      Left            =   4350
      TabIndex        =   38
      Top             =   60
      Width           =   120
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   120
      Picture         =   "Form1.frx":1428D
      Top             =   0
      Width           =   4500
   End
   Begin VB.Image Image8 
      Height          =   150
      Left            =   240
      Picture         =   "Form1.frx":1776F
      Top             =   240
      Width           =   285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim thecount As Integer, thecol As Integer
Dim col, theoldtext As String
Dim thecat As String, rad As Integer



Sub Skipto(letter As String)
For X = 0 To thecount
theletter = Mid(List1.List(X), 1, 1)
If LCase(theletter) = LCase(letter) Then List1.ListIndex = X
If LCase(theletter) = LCase(letter) Then Exit Sub
Next X
MsgBox "There Are Currently No Sites Starting With This Letter/Number", vbInformation, App.Title
End Sub

Private Sub Command1_Click()
    ShellExecute Me.hwnd, "open", List2.List(List1.ListIndex) + Text1.Text, "", "", 1
End Sub


Private Sub Image4_Click()
List1.Clear
List2.Clear
List3.Clear
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub


Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Dim Pnt As PointAPI
    GetCursorPos Pnt
Form3.Show
Form3.Left = Pnt.X * 15
Form3.Top = Pnt.Y * 15
Form3.Visible = True
'y...up/down
Else
Exit Sub
End If
End Sub

Private Sub Label1_Change()
If Len(Label1.Caption) >= 56 Then Label1.Caption = Mid(Label1.Caption, 1, 55) & "..."
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then MsgBox "Click on the letter that represents the first letter in the name of the search engine in which you are looking for", vbExclamation, App.Title
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Command1_Click
End Sub

Private Sub Form_Load()
Me.Caption = Label4.Caption
End Sub

Private Sub Image2_Click()
Form2.Show
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormDrag(Form1)
End Sub

Private Sub Label3_Click(Index As Integer)
bubba = Label3(Index).Caption
If bubba = "#" Then bubba = "1"
Skipto (bubba)
End Sub

Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For X = 0 To 26
Label3(X).BackColor = vbBlack
Next X
Call favhigh(Label3(Index))
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormDrag(Form1)
End Sub


Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
col = Label7.ForeColor
Label5.Visible = False
Label7.ForeColor = Label5.ForeColor
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Visible = True
Label7.ForeColor = col
End
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
col = Label8.ForeColor
Label6.Visible = False
Label8.ForeColor = Label6.ForeColor
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.Visible = True
Label8.ForeColor = col
Me.WindowState = 1
End Sub

Private Sub List1_Click()
On Error Resume Next
Command1.Caption = "Search"
Text1.Enabled = True
Label1.Caption = List2.List(List1.ListIndex)
If Text1.Text = "" Then GoTo getit
theoldtext = Text1.Text
getit:
Text1.Text = theoldtext
If List3.List(List1.ListIndex) = "True" Then Text1.Enabled = False
If List3.List(List1.ListIndex) = "True" Then Label1.Caption = "This Engine Does Not Allow Outside Searching"
If List3.List(List1.ListIndex) = "True" Then Command1.Caption = "Goto Site"
If List3.List(List1.ListIndex) = "True" Then Text1.Text = ""
theletter = Mid(List1.List(List1.ListIndex), 1, 1)
Call highlet(theletter)
End Sub

Function favhigh(nameit As Label)
On Error GoTo errorH:
thecol = thecol + 1
nameit.BackColor = QBColor(thecol)
Exit Function
errorH:
thecol = 0
End Function

Private Sub Text1_Change()
    Buffer = Text1.Text
    For I = 1 To Len(Buffer)
        Select Case Asc(Mid(Buffer, I, 1))
        Case 42, 43, 45 To 57, 64 To 90, 95, 97 To 122
            CBuffer = CBuffer + Mid(Buffer, I, 1)
        Case Else
            CBuffer = CBuffer + "%" & Hex(Asc(Mid(Buffer, I, 1)))
        End Select
    Next I
    Text2.Text = CBuffer
End Sub

Private Sub Text3_Change()
If Text3.Text = "Type In The First Few Letters Of The Search Engine Your Looking For" Then Exit Sub
For X = 0 To thecount
If LCase(Text3.Text) = LCase(Mid(List1.List(X), 1, Len(Text3.Text))) Then List1.ListIndex = X
If LCase(Text3.Text) = LCase(Mid(List1.List(X), 1, Len(Text3.Text))) Then Exit Sub
Next X
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then MsgBox "Type in what you are lookig for in this box, then click on the button labeled " & """" & "Search" & """", vbExclamation, App.Title
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then MsgBox "Click on this button to refresh the list of available search engines", vbExclamation, App.Title
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then MsgBox "This label tells how many search engines are available at this time", vbExclamation, App.Title
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then MsgBox "Click on this button to search for what your looking for with the selected search engine", vbExclamation, App.Title
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then MsgBox "This is a list of all available search engines at this time." & vbCr & "Click on a search engines name to use that specific search engine", vbExclamation, App.Title
End Sub

Private Sub Text3_GotFocus()
Text3.Text = ""
End Sub

Private Sub Text3_LostFocus()
Text3.Text = "Type In The First Few Letters Of The Search Engine Your Looking For"
End Sub

Private Sub Text3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then MsgBox "Type in the name of a search engine to see if it is listed", vbExclamation, App.Title
End Sub

Private Sub Timer1_Timer()
Dim db As Database, rst As Recordset
On Error GoTo errorH
GetINIlocation
Set db = OpenDatabase(INIloc)
With db
    Set rst = .OpenRecordset("Select SearchSites.SiteName,SearchSites.SiteUrl,SearchSites.SiteSearch From SearchSites Order By SearchSites.SiteName", dbOpenDynaset)
    If rst.RecordCount = 0 Then Exit Sub
    With rst
        .Edit
Do
List1.AddItem !SiteName
List2.AddItem !SiteUrl
List3.AddItem !SiteSearch
ProgressBar1.Max = rst.RecordCount
ProgressBar1.Value = ProgressBar1.Value + 1
    rst.MoveNext
    Loop
    .Close
End With
End With
errorH:
If Err.Number = 3024 Then
MsgBox "Could not find the database[" & INIloc & "]", vbCritical, App.Title
MsgBox "The program can not run without this file", vbCritical, App.Title
MsgBox "Please move the file to this location or download a new copy from http://www.angelfire.com/biz/warez86/", vbCritical, App.Title
End
Else
thecount = rst.RecordCount
Label2.Caption = "Viewing " & rst.RecordCount & " Listings"
rst.Close
db.Close
ProgressBar1.Visible = False
Timer1.Enabled = False
End If
End Sub

Function highlet(letter)
For X = 0 To 26
Label3(X).BackColor = vbBlack
Next X
For X = 0 To 26
If Label3(X).Caption = letter Then Exit For
Next X
Label3(X).BackColor = vbBlue
End Function

Private Sub Timer2_Timer()
rad = rad + 1
Image4.Picture = ImageList1.ListImages(rad).Picture
If rad = 7 Then rad = 0
End Sub
