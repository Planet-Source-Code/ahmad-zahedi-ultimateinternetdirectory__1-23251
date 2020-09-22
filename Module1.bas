Attribute VB_Name = "Module1"
Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory _
As String, ByVal nShowCmd As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Global INIloc As String
Global INIloc2 As String
Type PointAPI
    X As Long
    Y As Long
    End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Public Function GetINIlocation()
On Error GoTo errorH
Dim theDIR1 As String, thedir2 As String, bubba As String
theDIR1 = "\ud.mdb"
thedir2 = "\"
redoit:
bubba = Dir(App.Path & theDIR1)
INIloc = App.Path & theDIR1
Exit Function
errorH:
theDIR1 = "ud.mdb"
INIloc = App.Path & theDIR1
thedir2 = ""
GoTo redoit
End Function

Public Function GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String) As Boolean
    ' Written by: Blake Pell
    
    On Local Error GoTo 100
    
    Dim myData() As Byte
    If Inet1.StillExecuting = True Then Exit Function
    myData() = Inet1.OpenURL(myURL, icByteArray)


    For X = Len(myURL) To 1 Step -1
        If Left$(Right$(myURL, X), 1) = "/" Then RealFile$ = Right$(myURL, X - 1)
    Next X
    myFile$ = DestDIR + "\" + RealFile$
    Open myFile$ For Binary Access Write As #1
    Put #1, , myData()
    Close #1
    
    GetInternetFile = True
    Exit Function

' error handler
100 X = MsgBox("An error has occured in the file transfer or write.  Please try again later.", vbInformation)
    GetInternetFile = False
    Resume 105
105 End Function

Public Function Getverlocation()
On Error GoTo errorH
Dim theDIR1 As String, thedir2 As String, bubba As String
theDIR1 = "\data.ver"
thedir2 = "\"
redoit:
bubba = Dir(App.Path & theDIR1)
INIloc2 = App.Path & theDIR1
Exit Function
errorH:
theDIR1 = "data.ver"
INIloc2 = App.Path & theDIR1
thedir2 = ""
GoTo redoit
End Function

Public Function update()
Form1.Label1.Caption = "Checking For Database Updates..."
Getverlocation
Form1.RichTextBox1.LoadFile (INIloc2)
Form1.Inet1.AccessType = icUseDefault
    Form1.Text4.Text = Form1.Inet1.OpenURL("www.angelfire.com/biz/warez86/data.ver")
If Form1.RichTextBox1.Text >= Form1.Text4.Text Then
Form1.Label1.Caption = "There Are Currently No Database Updates Available"
Call update2
Else
Form1.Label1.Caption = "An Updated Database Is Available"
Form1.Label1.Caption = "Getting Version Information"
Call GetInternetFile(Form1.Inet1, "www.angelfire.com/biz/warez86/data.ver", App.Path)
Form1.Label1.Caption = "Getting And Installing New Database"
Call GetInternetFile(Form1.Inet1, "www.angelfire.com/biz/warez86/ud.mdb", App.Path)
Form1.Label1.Caption = "New Database Installed"
Call update2
End If
End Function

Public Function update2()
On Error GoTo errorH
Form1.Label1.Caption = "Checking For Program Updates..."
Form1.Inet1.AccessType = icUseDefault
    Form1.Text4.Text = Form1.Inet1.OpenURL("www.angelfire.com/biz/warez86/program.ver")
If App.Major >= Form1.Text4.Text Then
Form1.Label1.Caption = "There Are Currently No Program Updates Available"
Else
Form1.Label1.Caption = "An Updated Program Is Available"
MsgBox "Make browser goto homepage"
End If
Exit Function
errorH:
Form1.Label1.Caption = "Error While Checking For Updates"
MsgBox "There was an error while checking for updates" & vbCr & "Make sure that you are connected to the internet and there are no programs blocking internet communications", vbCritical, "[ERROR]"
Exit Function
End Function
