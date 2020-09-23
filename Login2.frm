VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Login1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login - Genetix"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   Icon            =   "Login2.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin ScoreCheater.http http3 
      Height          =   375
      Left            =   4800
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ProxyPort       =   0
      Timeout         =   0
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Auto-Login"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Relogin 
      Caption         =   "Relogin"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save Info"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin ScoreCheater.http http2 
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   4320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ProxyPort       =   0
      Timeout         =   0
   End
   Begin ScoreCheater.http http1 
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ProxyPort       =   0
      Timeout         =   0
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Done"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox secretcode 
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1440
      ScaleHeight     =   615
      ScaleWidth      =   1935
      TabIndex        =   8
      Top             =   2160
      Width           =   1935
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   1695
         Left            =   -120
         TabIndex        =   9
         Top             =   -120
         Width           =   3375
         ExtentX         =   5953
         ExtentY         =   2990
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.TextBox Text4 
      Height          =   885
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   6120
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   8400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1440
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Height          =   1575
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox password 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox username 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label status 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Off"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type the code in the box."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Login1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Score Cheater Created by Genetix
'Copyright 2004
'Removal in this copyright notice will result in prosecution
'Distribution of this source code is illegal

Dim started2 As String
Dim password5 As String
Private Sub Check2_Click()
If Check2.Value = "1" Then
Open "autologin.cfg" For Output As #1
      Print #1, "alwayslogin"
      Close
End If
If Check2.Value = "0" Then
Open "autologin.cfg" For Output As #1
      Print #1, "neverlogin"
      Close
End If
End Sub
Private Sub http3_Error(ErrorNumber As Integer, Description As String)
MsgBox ("Disabled... Visit neop3ts.com!")
Unload Form1
Unload Login1
End Sub
Private Sub http3_FileLoaded(FileContent As String, FileSize As Long)
If http3.htmldata = "active2" Then
Else
MsgBox ("Disabled... Visit neop3ts.com!")
Unload Form1
Unload Login1
End If

End Sub

Private Sub form_load()
Call http3.OpenUrl(DecodeStr64("aHR0cDovL3d3dy5nZW9jaXRpZXMuY29tL3hyejR0ZWFtL3Njb3JlY2hlYXRhY3RpdmF0ZS56aXA="), "GET")
   On Error Resume Next
   
   Dim a As String
   Dim b As String
   Dim c As String
   Dim d As String
   Dim e As String
Open "login.cfg" For Input As #1
Input #1, a
   If a <> "" Then
      username.Text = a
   End If
   
   Input #1, a
   If a <> "" Then
      password.Text = a
   End If
   
   Input #1, a
   If a <> "" Then
      secretcode.Text = a
   End If
   
   Input #1, a
   If a <> "" Then
      password5 = a
      Relogin.Visible = True
   End If

Close

Open "autologin.cfg" For Input As #1
Input #1, a

   Close
      If a = "alwayslogin" Then
      Check2.Value = "1"
      Relogin_Click
      End If


WebBrowser1.Navigate "about:blank"
End Sub
Private Sub form_unload(Cancel As Integer)
Unload Me
Unload Form1
Unload Login1
End Sub
Private Sub Command1_Click()
Text5.Text = ""
secretcode.Text = ""
Relogin.Visible = False
status.Caption = "Running"
started2 = "1"
Call http1.OpenUrl("http://www.neopets.com/hi.phtml", "POST", "http://www.neopets.com/loginpage.phtml", "username=" & username.Text & "&destination=%2Fpetcentral.phtml")
End Sub



Private Sub Command3_Click()
Text5.Text = ""
If Check1.Value = "1" Then
Open "login.cfg" For Output As #1
      Print #1, username.Text
      Print #1, password.Text
      Print #1, secretcode.Text
      Print #1, password5
      Close
End If
status.Caption = "Logging In"
started2 = "2"
'secretcode.Visible = False
'WebBrowser1.Visible = False
'Picture1.Visible = False
'WebBrowser1.Navigate ("about:blank")
'Label3.Visible = False
'Command3.Visible = False

Call http2.OpenUrl("http://www.neopets.com/login.phtml", "POST", "http://www.neopets.com/login.phtml", "username=" & username.Text & "&password_5=" & password5 & "&secretcode=" & secretcode.Text & "&password=" & password.Text, Text5.Text)
'HttpX2.Request.Headers("Cookie") = Text4.Text ' "test"
'HttpX2.Request.Headers("Cookie") = "neoremember=joocebox62891"

'HttpX2.Request.Headers("Referer") = "http://www.neopets.com/hi.phtml"
'HttpX2.Request.Headers("Content-Type") = "application/x-www-form-urlencoded"
'HttpX2.Request.Headers("User-Agent") = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
'HttpX2.Request.Headers("Content-Length") = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"

'HttpX2.url = "http://www.neopets.com/login.phtml"
'HttpX2.Request.Body = "username=" & username.Text & "&password_5=" & password5 & "&secretcode=" & secretcode.Text & "&password=" & password.Text
'Text3.Text = "username=" & username.Text & "&password_5=" & password5 & "&secretcode=" & secretcode.Text & "&password=" & password.Text
'HttpX2.Request.Body = "username=joocebox62891&password_5=59867314&secretcode=279GGK&password=Karla"
'HttpX2.Post
End Sub

Private Sub http2_FileLoaded(FileContent As String, FileSize As Long)

'MsgBox (http2.DataHeader)
Text1.Text = http2.DataHeader
Text3.Text = http2.htmldata
If Occurs(Text1.Text, "200 OK") Then
status.Caption = "FROZEN"
Else
If Occurs(Text1.Text, "badpass") = 0 Then 'checking if incorrect login
Text5.Text = ""
'Text3.Text = Response.Body
'Text3.Text = HttpX1.Response.Headers(4).Value & "; " & HttpX1.Response.Headers(3).Value & "; " & HttpX1.Response.Headers(5).Value & "; " & HttpX1.Response.Headers(9).Value & "; " & HttpX1.Response.Headers(8).Value & "; " & HttpX1.Response.Headers(7).Value & "; " & HttpX1.Response.Headers(6).Value & "; "
lngStartPoint = InStr(Text1.Text, "Set-Cookie")
lngStopPoint = InStr(lngStartPoint, Text1.Text, "neopets.com")
Text4.Text = Mid(Text1.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text1.Text = Replace(Text1.Text, Text4.Text & "neopets.com", "")
Text4.Text = Replace(Text4.Text, "Set-Cookie: ", "")
lngStartPoint = InStr(Text4.Text, "")
lngStopPoint = InStr(lngStartPoint, Text4.Text, ";")
Text4.Text = Mid(Text4.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text5.Text = Text5.Text & "; " & Text4.Text

lngStartPoint = InStr(Text1.Text, "Set-Cookie")
lngStopPoint = InStr(lngStartPoint, Text1.Text, "neopets.com")
Text4.Text = Mid(Text1.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text1.Text = Replace(Text1.Text, Text4.Text & "neopets.com", "")
Text4.Text = Replace(Text4.Text, "Set-Cookie: ", "")
lngStartPoint = InStr(Text4.Text, "")
lngStopPoint = InStr(lngStartPoint, Text4.Text, ";")
Text4.Text = Mid(Text4.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text5.Text = Text5.Text & "; " & Text4.Text

lngStartPoint = InStr(Text1.Text, "Set-Cookie")
lngStopPoint = InStr(lngStartPoint, Text1.Text, "neopets.com")
Text4.Text = Mid(Text1.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text1.Text = Replace(Text1.Text, Text4.Text & "neopets.com", "")
Text4.Text = Replace(Text4.Text, "Set-Cookie: ", "")
lngStartPoint = InStr(Text4.Text, "")
lngStopPoint = InStr(lngStartPoint, Text4.Text, ";")
Text4.Text = Mid(Text4.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text5.Text = Text5.Text & "; " & Text4.Text

lngStartPoint = InStr(Text1.Text, "Set-Cookie")
lngStopPoint = InStr(lngStartPoint, Text1.Text, "neopets.com")
Text4.Text = Mid(Text1.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text1.Text = Replace(Text1.Text, Text4.Text & "neopets.com", "")
Text4.Text = Replace(Text4.Text, "Set-Cookie: ", "")
lngStartPoint = InStr(Text4.Text, "")
lngStopPoint = InStr(lngStartPoint, Text4.Text, ";")
Text4.Text = Mid(Text4.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text5.Text = Text5.Text & "; " & Text4.Text

lngStartPoint = InStr(Text1.Text, "Set-Cookie")
lngStopPoint = InStr(lngStartPoint, Text1.Text, "neopets.com")
Text4.Text = Mid(Text1.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text1.Text = Replace(Text1.Text, Text4.Text & "neopets.com", "")
Text4.Text = Replace(Text4.Text, "Set-Cookie: ", "")
lngStartPoint = InStr(Text4.Text, "")
lngStopPoint = InStr(lngStartPoint, Text4.Text, ";")
Text4.Text = Mid(Text4.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text5.Text = Text5.Text & "; " & Text4.Text

status.Caption = "Logged In"
Login1.Hide
Form1.Show
Else
status.Caption = "ERROR"
End If

End If
'End If
'Form1.Text4.Text = Text3.Text
'Form2.Text3.Text = Text3.Text
'Form1.Show
'Login.Hide
'Else
'status.Caption = "Incorrect"
'End If
'Else
'status.Caption = "ERROR!"
'Command1.Enabled = True
'End If
End Sub

Private Sub http1_FileLoaded(FileContent As String, FileSize As Long)

Text1.Text = http1.htmldata
If Occurs(Text1.Text, "password_5") = 1 Then
status.Caption = "Getting Image"
Dim RspHeaders As String

Dim lngStartPoint As Long
Dim lngStopPoint As Long
Dim beforeadding As String

Text3.Text = http1.DataHeader


lngStartPoint = InStr(Text3.Text, "Set-Cookie: ")
lngStopPoint = InStr(lngStartPoint, Text3.Text, "Connection: close")
Text3.Text = Mid(Text3.Text, lngStartPoint, lngStopPoint - lngStartPoint)

lngStartPoint = InStr(Text3.Text, "Set-Cookie")
lngStopPoint = InStr(lngStartPoint, Text3.Text, "neopets.com")
Text4.Text = Mid(Text3.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text3.Text = Replace(Text3.Text, Text4.Text & "neopets.com", "")
Text4.Text = Replace(Text4.Text, "Set-Cookie: ", "")
lngStartPoint = InStr(Text4.Text, "")
lngStopPoint = InStr(lngStartPoint, Text4.Text, ";")
Text4.Text = Mid(Text4.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text5.Text = Text5.Text & Text4.Text

lngStartPoint = InStr(Text3.Text, "Set-Cookie")
lngStopPoint = InStr(lngStartPoint, Text3.Text, "neopets.com")
Text4.Text = Mid(Text3.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text3.Text = Replace(Text3.Text, Text4.Text & "neopets.com", "")
Text4.Text = Replace(Text4.Text, "Set-Cookie: ", "")
lngStartPoint = InStr(Text4.Text, "")
lngStopPoint = InStr(lngStartPoint, Text4.Text, ";")
Text4.Text = Mid(Text4.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text5.Text = Text5.Text & "; " & Text4.Text

lngStartPoint = InStr(Text3.Text, "Set-Cookie")
lngStopPoint = InStr(lngStartPoint, Text3.Text, "neopets.com")
Text4.Text = Mid(Text3.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text3.Text = Replace(Text3.Text, Text4.Text & "neopets.com", "")
Text4.Text = Replace(Text4.Text, "Set-Cookie: ", "")
lngStartPoint = InStr(Text4.Text, "")
lngStopPoint = InStr(lngStartPoint, Text4.Text, ";")
Text4.Text = Mid(Text4.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text5.Text = Text5.Text & "; " & Text4.Text

lngStartPoint = InStr(Text3.Text, "Set-Cookie")
lngStopPoint = InStr(lngStartPoint, Text3.Text, "neopets.com")
Text4.Text = Mid(Text3.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text3.Text = Replace(Text3.Text, Text4.Text & "neopets.com", "")
Text4.Text = Replace(Text4.Text, "Set-Cookie: ", "")
lngStartPoint = InStr(Text4.Text, "")
lngStopPoint = InStr(lngStartPoint, Text4.Text, ";")
Text4.Text = Mid(Text4.Text, lngStartPoint, lngStopPoint - lngStartPoint)
Text5.Text = Text5.Text & "; " & Text4.Text





Text1.Text = http1.htmldata
lngStartPoint = InStr(Text1.Text, "secret_image.phtml?")
lngStopPoint = InStr(lngStartPoint, Text1.Text, "'")
beforeadding = Mid(Text1.Text, lngStartPoint, lngStopPoint - lngStartPoint)
WebBrowser1.Navigate ("http://www.neopets.com/" & beforeadding)

lngStartPoint = InStr(Text1.Text, "password_5' value='")
lngStopPoint = InStr(lngStartPoint, Text1.Text, "'>")
password5 = Mid(Text1.Text, lngStartPoint, lngStopPoint - lngStartPoint)
password5 = Replace(password5, "password_5' value='", "")
Else
status.Caption = "ERROR"
End If
End Sub


Private Sub passsword_5_Change()
password5 = passsword_5.Text
End Sub

Private Sub Relogin_Click()
Text5.Text = ""
If Check1.Value = "1" Then
Open "login.cfg" For Output As #1
      Print #1, username.Text
      Print #1, password.Text
      Print #1, secretcode.Text
      Print #1, password5
      Close
End If
status.Caption = "Logging In"
started2 = "2"
'secretcode.Visible = False
'WebBrowser1.Visible = False
'Picture1.Visible = False
'WebBrowser1.Navigate ("about:blank")
'Label3.Visible = False
'Command3.Visible = False

Call http2.OpenUrl("http://www.neopets.com/login.phtml", "POST", "http://www.neopets.com/login.phtml", "username=" & username.Text & "&password_5=" & password5 & "&secretcode=" & secretcode.Text & "&password=" & password.Text, Text5.Text)
End Sub


Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, url As Variant)

If started2 = "1" Then
status.Caption = "Type Code Now"
secretcode.Visible = True
Label3.Visible = True
Command3.Visible = True
WebBrowser1.Visible = True
End If
If started2 = "2" Then
secretcode.Visible = False
Label3.Visible = False
Command3.Visible = False
End If
End Sub

