VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gamename"
   ClientHeight    =   2400
   ClientLeft      =   10650
   ClientTop       =   1680
   ClientWidth     =   2865
   Icon            =   "highscores.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   2865
   Begin ScoreCheater.http http1 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ProxyPort       =   0
      Timeout         =   0
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Score"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Score Cheater Created by Genetix
'Copyright 2004
'Removal in this copyright notice will result in prosecution
'Distribution of this source code is illegal
Dim htmldata As String
Private Sub http1_FileLoaded(FileContent As String, FileSize As Long)
On Error Resume Next
List1.Clear
List2.Clear
htmlbody = http1.htmldata
lngStartPoint = InStr(htmlbody, "<td align=center>1</td>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "<td align=center>20</td>")
htmlbody = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
lngStartPoint = InStr(htmlbody, "?user=")
lngStopPoint = InStr(lngStartPoint, htmlbody, "'")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "?user=", "")
List1.AddItem beforeadding

htmlbody = Replace(htmlbody, "<td align=center><", "")
htmlbody = Replace(htmlbody, "<td align=center>1</td>", "")
lngStartPoint = InStr(htmlbody, "<td align=center>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "</td")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "<td align=center>", "")
List2.AddItem beforeadding


lngStartPoint = InStr(htmlbody, "?user=")
lngStopPoint = InStr(lngStartPoint, htmlbody, "'")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "?user=", "")
List1.AddItem beforeadding

htmlbody = Replace(htmlbody, "<td align=center><", "")
htmlbody = Replace(htmlbody, "<td align=center>2</td>", "")
lngStartPoint = InStr(htmlbody, "<td align=center>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "</td")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "<td align=center>", "")
List2.AddItem beforeadding



lngStartPoint = InStr(htmlbody, "?user=")
lngStopPoint = InStr(lngStartPoint, htmlbody, "'")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "?user=", "")
List1.AddItem beforeadding

htmlbody = Replace(htmlbody, "<td align=center><", "")
htmlbody = Replace(htmlbody, "<td align=center>3</td>", "")
lngStartPoint = InStr(htmlbody, "<td align=center>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "</td")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "<td align=center>", "")
List2.AddItem beforeadding



lngStartPoint = InStr(htmlbody, "?user=")
lngStopPoint = InStr(lngStartPoint, htmlbody, "'")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "?user=", "")
List1.AddItem beforeadding

htmlbody = Replace(htmlbody, "<td align=center><", "")
htmlbody = Replace(htmlbody, "<td align=center>4</td>", "")
lngStartPoint = InStr(htmlbody, "<td align=center>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "</td")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "<td align=center>", "")
List2.AddItem beforeadding


lngStartPoint = InStr(htmlbody, "?user=")
lngStopPoint = InStr(lngStartPoint, htmlbody, "'")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "?user=", "")
List1.AddItem beforeadding

htmlbody = Replace(htmlbody, "<td align=center><", "")
htmlbody = Replace(htmlbody, "<td align=center>5</td>", "")
lngStartPoint = InStr(htmlbody, "<td align=center>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "</td")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "<td align=center>", "")
List2.AddItem beforeadding

lngStartPoint = InStr(htmlbody, "?user=")
lngStopPoint = InStr(lngStartPoint, htmlbody, "'")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "?user=", "")
List1.AddItem beforeadding

htmlbody = Replace(htmlbody, "<td align=center><", "")
htmlbody = Replace(htmlbody, "<td align=center>6</td>", "")
lngStartPoint = InStr(htmlbody, "<td align=center>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "</td")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "<td align=center>", "")
List2.AddItem beforeadding

lngStartPoint = InStr(htmlbody, "?user=")
lngStopPoint = InStr(lngStartPoint, htmlbody, "'")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "?user=", "")
List1.AddItem beforeadding

htmlbody = Replace(htmlbody, "<td align=center><", "")
htmlbody = Replace(htmlbody, "<td align=center>7</td>", "")
lngStartPoint = InStr(htmlbody, "<td align=center>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "</td")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "<td align=center>", "")
List2.AddItem beforeadding

lngStartPoint = InStr(htmlbody, "?user=")
lngStopPoint = InStr(lngStartPoint, htmlbody, "'")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "?user=", "")
List1.AddItem beforeadding

htmlbody = Replace(htmlbody, "<td align=center><", "")
htmlbody = Replace(htmlbody, "<td align=center>8</td>", "")
lngStartPoint = InStr(htmlbody, "<td align=center>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "</td")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "<td align=center>", "")
List2.AddItem beforeadding

lngStartPoint = InStr(htmlbody, "?user=")
lngStopPoint = InStr(lngStartPoint, htmlbody, "'")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "?user=", "")
List1.AddItem beforeadding

htmlbody = Replace(htmlbody, "<td align=center><", "")
htmlbody = Replace(htmlbody, "<td align=center>9</td>", "")
lngStartPoint = InStr(htmlbody, "<td align=center>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "</td")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "<td align=center>", "")
List2.AddItem beforeadding

lngStartPoint = InStr(htmlbody, "?user=")
lngStopPoint = InStr(lngStartPoint, htmlbody, "'")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "?user=", "")
List1.AddItem beforeadding

htmlbody = Replace(htmlbody, "<td align=center>10</td>", "")
lngStartPoint = InStr(htmlbody, "<td align=center>")
lngStopPoint = InStr(lngStartPoint, htmlbody, "</td")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
htmlbody = Replace(htmlbody, beforeadding, "")
beforeadding = Replace(beforeadding, "<td align=center>", "")
List2.AddItem beforeadding

End Sub
