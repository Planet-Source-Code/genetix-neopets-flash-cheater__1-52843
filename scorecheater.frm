VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Score Cheater - Genetix - Build 4"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   Icon            =   "scorecheater.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "High Scores"
      Height          =   255
      Left            =   7440
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4575
      ScaleWidth      =   8655
      TabIndex        =   9
      Top             =   960
      Width           =   8655
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   4575
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   8895
         ExtentX         =   15690
         ExtentY         =   8070
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
         Location        =   "http:///"
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   2295
   End
   Begin ScoreCheater.http http2 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ProxyPort       =   0
      Timeout         =   25
   End
   Begin ScoreCheater.http http1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ProxyPort       =   0
      Timeout         =   25
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Instead of"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Play"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label status 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Off"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Score Cheater Created by Genetix
'Copyright 2004
'Removal in this copyright notice will result in prosecution
'Distribution of this source code is illegal
Dim accountcookies As String
Dim htmlbody As String
Dim beforeadding As String
Dim play As String
Dim insteadof As String
Dim gameid As String
Dim pageid As String
Dim gameid2 As String
Dim checkoccurs As Integer
Dim started As String

Public Function Occurs(ByVal strtochk As String, ByVal searchstr As String) As Long
    ' remember SPLIT returns a zero-based ar
    '     ray
    Occurs = UBound(Split(strtochk, searchstr))
End Function

Private Sub Combo1_Click()
If Combo1.Text = "Hubrid's Heist" Then gameid2 = "g314_v12_39358"
If Combo1.Text = "Evil Thade" Then gameid2 = "g230_v19"
If Combo1.Text = "Chemistry" Then gameid2 = "g239_v10"
If Combo1.Text = "Codebreakers" Then gameid2 = "g2_v18"
If Combo1.Text = "Deckswaber" Then gameid2 = "g19_v18"
If Combo1.Text = "Destruct-O-Match" Then gameid2 = "g53_v19"
If Combo1.Text = "Gadgadsgame" Then gameid2 = "g159_v14"
If Combo1.Text = "Kiko Match II" Then gameid2 = "g93_v22"
If Combo1.Text = "Marble Men" Then gameid2 = "g201_v15"
If Combo1.Text = "Meepit Juice Break" Then gameid2 = "g379_v9_80428"
If Combo1.Text = "Maths Nightmare" Then gameid2 = "g150_v12"
If Combo1.Text = "Spell or Starve" Then gameid2 = "g202_v14"
If Combo1.Text = "Sutek's Tomb" Then gameid2 = "g306_v11_92289"
If Combo1.Text = "Toybox" Then gameid2 = "g367_v2_24972"
If Combo1.Text = "200 Meter Dash" Then gameid2 = "g189_v8"
If Combo1.Text = "Advert Attack" Then gameid2 = "g204_v23"
If Combo1.Text = "Bumper Cars" Then gameid2 = "g61_v18"
If Combo1.Text = "Honey O Throw" Then gameid2 = "g347_v14_22918"
If Combo1.Text = "Carnival of Terror" Then gameid2 = "g131_v14"
If Combo1.Text = "Chia Bomber" Then gameid2 = "g62_v16"
If Combo1.Text = "Chomby & the Fungus Balls" Then gameid2 = "g49_v10"
If Combo1.Text = "Deckball" Then gameid2 = "g82_v24"
If Combo1.Text = "Dubloon Disaster" Then gameid2 = "g143_v15"
If Combo1.Text = "Escape from Meridell" Then gameid2 = "g197_v7"
If Combo1.Text = "Evil Fuzzles" Then gameid2 = "g128_v15"
If Combo1.Text = "Extreme Herder" Then gameid2 = "g149_v14"
If Combo1.Text = "Xtreme Potato Counter" Then gameid2 = "g226_v17"
If Combo1.Text = "Faerie Bubbles" Then gameid2 = "g358_v14_17206"
If Combo1.Text = "Faerie Cloud Racers" Then gameid2 = "g137_v15"
If Combo1.Text = "Feed Florg" Then gameid2 = "g156_v12"
If Combo1.Text = "Grand Theft" Then gameid2 = "g212_v14"
If Combo1.Text = "Hasee Bounce" Then gameid2 = "g368_v31_40827"
If Combo1.Text = "Snowthrow" Then gameid2 = "g31_v18"
If Combo1.Text = "Ice Cream Factory" Then gameid2 = "g57_v19"
If Combo1.Text = "Igloo Garage Sale" Then gameid2 = "g169_v4"
If Combo1.Text = "Jelly Processing" Then gameid2 = "g95_v13"
If Combo1.Text = "Kenny the Shark" Then gameid2 = "g378_v9_64214"
If Combo1.Text = "Korbats Lab" Then gameid2 = "g85_v12"
If Combo1.Text = "Magmax" Then gameid2 = "g162_v8"
If Combo1.Text = "Magic Mates" Then gameid2 = "g325_v5_78694"
If Combo1.Text = "Meerca Chase" Then gameid2 = "g46_v38"
If Combo1.Text = "Meriball" Then gameid2 = "g173_v9"
If Combo1.Text = "Mutant Graveyard" Then gameid2 = "g65_v22"
If Combo1.Text = "Mynci Volleyball" Then gameid2 = "g315_v5_10166"
If Combo1.Text = "National Neo" Then gameid2 = "g371_v2_43277"
If Combo1.Text = "Nimmos Pond" Then gameid2 = "g74_v15"
If Combo1.Text = "Petpet Rescue" Then gameid2 = "g228_v12"
If Combo1.Text = "Pterattack" Then gameid2 = "g63_v25"
If Combo1.Text = "Reese's Mini-golf" Then gameid2 = "g345_v7_36027"
If Combo1.Text = "Rink Runner" Then gameid2 = "g220_v10"
If Combo1.Text = "Skies Over Meridell" Then gameid2 = "g340_v31_91377"
If Combo1.Text = "Splat A Sloth" Then gameid2 = "g81_v14"
If Combo1.Text = "Spy Kids" Then gameid2 = "g382_v11_82934"
If Combo1.Text = "Swarm" Then gameid2 = "g66_v14"
If Combo1.Text = "Buzzer Game" Then gameid2 = "g307_v6_18026"
If Combo1.Text = "Lost Plushies" Then gameid2 = "g322_v9_31132"
If Combo1.Text = "Trix Birthday" Then gameid2 = "g357_v3_46157"
If Combo1.Text = "Tug O War" Then gameid2 = "g52_v19"
If Combo1.Text = "Turmac Roll" Then gameid2 = "g366_v2_53506"
If Combo1.Text = "Ultimate Bullseyes" Then gameid2 = "g152_v15"
If Combo1.Text = "Usuki Frenzy" Then gameid2 = "g129_v11"
If Combo1.Text = "Volcano Run" Then gameid2 = "g140_v10"
If Combo1.Text = "Web of Vernax" Then gameid2 = "g353_v6_27634"
If Combo1.Text = "Warf Rescue" Then gameid2 = "g305_v6_76735"
If Combo1.Text = "Whack A Kass" Then gameid2 = "g381_v35_67047"
If Combo1.Text = "Zurroball" Then gameid2 = "g207_v21"

End Sub

Private Sub Combo2_Click()

If Combo2.Text = "Hubrid's Heist" Then
gameid = "g314_v12_39358"
pageid = "wizards.phtml"
End If

If Combo2.Text = "Evil Thade" Then
gameid = "g230_v19"
pageid = "elivthade.phtml"
End If
If Combo2.Text = "Chemistry" Then
gameid = "g239_v10"
pageid = "chemistry.phtml"
End If
If Combo2.Text = "Codebreakers" Then
gameid = "g2_v18"
pageid = "codebreakers.phtml"
End If
If Combo2.Text = "Deckswaber" Then
gameid = "g19_v18"
pageid = "deckswabber.phtml"
End If
If Combo2.Text = "Destruct-O-Match" Then
gameid = "g53_v19"
pageid = "destructomatch.phtml"
End If
If Combo2.Text = "Gadgadsgame" Then
gameid = "g159_v14"
pageid = "gadgadsgame.phtml"
End If
If Combo2.Text = "Kiko Match II" Then
gameid = "g93_v22"
pageid = "kikomatch.phtml"
End If
If Combo2.Text = "Marble Men" Then
gameid = "g201_v15"
pageid = "marblemen.phtml"
End If
If Combo2.Text = "Maths Nightmare" Then
gameid = "g150_v12"
pageid = "mathsnightmare.phtml"
End If
If Combo2.Text = "Meepit Juice Break" Then
gameid = "g379_v9_80428"
pageid = "meepits.phtml"
End If
If Combo2.Text = "Spell or Starve" Then
gameid = "g202_v14"
pageid = "spellorstarve.phtml"
End If
If Combo2.Text = "Sutek's Tomb" Then
gameid = "g306_v11_92289"
pageid = "sutekstomb.phtml"
End If
If Combo2.Text = "Toybox" Then
gameid = "g367_v2_24972"
pageid = "toyboxescape.phtml"
End If



If Combo2.Text = "200 Meter Dash" Then
gameid = "g189_v8"
pageid = "puppyblew.phtml"
End If
If Combo2.Text = "Advert Attack" Then
gameid = "g204_v23"
pageid = "advertattack.phtml"
End If
If Combo2.Text = "Bumper Cars" Then
gameid = ""
pageid = "bumpercars.phtml"
End If
If Combo2.Text = "Honey O Throw" Then
gameid = "g347_v14_22918"
pageid = "hnc.phtml"
End If
If Combo2.Text = "Carnival of Terror" Then
gameid = "g131_v14"
pageid = "carnival.phtml"
End If
If Combo2.Text = "Chia Bomber" Then
gameid = "g62_v16"
pageid = "chiabomber.phtml"
End If
If Combo2.Text = "Chomby & the Fungus Balls" Then
gameid = "g49_v10"
pageid = "fungusball.phtml"
End If
If Combo2.Text = "Deckball" Then
gameid = "g82_v24"
pageid = "deckball.phtml"
End If
If Combo2.Text = "Dubloon Disaster" Then
gameid = "g143_v15"
pageid = "dubloondisaster.phtml"
End If
If Combo2.Text = "Escape from Meridell" Then
gameid = "g197_v7"
pageid = "draikcastle.phtml"
End If
If Combo2.Text = "Evil Fuzzles" Then
gameid = "g128_v15"
pageid = "fuzzles.phtml"
End If
If Combo2.Text = "Extreme Herder" Then
gameid = "g149_v14"
pageid = "extremeherder.phtml"
End If
If Combo2.Text = "Xtreme Potato Counter" Then
gameid = "g226_v17"
pageid = "epc.phtml"
End If
If Combo2.Text = "Faerie Bubbles" Then
gameid = "g358_v14_17206"
pageid = "faeriebubbles.phtml"
End If
If Combo2.Text = "Faerie Cloud Racers" Then
gameid = "g137_v15"
pageid = "cloudracersphtml"
End If
If Combo2.Text = "Feed Florg" Then
gameid = "g156_v12"
pageid = "feedflorg.phtml"
End If
If Combo2.Text = "Grand Theft" Then
gameid = "g212_v14"
pageid = "grandtheft.phtml"
End If
If Combo2.Text = "Hasee Bounce" Then
gameid = "g368_v31_40827"
pageid = "haseebounce.phtml"
End If
If Combo2.Text = "Snowthrow" Then
gameid = "g31_v18"
pageid = "snowthrow.phtml"
End If
If Combo2.Text = "Ice Cream Factory" Then
gameid = "g57_v19"
pageid = "icecream.phtml"
End If
If Combo2.Text = "Igloo Garage Sale" Then
gameid = "g169_v4"
pageid = "igloogarage.phtml"
End If
If Combo2.Text = "Jelly Processing" Then
gameid = "g95_v13"
pageid = "jellyprocessing.phtml"
End If
If Combo2.Text = "Kenny the Shark" Then
gameid = "g378_v9_64214"
pageid = "kennytheshark.phtml"
End If
If Combo2.Text = "Korbats Lab" Then
gameid = "g85_v12"
pageid = "korbatslab.phtml"
End If
If Combo2.Text = "Magmax" Then
gameid = "g162_v8"
pageid = "maxmax.phtml"
End If
If Combo2.Text = "Magic Mates" Then
gameid = "g325_v5_78694"
pageid = "magicmates.phtml"
End If
If Combo2.Text = "Meerca Chase" Then
gameid = "g46_v38"
pageid = "meercachase.phtml"
End If
If Combo2.Text = "Meriball" Then
gameid = "g173_v9"
pageid = "meriball.phtml"
End If
If Combo2.Text = "Mutant Graveyard" Then
gameid = "g65_v22"
pageid = "graveyard.phtml"
End If
If Combo2.Text = "Mynci Volleyball" Then
gameid = "g315_v5_10166"
pageid = "volleyball.phtml"
End If
If Combo2.Text = "National Neo" Then
gameid = "g371_v2_43277"
pageid = "nationalneo.phtml"
End If
If Combo2.Text = "Nimmos Pond" Then
gameid = "g74_v15"
pageid = "nimmospond.phtml"
End If
If Combo2.Text = "Petpet Rescue" Then
gameid = "g228_v12"
pageid = "petpetrescue.phtml"
End If
If Combo2.Text = "Pterattack" Then
gameid = "g63_v25"
pageid = "pterattack.phtml"
End If
If Combo2.Text = "Reese's Mini-golf" Then
gameid = "g345_v7_36027"
pageid = "reesespuffs.phtml"
End If
If Combo2.Text = "Rink Runner" Then
gameid = "g220_v10"
pageid = "rinkrunner.phtml"
End If
If Combo2.Text = "Skies Over Meridell" Then
gameid = "g340_v31_91377"
pageid = "biplanes.phtml"
End If
If Combo2.Text = "Splat A Sloth" Then
gameid = "g81_v14"
pageid = "splatasloth.phtml"
End If
If Combo2.Text = "Spy Kids" Then
gameid = "g382_v11_82934"
pageid = "spykids3.phtml"
End If
If Combo2.Text = "Swarm" Then
gameid = "g66_v14"
pageid = "swarm.phtml"
End If
If Combo2.Text = "Buzzer Game" Then
gameid = "g307_v6_18026"
pageid = "thebuzzergame.phtml"
End If
If Combo2.Text = "Lost Plushies" Then
gameid = "g322_v9_31132"
pageid = "plushies.phtml"
End If
If Combo2.Text = "Trix Birthday" Then
gameid = "g357_v3_46157"
pageid = "trix.phtml"
End If
If Combo2.Text = "Tug O War" Then
gameid = "g52_v19"
pageid = "tugowar.phtml"
End If
If Combo2.Text = "Turmac Roll" Then
gameid = "g366_v2_53506"
pageid = "turmacroll2.phtml"
End If
If Combo2.Text = "Ultimate Bullseyes" Then
gameid = "g152_v15"
pageid = "bullseyes.phtml"
End If
If Combo2.Text = "Usuki Frenzy" Then
gameid = "g129_v11"
pageid = "usukifrenzy.phtml"
End If
If Combo2.Text = "Volcano Run" Then
gameid = "g140_v10"
pageid = "volcanorun.phtml"
End If
If Combo2.Text = "Web of Vernax" Then
gameid = "g353_v6_27634"
pageid = "webofvernax.phtml"
End If
If Combo2.Text = "Warf Rescue" Then
gameid = "g305_v6_76735"
pageid = "warfrescueteam.phtml"
End If
If Combo2.Text = "Whack A Kass" Then
gameid = "g381_v35_67047"
pageid = "kassgame.phtml"
End If
If Combo2.Text = "Zurroball" Then
gameid = "g207_v21"
pageid = "zurroball.phtml"
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Combo2.Text = "" Then
MsgBox ("Please select a game")
Else
If Combo2.Text = "Honey O Throw" Or Combo2.Text = "Kenny the Shark" Or Combo2.Text = "Trix Birthday" Or Combo2.Text = "Reese's Mini-golf" Then
MsgBox ("No Highscores")
Else
Form2.Show
Form2.Caption = Combo2.Text
gameid3 = Replace(gameid, "g", "")
lngStartPoint = InStr(gameid3, "")
lngStopPoint = InStr(lngStartPoint, gameid3, "_")
gameid3 = Mid(gameid3, lngStartPoint, lngStopPoint - lngStartPoint)
'MsgBox (gameid3)
Call Form2.http1.OpenUrl("http://www.neopets.com/gamescores.phtml?game_id=" & gameid3, "GET", "http://www.neopets.com/games/" & pageid, "", accountcookies)
End If
End If
End Sub

Private Sub form_unload(Cancel As Integer)
Unload Form1
Unload Login1
End Sub
Private Sub form_load()
WebBrowser1.Navigate "about:blank"
Combo2.AddItem "--Puzzle Games--"
Combo2.AddItem "Evil Thade"
Combo2.AddItem "Chemistry"
Combo2.AddItem "Codebreakers"
Combo2.AddItem "Deckswaber"
Combo2.AddItem "Destruct-O-Match"
Combo2.AddItem "Gadgadsgame"
Combo2.AddItem "Kiko Match II"
Combo2.AddItem "Marble Men"
Combo2.AddItem "Maths Nightmare"
Combo2.AddItem "Meepit Juice Break"
Combo2.AddItem "Spell or Starve"
Combo2.AddItem "Sutek's Tomb"
Combo2.AddItem "Toybox"

Combo2.AddItem "--Action Games"
Combo2.AddItem "200 Meter Dash"
Combo2.AddItem "Advert Attack"
Combo2.AddItem "Bumper Cars"
Combo2.AddItem "Honey O Throw"
Combo2.AddItem "Carnival of Terror"
Combo2.AddItem "Chia Bomber"
Combo2.AddItem "Chomby & the Fungus Balls"
Combo2.AddItem "Deckball"
Combo2.AddItem "Dubloon Disaster"
Combo2.AddItem "Escape from Meridell"
Combo2.AddItem "Evil Fuzzles"
Combo2.AddItem "Extreme Herder"
Combo2.AddItem "Xtreme Potato Counter"
Combo2.AddItem "Faerie Bubbles"
Combo2.AddItem "Faerie Cloud Racers"
Combo2.AddItem "Feed Florg"
Combo2.AddItem "Grand Theft"
Combo2.AddItem "Hasee Bounce"
Combo2.AddItem "Snowthrow"
Combo2.AddItem "Ice Cream Factory"
Combo2.AddItem "Igloo Garage Sale"
Combo2.AddItem "Jelly Processing"
Combo2.AddItem "Kenny the Shark"
Combo2.AddItem "Korbats Lab"
Combo2.AddItem "Magmax"
Combo2.AddItem "Magic Mates"
Combo2.AddItem "Meerca Chase"
Combo2.AddItem "Meriball"
Combo2.AddItem "Mutant Graveyard"
Combo2.AddItem "Mynci Volleyball"
Combo2.AddItem "National Neo"
Combo2.AddItem "Nimmos Pond"
Combo2.AddItem "Petpet Rescue"
Combo2.AddItem "Pterattack"
Combo2.AddItem "Reese's Mini-golf"
Combo2.AddItem "Rink Runner"
Combo2.AddItem "Skies Over Meridell"
Combo2.AddItem "Splat A Sloth"
Combo2.AddItem "Spy Kids"
Combo2.AddItem "Swarm"
Combo2.AddItem "Buzzer Game"
Combo2.AddItem "Lost Plushies"
Combo2.AddItem "Trix Birthday"
Combo2.AddItem "Tug O War"
Combo2.AddItem "Turmac Roll"
Combo2.AddItem "Ultimate Bullseyes"
Combo2.AddItem "Usuki Frenzy"
Combo2.AddItem "Volcano Run"
Combo2.AddItem "Web of Vernax"
Combo2.AddItem "Warf Rescue"
Combo2.AddItem "Whack A Kass"
Combo2.AddItem "Zurroball"


Combo1.AddItem "Hubrid's Heist"
Combo1.AddItem "--Puzzle Games--"
Combo1.AddItem "Evil Thade"
Combo1.AddItem "Chemistry"
Combo1.AddItem "Codebreakers"
Combo1.AddItem "Deckswaber"
Combo1.AddItem "Destruct-O-Match"
Combo1.AddItem "Gadgadsgame"
Combo1.AddItem "Kiko Match II"
Combo1.AddItem "Marble Men"
Combo1.AddItem "Maths Nightmare"
Combo1.AddItem "Meepit Juice Break"
Combo1.AddItem "Spell or Starve"
Combo1.AddItem "Sutek's Tomb"
Combo1.AddItem "Toybox"

Combo1.AddItem "--Action Games"
Combo1.AddItem "200 Meter Dash"
Combo1.AddItem "Advert Attack"
Combo1.AddItem "Bumper Cars"
Combo1.AddItem "Honey O Throw"
Combo1.AddItem "Carnival of Terror"
Combo1.AddItem "Chia Bomber"
Combo1.AddItem "Chomby & the Fungus Balls"
Combo1.AddItem "Deckball"
Combo1.AddItem "Dubloon Disaster"
Combo1.AddItem "Escape from Meridell"
Combo1.AddItem "Evil Fuzzles"
Combo1.AddItem "Extreme Herder"
Combo1.AddItem "Xtreme Potato Counter"
Combo1.AddItem "Faerie Bubbles"
Combo1.AddItem "Faerie Cloud Racers"
Combo1.AddItem "Feed Florg"
Combo1.AddItem "Grand Theft"
Combo1.AddItem "Hasee Bounce"
Combo1.AddItem "Snowthrow"
Combo1.AddItem "Ice Cream Factory"
Combo1.AddItem "Igloo Garage Sale"
Combo1.AddItem "Jelly Processing"
Combo1.AddItem "Kenny the Shark"
Combo1.AddItem "Korbats Lab"
Combo1.AddItem "Magmax"
Combo1.AddItem "Magic Mates"
Combo1.AddItem "Meerca Chase"
Combo1.AddItem "Meriball"
Combo1.AddItem "Mutant Graveyard"
Combo1.AddItem "Mynci Volleyball"
Combo1.AddItem "National Neo"
Combo1.AddItem "Nimmos Pond"
Combo1.AddItem "Petpet Rescue"
Combo1.AddItem "Pterattack"
Combo1.AddItem "Reese's Mini-golf"
Combo1.AddItem "Rink Runner"
Combo1.AddItem "Skies Over Meridell"
Combo1.AddItem "Splat A Sloth"
Combo1.AddItem "Spy Kids"
Combo1.AddItem "Swarm"
Combo1.AddItem "Buzzer Game"
Combo1.AddItem "Lost Plushies"
Combo1.AddItem "Trix Birthday"
Combo1.AddItem "Tug O War"
Combo1.AddItem "Turmac Roll"
Combo1.AddItem "Ultimate Bullseyes"
Combo1.AddItem "Usuki Frenzy"
Combo1.AddItem "Volcano Run"
Combo1.AddItem "Web of Vernax"
Combo1.AddItem "Warf Rescue"
Combo1.AddItem "Whack A Kass"
Combo1.AddItem "Zurroball"
accountcookies = Login1.Text5.Text
End Sub
Private Sub Command1_Click()
On Error Resume Next
started = "yes"
If Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox ("Game not selected")
Else
Call http1.OpenUrl("http://www.neopets.com/games/" & pageid, "GET", "http://www.neopets.com/gameroom.phtml?game_type=0", "", accountcookies)
status.Caption = "Getting game page"
End If
End Sub
Private Sub http1_Error(ErrorNumber As Integer, Description As String)
status.Caption = "ERROR, retrying"
Call http1.OpenUrl("http://www.neopets.com/games/" & pageid, "GET", "http://www.neopets.com/gameroom.phtml?game_type=0", "", accountcookies)
End Sub
Private Sub http2_Error(ErrorNumber As Integer, Description As String)
status.Caption = "ERROR, retrying"
Call http2.OpenUrl(beforeadding, "GET", beforeadding, "", accountcookies)
End Sub
Private Sub http1_FileLoaded(FileContent As String, FileSize As Long)
On Error Resume Next
status.Caption = "Getting Game"
htmlbody = http1.htmldata
checkoccurs = Occurs(htmlbody, "open(")
If checkoccurs > 0 Then
lngStartPoint = InStr(htmlbody, "open('")
lngStopPoint = InStr(lngStartPoint, htmlbody, "&quality")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
beforeadding = Replace(beforeadding, "open('", "")
beforeadding = beforeadding & "&quality=high"
Call http2.OpenUrl(beforeadding, "GET", beforeadding, "", accountcookies)
Else
status.Caption = "error page!"
Call http1.OpenUrl("http://www.neopets.com/games/" & pageid, "GET", "http://www.neopets.com/gameroom.phtml?game_type=0", "", accountcookies)
End If
End Sub
Private Sub http2_FileLoaded(FileContent As String, FileSize As Long)
On Error Resume Next
checkoccurs = Occurs(htmlbody, "value=")
If checkoccurs > 0 Then
status.Caption = "Loading Game"
htmlbody = http2.htmldata
lngStartPoint = InStr(htmlbody, "value=" & Chr(34))
lngStopPoint = InStr(lngStartPoint, htmlbody, Chr(34) & ">")
beforeadding = Mid(htmlbody, lngStartPoint, lngStopPoint - lngStartPoint)
beforeadding = Replace(beforeadding, "value=" & Chr(34), "")
beforeadding = Replace(beforeadding, "g=" & gameid, "g=" & gameid2)
'MsgBox (beforeadding)
WebBrowser1.Navigate beforeadding
Else
status.Caption = "error page!"
Call http2.OpenUrl(beforeadding, "GET", beforeadding, "", accountcookies)
End If
End Sub
Private Sub WebBrowser1_DownloadComplete()
If started = "yes" Then
status.Caption = "Game Loaded"
End If
End Sub
