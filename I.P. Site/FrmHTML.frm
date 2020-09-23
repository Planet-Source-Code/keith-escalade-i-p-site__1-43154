VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ip site -- keith_escalade"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8415
   Icon            =   "FrmHTML.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Broadcast hits"
      Height          =   255
      Left            =   2760
      TabIndex        =   25
      Top             =   5880
      Width           =   1335
   End
   Begin VB.ListBox List4 
      Height          =   1815
      Left            =   7560
      TabIndex        =   21
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Create info page"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   1095
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Text            =   "FrmHTML.frx":0442
      Top             =   4680
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "-"
      Height          =   315
      Left            =   7920
      TabIndex        =   17
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "+"
      Height          =   315
      Left            =   7560
      TabIndex        =   16
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4680
      TabIndex        =   15
      Top             =   4080
      Width           =   2775
   End
   Begin VB.ListBox List3 
      Height          =   1035
      Left            =   4680
      TabIndex        =   14
      Top             =   3000
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   315
      Left            =   3840
      TabIndex        =   12
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Visit url"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.ListBox List2 
      Height          =   2400
      ItemData        =   "FrmHTML.frx":04FF
      Left            =   120
      List            =   "FrmHTML.frx":0501
      TabIndex        =   7
      Top             =   3000
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy url"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "123456789"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "1560"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Broadcast"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1560
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "HTML"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Hits 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   24
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Hit Counter"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Times"
      Height          =   255
      Left            =   7560
      TabIndex        =   22
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Blocked ips see"
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Blocked ips"
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Random list %R% replaces itself with a item from the list"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Visitors"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
On Error GoTo asdf
If Text1.Text = "" Then Check1.Value = 0: Exit Sub
If Check1.Value = 1 Then Winsock1.LocalPort = Text3.Text: Winsock1.Listen: Text2.Enabled = False: Text3.Enabled = False: Exit Sub
Winsock1.Close
Text2.Enabled = True
Text3.Enabled = True
Exit Sub
asdf:
Check1.Value = 0
Exit Sub
End Sub
Private Sub Command1_Click()
Clipboard.SetText "http://" & Winsock1.LocalIP & ":" & Text3.Text
End Sub
Private Sub Command2_Click()
If Check1.Value = 1 Then Shell "Explorer http://" & Winsock1.LocalIP & ":" & Text3.Text
End Sub
Private Sub Command3_Click()
On Error Resume Next
List2.RemoveItem List2.ListIndex
End Sub
Private Sub Command4_Click()
List2.AddItem Text4.Text
Text4.Text = ""
End Sub
Private Sub Command5_Click()
List3.AddItem Text5.Text
Text5.Text = ""
End Sub
Private Sub Command6_Click()
List3.RemoveItem List3.ListIndex
End Sub
Private Sub Command7_Click()
Me.Hide
Form2.Show
End Sub
Private Sub Form_Load()
On Error GoTo DieError
Text2.Text = "http://" & Winsock1.LocalIP & ":1560"
Open App.Path & "/HTML.html" For Input As #1
While Not EOF(1)
Input #1, sText$
sText$ = Replace(sText$, "%CommA", ",")
If EOF(1) = True Then Text1.Text = Text1.Text & sText$
If EOF(1) = False Then Text1.Text = Text1.Text & sText$ & vbCrLf
Wend
Close #1
Open App.Path & "/RandomList.lst" For Input As #1
While Not EOF(1)
Input #1, sText1$
sText1$ = Replace(sText1$, "%CommA", ",")
List2.AddItem sText1$
Wend
Close #1
Open App.Path & "/Visitors.lst" For Input As #1
While Not EOF(1)
Input #1, sText2$
sText2$ = Replace(sText2$, "%CommA", ",")
List1.AddItem sText2$
Wend
Close #1
Open App.Path & "/BlockedIps.lst" For Input As #1
While Not EOF(1)
Input #1, sText3$
sText3$ = Replace(sText3$, "%CommA", ",")
List3.AddItem sText3$
Wend
Close #1
Open App.Path & "/BlockedIpText.txt" For Input As #1
Input #1, sText4$
sText4$ = Replace(sText4$, "%CommA", ",")
Text6.Text = sText4$
Close #1
Open App.Path & "/Times.lst" For Input As #1
While Not EOF(1)
Input #1, sText5$
List4.AddItem sText5$
Wend
Close #1
Hits.Caption = List1.ListCount
Exit Sub
DieError:
Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
Open App.Path & "/HTML.html" For Output As #1
asdf1$ = Replace(Text1.Text, ",", "%CommA")
Print #1, asdf1$
Close #1
Open App.Path & "/RandomList.lst" For Output As #1
For X = 0 To List2.ListCount - 1
asdf2$ = Replace(List2.List(X), ",", "%CommA")
Print #1, asdf2$
Next X
Close #1
Open App.Path & "/Visitors.lst" For Output As #1
For X = 0 To List1.ListCount - 1
asdf$ = Replace(List1.List(X), ",", "%CommA")
Print #1, List1.List(X)
Next X
Close #1
Open App.Path & "/BlockedIps.lst" For Output As #1
For X = 0 To List3.ListCount - 1
asdf3$ = Replace(List3.List(X), ",", "%CommA")
Print #1, asdf3$
Next X
Close #1
Open App.Path & "/BlockedIpText.txt" For Output As #1
asdf4$ = Replace(Text6.Text, ",", "%CommA")
Print #1, asdf4$
Close #1
Open App.Path & "/Times.lst" For Output As #1
For I = 0 To List4.ListCount - 1
Print #1, List4.List(I)
Next I
Close #1
Quit
End Sub

Private Sub Text3_Change()
Text2.Text = "http://" & Winsock1.LocalIP & ":" & Text3.Text
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command4_Click
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command5_Click
End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
For X = 0 To List3.ListCount - 1
If List3.List(X) = Winsock1.RemoteHostIP Then Winsock1.SendData Text6.Text: Exit Sub
Next X
Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
If Randd$ = "" Then Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
If Randd$ = "" Then Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
If Randd$ = "" Then Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
If Randd$ = "" Then Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
If Randd$ = "" Then Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
If Randd$ = "" Then Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
If Randd$ = "" Then Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
If Randd$ = "" Then Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
If Randd$ = "" Then Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
If Randd$ = "" Then Randd$ = List2.List(RandomNumber(List2.ListCount, 0))
Randd$ = Replace(Text1.Text, "%R%", Randd$)
If Check2.Value = 0 Then Winsock1.SendData Randd$
If Check2.Value = 1 Then Winsock1.SendData Randd$ & "<br><br><br><font size = ""1"">Hit Counter " & Hits.Caption
For X = 0 To List3.ListCount - 1
If List1.List(X) = Winsock1.RemoteHostIP Then Exit Sub
Next X
For X = 0 To List1.ListCount - 1
If Winsock1.RemoteHostIP = List1.List(X) Then List4.List(X) = List4.List(X) + 1: Exit Sub
Next X
List1.AddItem Winsock1.RemoteHostIP
List4.AddItem "1"
Hits.Caption = List1.ListCount
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Description = "Connection is aborted due to timeout or other failure" Then Exit Sub
Check1.Value = 0
Winsock1.Close
Me.Hide
MsgBox "ERROR: " & Description
Me.Show
Text2.Enabled = True
Text3.Enabled = True
End Sub
Private Sub Winsock1_SendComplete()
Pause 1#
Winsock1.Close
Winsock1.Listen
End Sub
