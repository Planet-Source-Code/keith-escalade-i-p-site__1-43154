VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create info site"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9615
   Icon            =   "FrmINFO.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   18
      Top             =   120
      Width           =   5055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   5040
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   4455
      Left            =   4440
      TabIndex        =   15
      Top             =   480
      Width           =   5055
      ExtentX         =   8916
      ExtentY         =   7858
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\system32\shdoclc.dll/dnserror.htm#http:///"
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Insert picture(not from harddrive)"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "MM/DD/YYYY"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmINFO.frx":0442
      Left            =   1080
      List            =   "FrmINFO.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "none@nowhere.com"
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4320
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Info about yourself"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Age :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Birthdate :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Sex :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "E-mail :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then Text6.Enabled = True: Exit Sub
Text6.Enabled = False
End Sub

Private Sub Command1_Click()
Form1.Show
Me.Hide
End Sub

Private Sub Command2_Click()
On Error GoTo asdf
WB1.Navigate "about:<img src = """ & Text6.Text & """ >"
Exit Sub
asdf:
Exit Sub
End Sub

Private Sub Command3_Click()
WB1.Navigate "about:People will not see the picture if it comes from your harddrive, it must be from a website. Example: http://microsoft.com/homepage/gif/bnr-microsoft.gif"
End Sub

Private Sub Command7_Click()
If Check1.Value = 1 Then Form1.Text1.Text = "<title>" & Text1.Text & "'s info site</title>" & vbCrLf & "Name: " & Text1.Text & vbCrLf & "<br>E-mail:<a href = ""mailto:" & Text2.Text & """>" & Text2.Text & "</a><br>Gender: " & Combo1.Text & vbCrLf & "<br>Birthdate: " & Text3.Text & vbCrLf & "<br>Age: " & Text4.Text & vbCrLf & "<br><center><b>Some info about me</b></center><br><center>" & vbCrLf & Text5.Text & "</center><br><img src = """ & Text6.Text & """>"
If Check1.Value = 0 Then Form1.Text1.Text = "<title>" & Text1.Text & "'s info site</title>" & vbCrLf & "Name: " & Text1.Text & vbCrLf & "<br>E-mail:<a href = ""mailto:" & Text2.Text & """>" & Text2.Text & "</a><br>Gender: " & Combo1.Text & vbCrLf & "<br>Birthdate: " & Text3.Text & vbCrLf & "<br>Age: " & Text4.Text & vbCrLf & "<br><center><b>Some info about me</b></center><br><center>" & vbCrLf & Text5.Text & "</center>"
Form1.Show
Me.Hide
End Sub

Private Sub Form_Load()
Combo1.Text = "Male"
WB1.Navigate "about:People will not see the picture if it comes from your harddrive, it must be from a website. Example: http://microsoft.com/homepage/gif/bnr-microsoft.gif"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form1.Show
Me.Hide
End Sub
