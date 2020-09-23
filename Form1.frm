VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weather"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2130
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   2130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   270
      Left            =   1890
      TabIndex        =   2
      Top             =   210
      Width           =   210
   End
   Begin Project1.Weather Weather1 
      Left            =   1725
      Top             =   1590
      _ExtentX        =   1508
      _ExtentY        =   1164
      Temp            =   ""
      Condition       =   ""
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   195
      Width           =   1830
   End
   Begin VB.Label lblHum 
      Caption         =   "Humidity:"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label lblFeel 
      BackStyle       =   0  'Transparent
      Caption         =   "Feels like 30* F"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   495
      TabIndex        =   5
      Top             =   900
      Width           =   1485
   End
   Begin VB.Label lblCondition 
      Caption         =   "Mostly Sunny"
      Height          =   240
      Left            =   420
      TabIndex        =   4
      Top             =   1140
      Width           =   1635
   End
   Begin VB.Label lblTemp 
      Caption         =   "40* F"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   390
      TabIndex        =   3
      Top             =   525
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "zip code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   1
      Top             =   30
      Width           =   1425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo poop
Weather1.ZIP = Text1
Weather1.UpdateWeather
lblTemp = Weather1.Temp
lblFeel = "Feels like " & Weather1.FeelsLike
lblHum = "Humidity: " & Weather1.Humidity
lblCondition = Weather1.Condition
Me.Height = 2295
poop:

End Sub

Private Sub Form_Load()
Me.Height = 945
End Sub

