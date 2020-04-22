VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H000080FF&
   Caption         =   "Tripazium"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14535
   LinkTopic       =   "Form9"
   ScaleHeight     =   9375
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   11
      Top             =   7320
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go to Main Menu"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      TabIndex        =   10
      Top             =   6000
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   9
      Top             =   6000
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   4680
      Width           =   6855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   3720
      Width           =   6855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   2520
      Width           =   6855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   1560
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "Area of Tripazium is---------"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   7
      Top             =   4680
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Area of Tripazium"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14415
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "Enter Value of First Parallal Side-----"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "Enter Value of Second Parallal Side---"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "Enter Height of Tripazium-------"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   0
      Top             =   3720
      Width           =   5655
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
Dim a As Integer
Dim b As Integer
a = Text1.Text
b = Text2.Text
c = Text3.Text
Text4.Text = (1 / 2) * (a + b) * c
Text1.SetFocus
End Sub

Private Sub Command2_Click()
calculator.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
