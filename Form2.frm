VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF0000&
   Caption         =   "Substract"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14580
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   9675
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Go To Main Menu"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      TabIndex        =   9
      Top             =   5160
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   1560
      Width           =   9855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   2520
      Width           =   9495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   2
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   1
      Top             =   6720
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   3480
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Substraction Menu"
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
      TabIndex        =   8
      Top             =   0
      Width           =   14175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter First Number"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter Second Number"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "Difference is:-"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
Dim a As Double
Dim b As Double
Text1.SetFocus
a = Text1.Text
b = Text2.Text
Text3.Text = a - b
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
calculator.Show
End Sub
