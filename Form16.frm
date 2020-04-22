VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H0000FF00&
   Caption         =   "Factorial"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Old English Text MT"
      Size            =   36
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form16"
   ScaleHeight     =   6075
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
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
      Left            =   6960
      TabIndex        =   7
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Main Menu"
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
      Left            =   3720
      TabIndex        =   6
      Top             =   3600
      Width           =   2655
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
      Left            =   360
      TabIndex        =   5
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text2 
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
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   7215
   End
   Begin VB.TextBox Text1 
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   7215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "n!"
      Height          =   735
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "n--"
      Height          =   735
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Factorial(n!)"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x, fact As Long
x = Text1.Text
fact = 1
For i = x To 1 Step -1
fact = fact * i
Next i
Text2.Text = fact
End Sub

Private Sub Command2_Click()
calculator.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
