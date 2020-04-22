VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF0000&
   Caption         =   "Area Menu"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14610
   LinkTopic       =   "Form5"
   ScaleHeight     =   9660
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
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
      Height          =   975
      Left            =   5760
      TabIndex        =   8
      Top             =   6960
      Width           =   3135
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Go to Main Menu"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   7
      Top             =   5520
      Width           =   4815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Rhombus"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10680
      TabIndex        =   6
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Parallalogram"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      TabIndex        =   5
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Tripazium"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rectangle"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   10680
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Square"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Triangle"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Area Calculation Menu"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   14295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Show
End Sub

Private Sub Command2_Click()
Form7.Show
End Sub

Private Sub Command3_Click(Index As Integer)
Form8.Show
End Sub

Private Sub Command4_Click()
Form9.Show
End Sub

Private Sub Command5_Click()
Form10.Show
End Sub

Private Sub Command6_Click()
Form11.Show
End Sub

Private Sub Command7_Click()
calculator.Show
End Sub

Private Sub Command8_Click()
Unload Me
End Sub
