VERSION 5.00
Begin VB.Form Form17 
   BackColor       =   &H0000FF00&
   Caption         =   "Square"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form17"
   ScaleHeight     =   5355
   ScaleWidth      =   9315
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
      Left            =   2640
      TabIndex        =   7
      Top             =   4320
      Width           =   3975
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
      Height          =   735
      Left            =   5760
      TabIndex        =   6
      Top             =   3360
      Width           =   2775
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
      Height          =   735
      Left            =   720
      TabIndex        =   5
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0000FF00&
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
      Left            =   3600
      TabIndex        =   4
      Top             =   2400
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
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
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Square is------------------"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Enter the Number---"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Square"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
Dim a As Integer
a = Text1.Text
Text2.Text = a * a
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
calculator.Show
End Sub
