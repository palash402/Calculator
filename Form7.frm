VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H0000FFFF&
   Caption         =   "Square"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14610
   LinkTopic       =   "Form7"
   ScaleHeight     =   9150
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H0000FFFF&
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
      TabIndex        =   4
      Top             =   3480
      Width           =   4935
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
      Left            =   6960
      TabIndex        =   3
      Top             =   6840
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
      Height          =   855
      Left            =   4080
      TabIndex        =   2
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0000FFFF&
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
      TabIndex        =   1
      Top             =   2520
      Width           =   9495
   End
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
      Left            =   8280
      TabIndex        =   0
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Area Is:---------------"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter Side Of Square-----"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Area Of Square"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.SetFocus
Dim b As Integer
a = Text2.Text
Text3.Text = a * a
Text2.SetFocus

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
calculator.Show
End Sub
