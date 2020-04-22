VERSION 5.00
Begin VB.Form Form23 
   BackColor       =   &H00FFFF00&
   Caption         =   "Sphere"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14610
   LinkTopic       =   "Form23"
   ScaleHeight     =   9345
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFF00&
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
      Left            =   5880
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
      Left            =   6360
      TabIndex        =   3
      Top             =   6600
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
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
      Left            =   3600
      TabIndex        =   2
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF00&
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
      Left            =   4680
      TabIndex        =   1
      Top             =   2520
      Width           =   9495
   End
   Begin VB.CommandButton Command4 
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
      Left            =   7680
      TabIndex        =   0
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Volume of Sphere"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Volume Is-------------"
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
      Left            =   2640
      TabIndex        =   6
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter Radius of Sphere---"
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
      Left            =   960
      TabIndex        =   5
      Top             =   2520
      Width           =   5055
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim r As Double
r = Text2.Text
Text3.Text = (4 / 3) * 3.1415 * (r ^ 3)
End Sub

Private Sub Command4_Click()
calculator.Show
End Sub
