VERSION 5.00
Begin VB.Form menu 
   BackColor       =   &H0000FF00&
   Caption         =   "Menu"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14580
   LinkTopic       =   "Form30"
   ScaleHeight     =   9120
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      TabIndex        =   3
      Top             =   6240
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go To Calculator"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Subject-ICT"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Regional Science Exibition"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   13815
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
calculator.Show
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
