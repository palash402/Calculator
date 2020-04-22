VERSION 5.00
Begin VB.Form Form26 
   BackColor       =   &H0000FF00&
   Caption         =   "Log"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form26"
   ScaleHeight     =   8595
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
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
      Left            =   1680
      TabIndex        =   11
      Top             =   6840
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
      Height          =   855
      Left            =   6720
      TabIndex        =   10
      Top             =   6840
      Width           =   2775
   End
   Begin VB.TextBox txtNumber 
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
      Height          =   525
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtBase 
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
      Height          =   525
      Left            =   6000
      TabIndex        =   3
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton cmdFindLog 
      Caption         =   "Find Log"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      Top             =   3240
      Width           =   3855
   End
   Begin VB.TextBox txtResult 
      Alignment       =   1  'Right Justify
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
      Height          =   525
      Left            =   3840
      TabIndex        =   1
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox txtVerify 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3840
      TabIndex        =   0
      Top             =   5760
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Log"
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
      Index           =   4
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Number:"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Base:"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   2
      Left            =   4560
      TabIndex        =   7
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Result:--------"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   4680
      Width           =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Verify:---"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   3
      Left            =   2400
      TabIndex        =   5
      Top             =   5760
      Width           =   1500
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFindLog_Click()
Dim num, base, result As Double

    num = Val(txtNumber.Text)
    base = Val(txtBase.Text)

    ' Calculate Log(num) in the base.
    result = Log(num) / Log(base)

    txtResult.Text = Format$(result)
    txtVerify.Text = Format$(base ^ result)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click()
calculator.Show
End Sub
