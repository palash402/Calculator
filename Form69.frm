VERSION 5.00
Begin VB.Form Form40 
   BackColor       =   &H00FFFF00&
   Caption         =   "mileage_consumption"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   14340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   28
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   27
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Metric"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   2
      Left            =   10800
      TabIndex        =   22
      Top             =   3120
      Width           =   3255
      Begin VB.TextBox txtKpl 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtLitersPer100km 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         Caption         =   "km/liter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   960
         TabIndex        =   26
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label txtMpg 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         Caption         =   "Liters/100 km"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "British"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   1
      Left            =   6240
      TabIndex        =   17
      Top             =   3120
      Width           =   3135
      Begin VB.TextBox txtUkMpg 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtUkGalPer100Miles 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label txtMpg 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         Caption         =   "MPG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         Caption         =   "Gal/100 miles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "American"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   0
      Left            =   960
      TabIndex        =   12
      Top             =   3120
      Width           =   3255
      Begin VB.TextBox txtUsGalPer100Miles 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtUsMpg 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         Caption         =   "Gal/100 miles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label txtMpg 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         Caption         =   "MPG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   11
      Top             =   2160
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5760
      ScaleHeight     =   255
      ScaleWidth      =   4095
      TabIndex        =   7
      Top             =   1440
      Width           =   4095
      Begin VB.OptionButton optLiters 
         BackColor       =   &H00FFFF00&
         Caption         =   "Liters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optUsGallons 
         BackColor       =   &H00FFFF00&
         Caption         =   "US Gallons"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optUkGallons 
         BackColor       =   &H00FFFF00&
         Caption         =   "UK Gallons"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5760
      ScaleHeight     =   255
      ScaleWidth      =   2895
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
      Begin VB.OptionButton optKilometers 
         BackColor       =   &H00FFFF00&
         Caption         =   "Kilometers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
      Begin VB.OptionButton optMiles 
         BackColor       =   &H00FFFF00&
         Caption         =   "Miles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtFuel 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtDistance 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3480
      TabIndex        =   30
      Top             =   720
      Width           =   7455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Milage Consumpsion"
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
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   14415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Fuel:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Distance:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   1320
   End
End
Attribute VB_Name = "Form40"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Conversion factors.
Private Const miles_per_km As Single = 0.621371192
Private Const us_gallons_per_liter As Single = 0.264172052637296
Private Const uk_gallons_per_liter As Single = 0.2199692483

' Calculate the mileage statistics.
Private Sub cmdCalculate_Click()
Dim miles As Single
Dim gallons As Single
Dim us_mpg As Single
Dim us_gal_per_100_miles As Single
Dim uk_gals As Single
Dim uk_mpg As Single
Dim uk_gal_per_100_miles
Dim liters As Single
Dim kilometers As Single
Dim km_per_liter As Single
Dim l_per_100_km As Single

    ' Get the inputs.
    miles = Val(txtDistance.Text)
    gallons = Val(txtFuel.Text)

    ' Convert to miles and US gallons.
    If (optKilometers.Value) Then miles = miles * miles_per_km
    If (optUkGallons.Value) Then
        gallons = gallons * us_gallons_per_liter / uk_gallons_per_liter
    ElseIf (optKilometers.Value) Then
        gallons = gallons * us_gallons_per_liter
    End If

    ' Calculate performance values.
    us_mpg = miles / gallons
    us_gal_per_100_miles = 100 / us_mpg

    ' US values.
    txtUsMpg.Text = Format$(us_mpg, "0.00")
    txtUsGalPer100Miles.Text = Format$(us_gal_per_100_miles, "0.00")

    ' UK values.
    uk_gals = gallons * uk_gallons_per_liter / us_gallons_per_liter
    uk_mpg = miles / uk_gals
    uk_gal_per_100_miles = 100 / uk_mpg
    txtUkMpg.Text = Format$(uk_mpg, "0.00")
    txtUkGalPer100Miles.Text = Format$(uk_gal_per_100_miles, "0.00")

    ' Metric values.
    liters = gallons / us_gallons_per_liter
    kilometers = miles / miles_per_km
    km_per_liter = kilometers / liters
    l_per_100_km = 100 / km_per_liter
    txtKpl.Text = Format$(km_per_liter, "0.00")
    txtLitersPer100km.Text = Format$(l_per_100_km, "0.00")
End Sub

Private Sub Command1_Click()
calculator.Show

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
