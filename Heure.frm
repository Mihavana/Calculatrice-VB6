VERSION 5.00
Begin VB.Form Heure 
   BackColor       =   &H00202020&
   Caption         =   "Calculatrice"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameSupport 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   -8000
      TabIndex        =   17
      Top             =   600
      Width           =   2895
      Begin VB.Frame FrameMenu 
         BackColor       =   &H00252525&
         BorderStyle     =   0  'None
         Caption         =   "cc"
         Height          =   10935
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   2655
         Begin VB.Image Logos 
            Height          =   375
            Index           =   15
            Left            =   0
            Picture         =   "Heure.frx":0000
            Stretch         =   -1  'True
            Top             =   10320
            Width           =   375
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   14
            Left            =   0
            Picture         =   "Heure.frx":03D6
            Top             =   9600
            Width           =   480
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   13
            Left            =   0
            Picture         =   "Heure.frx":07C6
            Top             =   9000
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   420
            Index           =   12
            Left            =   0
            Picture         =   "Heure.frx":0B91
            Top             =   7680
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   420
            Index           =   11
            Left            =   0
            Picture         =   "Heure.frx":0FC5
            Top             =   7080
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   570
            Index           =   10
            Left            =   0
            Picture         =   "Heure.frx":139D
            Top             =   8280
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   9
            Left            =   0
            Picture         =   "Heure.frx":1788
            Top             =   6480
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   8
            Left            =   0
            Picture         =   "Heure.frx":1B2B
            Top             =   5880
            Width           =   480
         End
         Begin VB.Image Logos 
            Height          =   375
            Index           =   7
            Left            =   80
            Picture         =   "Heure.frx":1F11
            Top             =   4800
            Width           =   330
         End
         Begin VB.Image Logos 
            Height          =   495
            Index           =   6
            Left            =   0
            Picture         =   "Heure.frx":22FF
            Stretch         =   -1  'True
            Top             =   5280
            Width           =   450
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   5
            Left            =   120
            Picture         =   "Heure.frx":26A2
            Top             =   4080
            Width           =   300
         End
         Begin VB.Image Logos 
            Height          =   525
            Index           =   4
            Left            =   50
            Picture         =   "Heure.frx":2A5C
            Top             =   3480
            Width           =   405
         End
         Begin VB.Label surf 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Surface"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   9
            Left            =   600
            TabIndex        =   37
            Top             =   6600
            Width           =   1335
         End
         Begin VB.Label energ 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Energie"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   8
            Left            =   600
            TabIndex        =   36
            Top             =   6000
            Width           =   975
         End
         Begin VB.Label température 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Température"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   7
            Left            =   600
            TabIndex        =   35
            Top             =   5400
            Width           =   2055
         End
         Begin VB.Label PoidsEtMasse 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Poids et Masse"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Index           =   6
            Left            =   600
            TabIndex        =   34
            Top             =   4800
            Width           =   2055
         End
         Begin VB.Label Longueur 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Longueur"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   5
            Left            =   600
            TabIndex        =   33
            Top             =   4200
            Width           =   2055
         End
         Begin VB.Label volume 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Volume"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   4
            Left            =   600
            TabIndex        =   32
            Top             =   3600
            Width           =   2055
         End
         Begin VB.Label Listes 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Convertisseur"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Index           =   17
            Left            =   0
            TabIndex        =   31
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label Listes 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Calculatrice"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   30
            Top             =   120
            Width           =   1935
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   3
            Left            =   120
            Picture         =   "Heure.frx":2EBF
            Top             =   2280
            Width           =   420
         End
         Begin VB.Label CalculDedate 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Calcul de la date"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   600
            TabIndex        =   29
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   0
            Left            =   120
            Picture         =   "Heure.frx":32E8
            Top             =   480
            Width           =   390
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   1
            Left            =   120
            Picture         =   "Heure.frx":372B
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   2
            Left            =   0
            Picture         =   "Heure.frx":3AF7
            Top             =   1680
            Width           =   555
         End
         Begin VB.Label Graphique 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Graphique"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   600
            TabIndex        =   28
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label Scientifique 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Scientifique"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   600
            TabIndex        =   27
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Standard 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Standard"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   600
            TabIndex        =   26
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label vit 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Vitesse"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   10
            Left            =   600
            TabIndex        =   25
            Top             =   7200
            Width           =   1095
         End
         Begin VB.Label heur 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Heure"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   11
            Left            =   600
            TabIndex        =   24
            Top             =   7800
            Width           =   1095
         End
         Begin VB.Label puissa 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Puissance"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   12
            Left            =   600
            TabIndex        =   23
            Top             =   8400
            Width           =   1455
         End
         Begin VB.Label don 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Données"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   600
            TabIndex        =   22
            Top             =   9000
            Width           =   1335
         End
         Begin VB.Label press 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Pression"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   14
            Left            =   600
            TabIndex        =   21
            Top             =   9720
            Width           =   1095
         End
         Begin VB.Label Ang 
            BackColor       =   &H00252525&
            BackStyle       =   0  'Transparent
            Caption         =   "Angle"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   15
            Left            =   600
            TabIndex        =   20
            Top             =   10320
            Width           =   975
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   7215
         LargeChange     =   300
         Left            =   2640
         Max             =   4100
         SmallChange     =   300
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.ComboBox listVol 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      ItemData        =   "Heure.frx":3F07
      Left            =   120
      List            =   "Heure.frx":3F1A
      TabIndex        =   14
      Top             =   6000
      Width           =   3135
   End
   Begin VB.ComboBox listVol 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      ItemData        =   "Heure.frx":3F57
      Left            =   120
      List            =   "Heure.frx":3F6A
      TabIndex        =   13
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton Virgule 
      BackColor       =   &H00404040&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton CLR 
      BackColor       =   &H00404040&
      Caption         =   "CE"
      Height          =   1095
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Delete 
      BackColor       =   &H00404040&
      Caption         =   "Õ"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton chiffre 
      BackColor       =   &H00404040&
      Caption         =   "9"
      Height          =   1095
      Index           =   9
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton chiffre 
      BackColor       =   &H00404040&
      Caption         =   "8"
      Height          =   1095
      Index           =   8
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton chiffre 
      BackColor       =   &H00404040&
      Caption         =   "7"
      Height          =   1095
      Index           =   7
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton chiffre 
      BackColor       =   &H00404040&
      Caption         =   "6"
      Height          =   1095
      Index           =   6
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton chiffre 
      BackColor       =   &H00404040&
      Caption         =   "5"
      Height          =   1095
      Index           =   5
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton chiffre 
      BackColor       =   &H00404040&
      Caption         =   "4"
      Height          =   1095
      Index           =   4
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton chiffre 
      BackColor       =   &H00404040&
      Caption         =   "3"
      Height          =   1095
      Index           =   3
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton chiffre 
      BackColor       =   &H00404040&
      Caption         =   "2"
      Height          =   1095
      Index           =   2
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton chiffre 
      BackColor       =   &H00404040&
      Caption         =   "1"
      Height          =   1095
      Index           =   1
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton chiffre 
      BackColor       =   &H00404040&
      Caption         =   "0"
      Height          =   1095
      Index           =   0
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00202020&
      Caption         =   "Heure"
      BeginProperty Font 
         Name            =   "Sans Serif Collection"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   960
      TabIndex        =   39
      Top             =   0
      Width           =   990
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   600
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   600
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      BackColor       =   &H00202020&
      Height          =   495
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   615
   End
   Begin VB.Label panConversion 
      BackColor       =   &H00202020&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Label panConversion 
      BackColor       =   &H00202020&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   3495
   End
End
Attribute VB_Name = "Heure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numTabPan
Dim tabPanneau(10000)
Dim j, l
Dim vir

Private Sub Command1_Click(Index As Integer)
End Sub

Private Sub Ang_Click(Index As Integer)
Unload Me
Angle.Show
End Sub

Private Sub CalculDedate_Click(Index As Integer)
Unload Me
CalculdelaDate.Show
End Sub

Private Sub chiffre_Click(i As Integer)
numTabPan = numTabPan + 1
tabPanneau(numTabPan) = i
If panConversion(j).Caption <> "0" Then panConversion(j).Caption = panConversion(j).Caption & i
If panConversion(j).Caption = 0 Then panConversion(j).Caption = i
Call conversion
End Sub

Private Sub Command2_Click()

End Sub

Private Sub CLR_Click()
panConversion(0) = 0
panConversion(1) = 0
vir = True
End Sub

Private Sub Command4_Click()

End Sub

Function conversion()
If j = 0 Then
l = 1
ElseIf j = 1 Then
l = 0
End If
If listVol(j) = "Microsecondes" Then
   If listVol(l) = "Microsecondes" Then
    panConversion(l) = panConversion(j)
   ElseIf listVol(l) = "Millisecondes" Then
    panConversion(l) = panConversion(j) * 0.001
   ElseIf listVol(l) = "Secondes" Then
    panConversion(l) = panConversion(j) * 0.000001
   ElseIf listVol(l) = "Minutes" Then
    panConversion(l) = panConversion(j) * 0.000000016666667
   ElseIf listVol(l) = "Heures" Then
    panConversion(l) = panConversion(j) * 0.000000000277778
   End If
ElseIf listVol(j) = "Millisecondes" Then
   If listVol(l) = "Millisecondes" Then
    panConversion(l) = panConversion(j)
   ElseIf listVol(l) = "Microsecondes" Then
    panConversion(l) = panConversion(j) * 1000
   ElseIf listVol(l) = "Secondes" Then
    panConversion(l) = panConversion(j) * 0.001
   ElseIf listVol(l) = "Minutes" Then
    panConversion(l) = panConversion(j) * 0.000017
   ElseIf listVol(l) = "Heures" Then
    panConversion(l) = panConversion(j) * 0.000000277777778
   End If
ElseIf listVol(j) = "Secondes" Then
   If listVol(l) = "Secondes" Then
    panConversion(l) = panConversion(j)
   ElseIf listVol(l) = "Microsecondes" Then
    panConversion(l) = panConversion(j) * 1000000
   ElseIf listVol(l) = "Millisecondes" Then
    panConversion(l) = panConversion(j) * 1000
   ElseIf listVol(l) = "Minutes" Then
    panConversion(l) = panConversion(j) * 0.016667
   ElseIf listVol(l) = "Heures" Then
    panConversion(l) = panConversion(j) * 0.000278
   End If
ElseIf listVol(j) = "Minutes" Then
   If listVol(l) = "Minutes" Then
    panConversion(l) = panConversion(j)
   ElseIf listVol(l) = "Microsecondes" Then
    panConversion(l) = panConversion(j) * 60000000
   ElseIf listVol(l) = "Millisecondes" Then
    panConversion(l) = panConversion(j) * 60000
   ElseIf listVol(l) = "Secondes" Then
    panConversion(l) = panConversion(j) * 60
   ElseIf listVol(l) = "Heures" Then
    panConversion(l) = panConversion(j) * 0.016667
   End If
ElseIf listVol(j) = "Heures" Then
   If listVol(l) = "Heures" Then
    panConversion(l) = panConversion(j)
   ElseIf listVol(l) = "Microsecondes" Then
    panConversion(l) = panConversion(j) * 3600000000#
   ElseIf listVol(l) = "Millisecondes" Then
    panConversion(l) = panConversion(j) * 3600000
   ElseIf listVol(l) = "Secondes" Then
    panConversion(l) = panConversion(j) * 3600
   ElseIf listVol(l) = "Minutes" Then
    panConversion(l) = panConversion(j) * 60
   End If
End If
End Function

Private Sub Delete_Click()
panConversion(j) = Left(panConversion(j), Len(panConversion(j)) - 1)
If Len(panConversion(j)) = 0 Then panConversion(j) = 0
If InStr(1, panConversion(j), ",") = 0 Then vir = True
Call conversion
End Sub

Private Sub don_Click(Index As Integer)
Unload Me
Données.Show
End Sub

Private Sub energ_Click(Index As Integer)
Unload Me
Energie.Show
End Sub

Private Sub Form_Load()
j = 0
vir = True
End Sub
Private Sub Form_Click()
If FrameSupport.Left < 13000 Then
Do While FrameSupport.Left < 13000
FrameSupport.Left = FrameSupport.Left + 10
Loop
End If
End Sub
Private Sub Graphique_Click(Index As Integer)
Unload Me
Graph.Show
End Sub

Private Sub heur_Click(Index As Integer)
Unload Me
Heure.Show
End Sub

Private Sub Label2_Click()
If FrameSupport.Left < 0 Then
Do While FrameSupport.Left < 0
FrameSupport.Left = FrameSupport.Left + 10
Loop
Else
Do While FrameSupport.Left > -8000
FrameSupport.Left = FrameSupport.Left - 10
Loop
End If
End Sub



Private Sub Longueur_Click(Index As Integer)
Unload Me
Longue.Show
End Sub

Private Sub panConversion_Click(k As Integer)
j = k
End Sub

Private Sub PoidsEtMasse_Click(Index As Integer)
Unload Me
PoidsMasse.Show
End Sub

Private Sub press_Click(Index As Integer)
Unload Me
Pression.Show
End Sub

Private Sub puissa_Click(Index As Integer)
Unload Me
puissance.Show
End Sub

Private Sub Scientifique_Click(Index As Integer)
Unload Me
Scient.Show
End Sub

Private Sub Standard_Click(Index As Integer)
Unload Me
Stand.Show
End Sub

Private Sub surf_Click(Index As Integer)
Unload Me
Surface.Show
End Sub

Private Sub température_Click(Index As Integer)
Unload Me
Temp.Show
End Sub

Private Sub Virgule_Click()
If vir = True Then panConversion(j) = panConversion(j) & ","
vir = False
End Sub


Private Sub vit_Click(Index As Integer)
Unload Me
Vitesse.Show
End Sub

Private Sub volume_Click(Index As Integer)
Unload Me
Vol.Show
End Sub

Private Sub VScroll1_Change()
FrameMenu.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
Call VScroll1_Change
End Sub
