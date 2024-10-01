VERSION 5.00
Begin VB.Form CalculdelaDate 
   BackColor       =   &H00202020&
   Caption         =   "Calculatrice"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   11170.36
   ScaleMode       =   0  'User
   ScaleWidth      =   25515.48
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameSupport 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   -2836
      TabIndex        =   33
      Top             =   840
      Width           =   2775
      Begin VB.Frame FrameMenu 
         BackColor       =   &H00252525&
         BorderStyle     =   0  'None
         Caption         =   "cc"
         Height          =   10935
         Left            =   0
         TabIndex        =   35
         Top             =   -360
         Width           =   2535
         Begin VB.Image Logos 
            Height          =   375
            Index           =   15
            Left            =   0
            Picture         =   "Calcul de la date.frx":0000
            Stretch         =   -1  'True
            Top             =   10320
            Width           =   375
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   14
            Left            =   0
            Picture         =   "Calcul de la date.frx":03D6
            Top             =   9600
            Width           =   480
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   13
            Left            =   0
            Picture         =   "Calcul de la date.frx":07C6
            Top             =   9000
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   420
            Index           =   12
            Left            =   0
            Picture         =   "Calcul de la date.frx":0B91
            Top             =   7680
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   420
            Index           =   11
            Left            =   0
            Picture         =   "Calcul de la date.frx":0FC5
            Top             =   7080
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   570
            Index           =   10
            Left            =   0
            Picture         =   "Calcul de la date.frx":139D
            Top             =   8280
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   9
            Left            =   0
            Picture         =   "Calcul de la date.frx":1788
            Top             =   6480
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   8
            Left            =   0
            Picture         =   "Calcul de la date.frx":1B2B
            Top             =   5880
            Width           =   480
         End
         Begin VB.Image Logos 
            Height          =   375
            Index           =   7
            Left            =   80
            Picture         =   "Calcul de la date.frx":1F11
            Top             =   4800
            Width           =   330
         End
         Begin VB.Image Logos 
            Height          =   495
            Index           =   6
            Left            =   0
            Picture         =   "Calcul de la date.frx":22FF
            Stretch         =   -1  'True
            Top             =   5280
            Width           =   450
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   5
            Left            =   120
            Picture         =   "Calcul de la date.frx":26A2
            Top             =   4080
            Width           =   300
         End
         Begin VB.Image Logos 
            Height          =   525
            Index           =   4
            Left            =   50
            Picture         =   "Calcul de la date.frx":2A5C
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            Left            =   120
            TabIndex        =   47
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
            TabIndex        =   46
            Top             =   120
            Width           =   1935
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   3
            Left            =   120
            Picture         =   "Calcul de la date.frx":2EBF
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
            TabIndex        =   45
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   0
            Left            =   120
            Picture         =   "Calcul de la date.frx":32E8
            Top             =   480
            Width           =   390
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   1
            Left            =   120
            Picture         =   "Calcul de la date.frx":372B
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   2
            Left            =   0
            Picture         =   "Calcul de la date.frx":3AF7
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
            Top             =   10320
            Width           =   975
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   7215
         LargeChange     =   300
         Left            =   2520
         Max             =   3600
         SmallChange     =   300
         TabIndex        =   34
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00252525&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label choix1 
         Alignment       =   2  'Center
         BackColor       =   &H00252525&
         Caption         =   "Difference entre les dates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label choix2 
         Alignment       =   2  'Center
         BackColor       =   &H00252525&
         Caption         =   "Ajouter ou soustraire des jours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   480
      ScaleHeight     =   5775
      ScaleWidth      =   5415
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton Calcul2 
         BackColor       =   &H00252525&
         Caption         =   "Calcul"
         Height          =   735
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton Soustraire 
         BackColor       =   &H00202020&
         Caption         =   "Soustraire"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Réponse 
         BackColor       =   &H00202020&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   4440
         Width           =   4455
      End
      Begin VB.ComboBox jour 
         BackColor       =   &H00252525&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3240
         TabIndex        =   17
         Text            =   "0"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox mois 
         BackColor       =   &H00252525&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         TabIndex        =   16
         Text            =   "0"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox Année 
         BackColor       =   &H00252525&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   360
         TabIndex        =   15
         Text            =   "0"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton Ajouter 
         BackColor       =   &H00202020&
         Caption         =   "Ajouter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Date2 
         BackColor       =   &H00202020&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   600
         TabIndex        =   13
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00202020&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00202020&
         Caption         =   "Jours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3240
         TabIndex        =   20
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00202020&
         Caption         =   "Mois"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00202020&
         Caption         =   "Années"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00202020&
         Caption         =   "De"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.ComboBox annéeChoix 
      BackColor       =   &H00252525&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   3120
      TabIndex        =   32
      Text            =   "Année"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox moisChoix 
      BackColor       =   &H00252525&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   31
      Text            =   "Mois"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox jourChoix 
      BackColor       =   &H00252525&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   30
      Text            =   "Jour"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox annéeChoix 
      BackColor       =   &H00252525&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   29
      Text            =   "Année"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ComboBox moisChoix 
      BackColor       =   &H00252525&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   28
      Text            =   "Mois"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ComboBox jourChoix 
      BackColor       =   &H00252525&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   27
      Text            =   "Jour"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Calcul 
      BackColor       =   &H00252525&
      Caption         =   "Calcul"
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Resultat 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   720
      TabIndex        =   6
      Top             =   5520
      Width           =   5055
   End
   Begin VB.TextBox dateFin 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   600
      TabIndex        =   5
      Top             =   4080
      Width           =   3735
   End
   Begin VB.TextBox dateDebut 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   12525.01
      X2              =   12863.53
      Y1              =   1624.78
      Y2              =   1444.249
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   12186.5
      X2              =   12525.01
      Y1              =   1444.249
      Y2              =   1624.78
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   677.028
      X2              =   1692.569
      Y1              =   722.124
      Y2              =   722.124
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   677.028
      X2              =   1692.569
      Y1              =   361.062
      Y2              =   361.062
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   677.028
      X2              =   1692.569
      Y1              =   541.593
      Y2              =   541.593
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00202020&
      Caption         =   "Calcul de la Date"
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
      Left            =   840
      TabIndex        =   26
      Top             =   120
      Width           =   2640
   End
   Begin VB.Label Choix 
      BackColor       =   &H00202020&
      BackStyle       =   0  'Transparent
      Caption         =   "Difference entre les dates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label difference 
      BackColor       =   &H00202020&
      Caption         =   "Difference"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label A 
      BackColor       =   &H00202020&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00202020&
      Caption         =   "De"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00202020&
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "CalculdelaDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stockAnnée, stockMois, stockJour
Dim i
Dim ajout, soustrait
Dim vrai
Dim debut, fin

Private Sub Text3_Change()

End Sub


Private Sub Ajouter_Click()
ajout = True
soustrait = False
End Sub

Private Sub Ang_Click(Index As Integer)
Unload Me
Angle.Show
End Sub

Private Sub Calcul_Click()
If jourChoix(0) <> "Jour" And moisChoix(0) <> "Mois" And annéeChoix(0) <> "Année" Then
dateDebut = jourChoix(0) & "/" & moisChoix(0) & "/" & annéeChoix(0)
End If
If jourChoix(1) <> "Jour" And moisChoix(1) <> "Mois" And annéeChoix(1) <> "Année" Then
dateFin = jourChoix(1) & "/" & moisChoix(1) & "/" & annéeChoix(1)
End If
debut = CDate(dateDebut)
fin = CDate(dateFin)

If debut = fin Then
Resultat = "Dates identiques"
Else
stockJour = DateDiff("d", debut, fin)
    If stockJour > 30 Then
        Do While stockJour > 30
            stockJour = stockJour - 31
            stockMois = stockMois + 1
            If vrai = True Then stockJour = stockJour + 1
            vrai = Not vrai
        Loop
    End If
    If stockMois > 11 Then
        Do While stockMois > 11
            stockMois = stockMois - 12
            stockAnnée = stockAnnée + 1
        Loop
    End If
    If stockJour <> 0 Then
        If stockMois <> 0 Then
            If stockAnnée <> 0 Then
                Resultat = stockAnnée & " Année(s); " & stockMois & " Mois; " & stockJour & " Jour(s)"
            Else
                Resultat = stockMois & " Mois; " & stockJour & " Jour(s)"
            End If
        ElseIf stockMois = 0 Then
            If stockAnnée <> 0 Then
                Resultat = stockAnnée & " Année(s); " & stockJour & " Jour(s)"
            Else
                Resultat = stockJour & " Jour(s)"
            End If
        End If
    ElseIf stockJour = 0 Then
        If stockMois <> 0 Then
            If stockAnnée <> 0 Then
                Resultat = stockAnnée & " Année(s); " & stockMois & " Moiss "
            Else
                Resultat = stockMois & " Mois; " & stockJour & ""
            End If
        ElseIf stockMois = 0 Then
            If stockAnnée <> 0 Then
                Resultat = stockAnnée & " Année(s) "
            End If
        End If
    End If
End If
stockJour = 0
stockMois = 0
stocckannée = 0
End Sub

Private Sub Calcul2_Click()
If ajout = True Then
    If Année <> 0 Then
        Réponse = DateAdd("yyyy", Année, Date2)
        stockAnnée = DateAdd("yyyy", Année, Date2)
    End If
    If mois <> 0 Then
        If Année <> 0 Then
            Réponse = DateAdd("m", mois, stockAnnée)
            stockMois = DateAdd("m", mois, stockAnnée)
        Else
            Réponse = DateAdd("m", mois, Date2)
            stockMois = DateAdd("m", mois, Date2)
        End If
    End If
    If jour <> 0 Then
        If mois <> 0 Then
            Réponse = DateAdd("d", jour, stockMois)
        ElseIf Année <> 0 And mois = 0 Then
            Réponse = DateAdd("d", jour, stockAnnée)
        Else
            Réponse = DateAdd("d", jour, Date2)
        End If
    End If
End If
If soustrait = True Then
    If Année <> 0 Then
        Réponse = DateAdd("yyyy", -Année, Date2)
        stockAnnée = DateAdd("yyyy", -Année, Date2)
    End If
    If mois <> 0 Then
        If Année <> 0 Then
            Réponse = DateAdd("m", -mois, stockAnnée)
            stockMois = DateAdd("m", -mois, stockAnnée)
        Else
            Réponse = DateAdd("m", -mois, Date2)
            stockMois = DateAdd("m", -mois, Date2)
        End If
    End If
    If jour <> 0 Then
        If mois <> 0 Then
            Réponse = DateAdd("d", -jour, stockMois)
        ElseIf Année <> 0 And mois = 0 Then
            Réponse = DateAdd("d", -jour, stockAnnée)
        Else
            Réponse = DateAdd("d", -jour, Date2)
        End If
    End If
End If
stockMois = 0
stockAnnée = 0
End Sub

Private Sub calculDeLDate_Click()
Stand.Hide
Scient.Hide
Graph.Hide
calculDate.Show
End Sub

Private Sub CalculDedate_Click(Index As Integer)
Unload Me
CalculdelaDate.Show
End Sub

Private Sub Choix_Click()
Frame1.Visible = True
End Sub

Private Sub choix1_Click()
Choix.Caption = "Difference entre les dates"
Frame1.Visible = False
Picture1.Visible = False
End Sub

Private Sub choix2_Click()
Choix.Caption = "Ajouter ou soustraire des jours"
Frame1.Visible = False
Picture1.Visible = True
End Sub

Private Sub don_Click(Index As Integer)
Unload Me
Données.Show
End Sub

Private Sub energ_Click(Index As Integer)
Unload Me
Energie.Show
End Sub
Private Sub Form_Click()
If FrameSupport.Left > -8000 Then
Do While FrameSupport.Left > -8000
FrameSupport.Left = FrameSupport.Left - 1
Loop
End If
End Sub
Private Sub Form_Load()
vrai = True
Date2 = Date
dateDebut = Date
dateFin = Date
If dateDebut = dateFin Then
Resultat = "Dates identiques"
End If
i = 0
Do While i < 999
Année.AddItem (i)
mois.AddItem (i)
jour.AddItem (i)
i = i + 1
Loop
i = 0
Do While i < 31
i = i + 1
If i < 10 Then
    jourChoix(0).AddItem ("0" & i)
    jourChoix(1).AddItem ("0" & i)
Else
    jourChoix(0).AddItem (i)
    jourChoix(1).AddItem (i)
End If
Loop
i = 0
Do While i < 12
i = i + 1
If i < 10 Then
    moisChoix(0).AddItem ("0" & i)
    moisChoix(1).AddItem ("0" & i)
Else
    moisChoix(0).AddItem (i)
    moisChoix(1).AddItem (i)
End If
Loop
i = 1959
Do While i < 2100
i = i + 1
annéeChoix(0).AddItem (i)
annéeChoix(1).AddItem (i)
Loop
End Sub

Private Sub Frame1_change()
End Sub



Private Sub Graphique_Click(Index As Integer)
Graph.Show
Unload Me
End Sub

Private Sub heur_Click(Index As Integer)
Unload Me
Heure.Show
End Sub

Private Sub Label7_Click()
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


Private Sub Mois1_Change()

End Sub

Private Sub Longueur_Click(Index As Integer)
Unload Me
Longue.Show
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

Private Sub Soustraire_Click()
ajout = False
soustrait = True
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
