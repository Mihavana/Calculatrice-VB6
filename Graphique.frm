VERSION 5.00
Begin VB.Form Graph 
   BackColor       =   &H00202020&
   Caption         =   "Calculatrice"
   ClientHeight    =   8280
   ClientLeft      =   225
   ClientTop       =   270
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   12780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fonct 
      BackColor       =   &H00353535&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   10080
      TabIndex        =   111
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
      Begin VB.CommandButton absol 
         BackColor       =   &H00353535&
         Caption         =   "|x|"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame messageAvertissement 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   8520
      TabIndex        =   107
      Top             =   1680
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton fermez 
         BackColor       =   &H000000FF&
         Caption         =   "x"
         Height          =   255
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   0
         Width           =   495
      End
      Begin VB.Label mess 
         BackStyle       =   0  'Transparent
         Caption         =   "Veuillez fermer la parenthèse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   109
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame FrameSupport 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   13000
      TabIndex        =   86
      Top             =   600
      Width           =   3015
      Begin VB.Frame FrameMenu 
         BackColor       =   &H00252525&
         BorderStyle     =   0  'None
         Caption         =   "cc"
         Height          =   10935
         Left            =   0
         TabIndex        =   88
         Top             =   0
         Width           =   2535
         Begin VB.Image Logos 
            Height          =   375
            Index           =   15
            Left            =   0
            Picture         =   "Graphique.frx":0000
            Stretch         =   -1  'True
            Top             =   10320
            Width           =   375
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   14
            Left            =   0
            Picture         =   "Graphique.frx":03D6
            Top             =   9600
            Width           =   480
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   13
            Left            =   0
            Picture         =   "Graphique.frx":07C6
            Top             =   9000
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   420
            Index           =   12
            Left            =   0
            Picture         =   "Graphique.frx":0B91
            Top             =   7680
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   420
            Index           =   11
            Left            =   0
            Picture         =   "Graphique.frx":0FC5
            Top             =   7080
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   570
            Index           =   10
            Left            =   0
            Picture         =   "Graphique.frx":139D
            Top             =   8280
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   9
            Left            =   0
            Picture         =   "Graphique.frx":1788
            Top             =   6480
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   8
            Left            =   0
            Picture         =   "Graphique.frx":1B2B
            Top             =   5880
            Width           =   480
         End
         Begin VB.Image Logos 
            Height          =   375
            Index           =   7
            Left            =   80
            Picture         =   "Graphique.frx":1F11
            Top             =   4800
            Width           =   330
         End
         Begin VB.Image Logos 
            Height          =   495
            Index           =   6
            Left            =   0
            Picture         =   "Graphique.frx":22FF
            Stretch         =   -1  'True
            Top             =   5280
            Width           =   450
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   5
            Left            =   120
            Picture         =   "Graphique.frx":26A2
            Top             =   4080
            Width           =   300
         End
         Begin VB.Image Logos 
            Height          =   525
            Index           =   4
            Left            =   50
            Picture         =   "Graphique.frx":2A5C
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
            TabIndex        =   106
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
            TabIndex        =   105
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
            TabIndex        =   104
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
            TabIndex        =   103
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
            TabIndex        =   102
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
            TabIndex        =   101
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
            TabIndex        =   100
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
            TabIndex        =   99
            Top             =   120
            Width           =   1935
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   3
            Left            =   120
            Picture         =   "Graphique.frx":2EBF
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
            TabIndex        =   98
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   0
            Left            =   120
            Picture         =   "Graphique.frx":32E8
            Top             =   480
            Width           =   390
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   1
            Left            =   120
            Picture         =   "Graphique.frx":372B
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   2
            Left            =   0
            Picture         =   "Graphique.frx":3AF7
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
            TabIndex        =   97
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
            TabIndex        =   96
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
            TabIndex        =   95
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
            TabIndex        =   94
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
            TabIndex        =   93
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
            Height          =   375
            Index           =   12
            Left            =   600
            TabIndex        =   92
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
            TabIndex        =   91
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
            TabIndex        =   90
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
            TabIndex        =   89
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
         TabIndex        =   87
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00353535&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   8280
      TabIndex        =   55
      Top             =   3600
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton tanhmoins 
         BackColor       =   &H00353535&
         Caption         =   "tanh-1"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cotanhmoins 
         BackColor       =   &H00353535&
         Caption         =   "coth-1"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cotanh 
         BackColor       =   &H00353535&
         Caption         =   "coth"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cschmoins 
         BackColor       =   &H00353535&
         Caption         =   "csch-1"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton csch 
         BackColor       =   &H00353535&
         Caption         =   "csch"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton secchmoins 
         BackColor       =   &H00353535&
         Caption         =   "sech-1"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton secch 
         BackColor       =   &H00353535&
         Caption         =   "sech"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton tanh 
         BackColor       =   &H00353535&
         Caption         =   "tanh"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton coshmoins 
         BackColor       =   &H00353535&
         Caption         =   "cosh-1"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cosh 
         BackColor       =   &H00353535&
         Caption         =   "cosh"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton sinhmoins 
         BackColor       =   &H00353535&
         Caption         =   "sinh-1"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton sinh 
         BackColor       =   &H00353535&
         Caption         =   "sinh"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton secondeTrigo1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "2nd"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton hyp1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "hyp"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton secondeTrigo 
         BackColor       =   &H00353535&
         Caption         =   "2nd"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton hyp 
         BackColor       =   &H00353535&
         Caption         =   "hyp"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton sin 
         BackColor       =   &H00353535&
         Caption         =   "sin"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cos 
         BackColor       =   &H00353535&
         Caption         =   "cos"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton secc 
         BackColor       =   &H00353535&
         Caption         =   "sec"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton csc 
         BackColor       =   &H00353535&
         Caption         =   "csc"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cotan 
         BackColor       =   &H00353535&
         Caption         =   "cot"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton tan 
         BackColor       =   &H00353535&
         Caption         =   "tan"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cotanMoins 
         BackColor       =   &H00353535&
         Caption         =   "cot-1"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cscMoins 
         BackColor       =   &H00353535&
         Caption         =   "csc-1"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton seccMoins 
         BackColor       =   &H00353535&
         Caption         =   "sec-1"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton tanMoins 
         BackColor       =   &H00353535&
         Caption         =   "tan-1"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cosMoins 
         BackColor       =   &H00353535&
         Caption         =   "cos-1"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton sinMoins 
         BackColor       =   &H00353535&
         Caption         =   "sin-1"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CommandButton ixs 
      BackColor       =   &H00353535&
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton ixs 
      BackColor       =   &H00353535&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   4560
      Width           =   735
   End
   Begin VB.Frame FrameCouleur 
      BackColor       =   &H00353535&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   9000
      TabIndex        =   45
      Top             =   1920
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Couleur"
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
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   3120
         TabIndex        =   53
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape shapePurple 
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   1
         Left            =   3120
         Shape           =   3  'Circle
         Top             =   720
         Width           =   495
      End
      Begin VB.Label yellow 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   2400
         TabIndex        =   49
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape shapePurple 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   0
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1560
         TabIndex        =   48
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape shapeGreen 
         BorderColor     =   &H00008000&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   1
         Left            =   1560
         Shape           =   3  'Circle
         Top             =   720
         Width           =   495
      End
      Begin VB.Label blue 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   840
         TabIndex        =   47
         Top             =   720
         Width           =   495
      End
      Begin VB.Label rouge 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00800000&
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   0
         Left            =   840
         Shape           =   3  'Circle
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   120
         Shape           =   3  'Circle
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.TextBox Panneau 
      BackColor       =   &H00202020&
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
      Height          =   495
      Index           =   3
      Left            =   8400
      TabIndex        =   44
      Top             =   2640
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox Panneau 
      BackColor       =   &H00202020&
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
      Height          =   495
      Index           =   2
      Left            =   8400
      TabIndex        =   43
      Top             =   2040
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox Panneau 
      BackColor       =   &H00202020&
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
      Height          =   495
      Index           =   1
      Left            =   8400
      TabIndex        =   42
      Top             =   1440
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox Panneau 
      BackColor       =   &H00202020&
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
      Height          =   495
      Index           =   0
      Left            =   8400
      TabIndex        =   41
      Top             =   840
      Width           =   3975
   End
   Begin VB.CommandButton Pi 
      BackColor       =   &H00353535&
      Caption         =   "Pi"
      BeginProperty Font 
         Name            =   "Blackadder ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton e 
      BackColor       =   &H00353535&
      Caption         =   "e"
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton unSurX 
      BackColor       =   &H00353535&
      Caption         =   "1/x"
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Module 
      BackColor       =   &H00353535&
      Caption         =   "|x|"
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton effacer 
      BackColor       =   &H00353535&
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
      Height          =   495
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton bracket 
      BackColor       =   &H00353535&
      Caption         =   "("
      Height          =   495
      Index           =   0
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton bracket 
      BackColor       =   &H00353535&
      Caption         =   ")"
      Height          =   495
      Index           =   1
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00353535&
      Caption         =   "="
      Height          =   495
      Index           =   0
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton signes 
      BackColor       =   &H00505050&
      Caption         =   "(-)"
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton virgule 
      BackColor       =   &H00505050&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton égale 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ß"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00353535&
      Caption         =   "+"
      Height          =   495
      Index           =   1
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00353535&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00353535&
      Caption         =   "X"
      Height          =   495
      Index           =   3
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00353535&
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton clear 
      BackColor       =   &H00353535&
      Caption         =   "C"
      Height          =   495
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton sec 
      BackColor       =   &H00C0C0FF&
      Caption         =   "2nd"
      Height          =   495
      Left            =   8400
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton second 
      BackColor       =   &H00353535&
      Caption         =   "2nd"
      Height          =   495
      Left            =   8400
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Logarithme 
      BackColor       =   &H00353535&
      Caption         =   "ln"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton logDix 
      BackColor       =   &H00353535&
      Caption         =   "log"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton dixPuissance 
      BackColor       =   &H00353535&
      Caption         =   "10^x"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton puissanceSpec 
      BackColor       =   &H00353535&
      Caption         =   "x^y"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton puissance 
      BackColor       =   &H00353535&
      Caption         =   "x²"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton expo 
      BackColor       =   &H00353535&
      Caption         =   "e^x"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton deuxPuissance 
      BackColor       =   &H00353535&
      Caption         =   "2^x"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton racineSpec 
      BackColor       =   &H00353535&
      Caption         =   "y^Vx"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton puissanceTrois 
      BackColor       =   &H00353535&
      Caption         =   "x^3"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton racineCarré 
      BackColor       =   &H00353535&
      Caption         =   "²Vx"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton racineCubique 
      BackColor       =   &H00353535&
      Caption         =   "3^Vx"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   13.044
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   14.446
      TabIndex        =   0
      Top             =   840
      Width           =   8250
      Begin VB.Line Line36 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   0.011
         Y2              =   0.011
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   1.011
         Y2              =   1.011
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   2.011
         Y2              =   2.011
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   3.009
         Y2              =   3.009
      End
      Begin VB.Line Line32 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   4.009
         Y2              =   4.009
      End
      Begin VB.Line Line31 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   5.009
         Y2              =   5.009
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   6.01
         Y2              =   6.01
      End
      Begin VB.Line Line29 
         X1              =   0
         X2              =   14.393
         Y1              =   13
         Y2              =   13
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   12.01
         Y2              =   12.01
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   11.01
         Y2              =   11.01
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   10.01
         Y2              =   10.01
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   9.01
         Y2              =   9.01
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   14.393
         Y1              =   8.01
         Y2              =   8.01
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   14
         X2              =   14
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   0.011
         X2              =   0.011
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   1.011
         X2              =   1.011
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   2.011
         X2              =   2.011
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   3.009
         X2              =   3.009
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   4.009
         X2              =   4.009
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   5.009
         X2              =   5.009
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   6.01
         X2              =   6.01
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   13.01
         X2              =   13.01
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   12.01
         X2              =   12.01
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   11.01
         X2              =   11.01
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   10.01
         X2              =   10.01
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   9.01
         X2              =   9.01
         Y1              =   0
         Y2              =   13.123
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   8.043
         X2              =   8.043
         Y1              =   0
         Y2              =   13.335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   6.985
         X2              =   6.985
         Y1              =   -0.423
         Y2              =   13.123
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   14.393
         X2              =   13.97
         Y1              =   6.985
         Y2              =   7.197
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   14.393
         X2              =   13.97
         Y1              =   6.985
         Y2              =   6.773
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   6.985
         X2              =   7.197
         Y1              =   0
         Y2              =   0.423
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   6.985
         X2              =   6.773
         Y1              =   0
         Y2              =   0.423
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   0
         X2              =   14.393
         Y1              =   6.985
         Y2              =   6.985
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00202020&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   12120
      TabIndex        =   50
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Fonc 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      Caption         =   "Fonction"
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
      Height          =   255
      Left            =   9960
      TabIndex        =   110
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Active 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11640
      TabIndex        =   52
      Top             =   3480
      Width           =   255
   End
   Begin VB.Image Supprimer 
      Height          =   375
      Index           =   3
      Left            =   12360
      Picture         =   "Graphique.frx":3F07
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Supprimer 
      Height          =   375
      Index           =   2
      Left            =   12360
      Picture         =   "Graphique.frx":4350
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Supprimer 
      Height          =   375
      Index           =   1
      Left            =   12360
      Picture         =   "Graphique.frx":4799
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Supprimer 
      Height          =   375
      Index           =   0
      Left            =   12360
      Picture         =   "Graphique.frx":4BE2
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00202020&
      Caption         =   "Graphique"
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
      Left            =   10320
      TabIndex        =   51
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image color 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   12000
      Picture         =   "Graphique.frx":502B
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label trigo 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      Caption         =   "Trigonométrie"
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
      Height          =   255
      Left            =   8280
      TabIndex        =   40
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   12240
      X2              =   12600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   12240
      X2              =   12600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   12240
      X2              =   12600
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j, l
Dim X, Y As Double
Dim valExcel, valPan
Dim traceX(10000), traceY(10000) As Double
Dim m(5), tabPanneau(10000), memoire(4)
Dim A, numTabPan
Dim n
Dim chiff, oper, ix, brac
Dim couleur, couleur1(4)
Dim panMem(4)
Dim vir, egale, operateur, parenthese, activeExc, initial, brack
Dim modulo, yroot, logBase, puisSpec
Dim mem, tmp
Dim deg, rad
Dim positionPar, panMod
Dim panneauSansPar
Dim segundo, segundo1
Dim lnep, lG, squirt
Dim loog, loog1
Dim AppExcel As Excel.Application
Dim wbExcel As Excel.Workbook
Dim wsExcel As Excel.Worksheet

Private Sub absol_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "ABS" & "("
Else
Panneau(l) = "ABS" & "("
End If
End Sub

Private Sub absol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
absol.BackColor = &H808080
End Sub

Private Sub Ang_Click(Index As Integer)
Unload Me
Angle.Show
End Sub

Private Sub blue_Click()
couleur = vbBlue
FrameCouleur.Visible = False
Active.BackColor = &H800000

End Sub

Private Sub bracket_Click(o As Integer)
If o = 0 Then
    If Panneau(l) <> "" Then
    If Right(Panneau(l), 1) <> m(n) Then
    Panneau(l) = Panneau(l) & "*" & "("
    ElseIf Right(Panneau(l), 1) = m(n) Then
    Panneau(l) = Panneau(l) & "("
    Else
    Panneau(l) = "("
    End If
    End If
    brack = brack + 1
ElseIf o = 1 And brack > 0 Then
Panneau(l) = Panneau(l) & ")"
brack = brack - 1
End If
parenthese = True
End Sub

Private Sub calculDeLaDate_Click()
Stand.Hide
Scient.Hide
Graph.Hide
calculDate.Show
End Sub

Private Sub bracket_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
bracket(brac).BackColor = &H353535
bracket(o).BackColor = &H808080
brac = o
End Sub

Private Sub CalculDedate_Click(Index As Integer)
Unload Me
CalculdelaDate.Show
End Sub

Private Sub ceill_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "CEILING" & "("
Else
Panneau(l) = "CEILING" & "("
End If
End Sub

Private Sub chiffre_Click(i As Integer)
numTabPan = numTabPan + 1
tabPanneau(numTabPan) = i
Panneau(l) = Panneau(l) & tabPanneau(numTabPan)
valPan = Panneau(l) & tabPanneau(numTabPan)
If Panneau(l) = "" Then
    Panneau(l) = i
    valPan = i
End If

End Sub

Private Sub Chiffre_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
chiffre(chiff).BackColor = &H505050
chiffre(o).BackColor = &H808080
chiff = o
End Sub

Private Sub clear_Click()
Panneau(l) = ""
A = 0
panMod = 0
vir = True
vir = True
operateur = False
egale = False
initial = True
brack = 0
End Sub



Private Sub clear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clear.BackColor = &H808080
End Sub

Private Sub color_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
color.Appearance = 0
End Sub

Private Sub cos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cos.BackColor = &H808080
End Sub

Private Sub cosh_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "cosh" & "("
Else
Panneau(l) = "cosh" & "("
End If
parenthese = False
End Sub

Private Sub cosh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cosh.BackColor = &H808080
End Sub

Private Sub coshmoins_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "acosh" & "("
Else
Panneau(l) = "acosh" & "("
End If
parenthese = False
End Sub

Private Sub coshmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
coshmoins.BackColor = &H808080
End Sub

Private Sub cosMoins_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "acos" & "("
Else
Panneau(l) = "acos" & "("
End If
parenthese = False
End Sub

Private Sub cosMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cosMoins.BackColor = &H808080
End Sub

Private Sub cotan_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "cot" & "("
Else
Panneau(l) = "cot" & "("
End If
parenthese = False
End Sub

Private Sub cotan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cotan.BackColor = &H808080
End Sub

Private Sub cotanh_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "coth" & "("
Else
Panneau(l) = "coth" & "("
End If
parenthese = False
End Sub

Private Sub cotanh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cotanh.BackColor = &H808080
End Sub

Private Sub cotanhmoins_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "acoth" & "("
Else
Panneau(l) = "acoth" & "("
End If
parenthese = False
End Sub

Private Sub cotanhmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cotanhmoins.BackColor = &H808080
End Sub

Private Sub cotanMoins_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "acot" & "("
Else
Panneau(l) = "acot" & "("
End If
parenthese = False

End Sub

Private Sub cotanMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cotanMoins.BackColor = &H808080
End Sub

Private Sub csc_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & 1 & "/" & "sin" & "("
Else
Panneau(l) = 1 & "/" & "sin" & "("
End If
parenthese = False
End Sub

Private Sub csc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
csc.BackColor = &H808080
End Sub

Private Sub csch_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & 1 & "/" & "sinh" & "("
Else
Panneau(l) = 1 & "/" & "sinh" & "("
End If
parenthese = False
End Sub

Private Sub csch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
csch.BackColor = &H808080
End Sub

Private Sub cschmoins_Click()
If Panneau(l) <> "" Then
brack = brack + 1
Panneau(l) = Panneau(l) & 1 & "/" & "asinh" & "("
Else
Panneau(l) = 1 & "/" & "asinh" & "("
End If
parenthese = False

End Sub

Private Sub cschmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cschmoins.BackColor = &H808080
End Sub

Private Sub cscMoins_Click()
If Panneau(l) <> "" Then
brack = brack + 1
Panneau(l) = Panneau(l) & 1 & "/" & "asin" & "("
Else
Panneau(l) = 1 & "/" & "asin" & "("
End If
parenthese = False

End Sub

Private Sub factoriel_Click()

End Sub

Private Sub cscMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cscMoins.BackColor = &H808080
End Sub

Private Sub deuxPuissance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
deuxPuissance.BackColor = &H808080
End Sub

Private Sub dixPuissance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
dixPuissance.BackColor = &H808080
End Sub

Private Sub don_Click(Index As Integer)
Unload Me
Données.Show
End Sub

Private Sub e_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
e.BackColor = &H808080
End Sub

Private Sub effacer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
effacer.BackColor = &H808080
End Sub

Private Sub égale_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
égale.BackColor = &H8080FF
End Sub

Private Sub energ_Click(Index As Integer)
Unload Me
Energie.Show
End Sub

Private Sub expo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
expo.BackColor = &H808080
End Sub

Private Sub fermez_Click()
messageAvertissement.Visible = False
End Sub

Private Sub floorr_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "FLOOR" & "("
Else
Panneau(l) = "FLOOR" & "("
End If
End Sub

Private Sub Fonc_Click()
If fonct.Visible = False Then
fonct.Visible = True
Else
fonct.Visible = False
End If
End Sub

Private Sub Fonc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
trigo.BackColor = &H202020
Fonc.BackColor = &H808080
End Sub

Private Sub fonct_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
absol.BackColor = &H353535
End Sub

Private Sub Form_Click()
If FrameSupport.Left < 13000 Then
Do While FrameSupport.Left < 13000
FrameSupport.Left = FrameSupport.Left + 10
Loop
End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
cschmoins.BackColor = &H353535
cscMoins.BackColor = &H353535
csch.BackColor = &H353535
csc.BackColor = &H353535
cotanhmoins.BackColor = &H353535
cotanMoins.BackColor = &H353535
cotanh.BackColor = &H353535
cotan.BackColor = &H353535
tanhmoins.BackColor = &H353535
tanMoins.BackColor = &H353535
tanh.BackColor = &H353535
tan.BackColor = &H353535
coshmoins.BackColor = &H353535
cosMoins.BackColor = &H353535
cosh.BackColor = &H353535
cos.BackColor = &H353535
sinhmoins.BackColor = &H353535
sinMoins.BackColor = &H353535
sinh.BackColor = &H353535
sin.BackColor = &H353535
seccMoins.BackColor = &H353535
secchmoins.BackColor = &H353535
secch.BackColor = &H353535
secc.BackColor = &H353535
Fonc.BackColor = &H202020
trigo.BackColor = &H202020
hyp.BackColor = &H353535
secondeTrigo.BackColor = &H353535
hyp1.BackColor = &HC0C0FF
secondeTrigo1.BackColor = &HC0C0FF
absol.BackColor = &H353535
color.Appearance = 1
Pi.BackColor = &H353535
e.BackColor = &H353535
clear.BackColor = &H353535
Module.BackColor = &H353535
unSurX.BackColor = &H353535
second.BackColor = &H353535
sec.BackColor = &HC0C0FF
puissanceTrois.BackColor = &H353535
puissance.BackColor = &H353535
bracket(brac).BackColor = &H353535
racineCubique.BackColor = &H353535
racineCarré.BackColor = &H353535
racineSpec.BackColor = &H353535
puissanceSpec.BackColor = &H353535
deuxPuissance.BackColor = &H353535
dixPuissance.BackColor = &H353535
expo.BackColor = &H353535
logDix.BackColor = &H353535
Logarithme.BackColor = &H353535
ixs(ix).BackColor = &H353535
effacer.BackColor = &H353535
égale.BackColor = &HC0C0FF
Virgule.BackColor = &H505050
signes.BackColor = &H505050
operation(oper).BackColor = &H353535
chiffre(chiff).BackColor = &H505050
If Frame1.Visible = True Then
Frame1.Visible = False
End If
If fonct.Visible = True Then
fonct.Visible = False
End If
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cschmoins.BackColor = &H353535
cscMoins.BackColor = &H353535
csch.BackColor = &H353535
csc.BackColor = &H353535
cotanhmoins.BackColor = &H353535
cotanMoins.BackColor = &H353535
cotanh.BackColor = &H353535
cotan.BackColor = &H353535
tanhmoins.BackColor = &H353535
tanMoins.BackColor = &H353535
tanh.BackColor = &H353535
tan.BackColor = &H353535
coshmoins.BackColor = &H353535
cosMoins.BackColor = &H353535
cosh.BackColor = &H353535
cos.BackColor = &H353535
sinhmoins.BackColor = &H353535
sinMoins.BackColor = &H353535
sinh.BackColor = &H353535
sin.BackColor = &H353535
seccMoins.BackColor = &H353535
secchmoins.BackColor = &H353535
secch.BackColor = &H353535
secc.BackColor = &H353535
Fonc.BackColor = &H202020
trigo.BackColor = &H202020
hyp.BackColor = &H353535
secondeTrigo.BackColor = &H353535
hyp1.BackColor = &HC0C0FF
secondeTrigo1.BackColor = &HC0C0FF
End Sub

Private Sub heur_Click(Index As Integer)
Unload Me
Heure.Show
End Sub

Private Sub hyp_Click()
hyp1.Visible = True
If segundo1 = False Then
sinh.Visible = True
cosh.Visible = True
tanh.Visible = True
secch.Visible = True
csch.Visible = True
cotanh.Visible = True
ElseIf segundo1 = True Then
sinhmoins.Visible = True
coshmoins.Visible = True
tanhmoins.Visible = True
secchmoins.Visible = True
cschmoins.Visible = True
cotanhmoins.Visible = True
End If
segundo = True
End Sub

Private Sub hyp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
hyp.BackColor = &H808080
End Sub

Private Sub hyp1_Click()
If segundo1 = True Then
sin.Visible = False
cos.Visible = False
tan.Visible = False
secc.Visible = False
csc.Visible = False
cotan.Visible = False
ElseIf segundo1 = False Then
sin.Visible = True
cos.Visible = True
tan.Visible = True
secc.Visible = True
csc.Visible = True
cotan.Visible = True
End If
sinh.Visible = False
cosh.Visible = False
tanh.Visible = False
secch.Visible = False
csch.Visible = False
cotanh.Visible = False

sinhmoins.Visible = False
coshmoins.Visible = False
tanhmoins.Visible = False
secchmoins.Visible = False
cschmoins.Visible = False
cotanhmoins.Visible = False

hyp1.Visible = False
segundo = False
End Sub






Private Sub hyp1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
hyp1.BackColor = &H8080FF
End Sub

Private Sub ixs_Click(i As Integer)
    If i = 0 Then
        If Panneau(l) <> "" Then
            If Right(Panneau(l), 15) = "333333333333333" Then
            Panneau(l) = Replace(Panneau(l), "(", "(x")
            Else
            Panneau(l) = Panneau(l) & "x"
        End If
    Else
        Panneau(l) = "x"
    End If
valPan = "x"
End If
End Sub

Private Sub ixs_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ixs(ix).BackColor = &H353535
ixs(o).BackColor = &H808080
ix = o
End Sub

Private Sub Label1_Click()
couleur = vbGreen
FrameCouleur.Visible = False
Active.BackColor = &H8000&
End Sub

Private Sub Degré_Click()
deg = False
Degré.Visible = False
rad = True
radiant.Visible = True
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub color_Click()
If FrameCouleur.Visible = False Then
FrameCouleur.Visible = True
Else
FrameCouleur.Visible = False
End If
End Sub

Private Sub cos_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "cos" & "("
Else
Panneau(l) = "cos" & "("
End If
parenthese = False
End Sub

Private Sub deuxPuissance_Click()
Panneau(l) = "2" & "^" & "("
End Sub

Private Sub dixPuissance_Click()
Panneau(l) = "10" & "^" & "("
brack = brack + 1
End Sub

Private Sub e_Click()
Panneau(l) = 2.71828182845905
End Sub

Private Sub effacer_Click()
If Right(Panneau(l), 1) = ")" Then
brack = brack + 1
Panneau(l) = Left(Panneau(l), Len(Panneau(l)) - 1)
ElseIf Right(Panneau(l), 1) = "(" Then
brack = brack - 1
Panneau(l) = Left(Panneau(l), Len(Panneau(l)) - 1)
ElseIf Panneau(l) <> "" Then
Panneau(l) = Left(Panneau(l), Len(Panneau(l)) - 1)
End If
If InStr(1, Panneau(l), ",") = 0 Then
vir = True
End If
End Sub

Private Sub égale_Click()
Dim erreur
If brack > 0 Then
messageAvertissement.Visible = True
Exit Sub
End If
On Error Resume Next
Picture1.DrawWidth = 2
Picture1.Line (7, 0)-(7, 1000)
Picture1.Line (0, 7)-(1000, 7)
Picture1.DrawWidth = 1
i = 0
j = 1
Do While j < 14
Picture1.CurrentX = j
Picture1.CurrentY = 7
Picture1.Print (-7 + j)
j = j + 1
Loop
Do While j > 0
Picture1.CurrentX = 7
Picture1.CurrentY = j
Picture1.Print (7 - j)
j = j - 1
Loop
Picture1.DrawWidth = 3
memoire(l) = Panneau(l)
Do While X < 10
valExcel = "=" & Replace(Panneau(l), "x", X)
valExcel = Replace(valExcel, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = valExcel
Y = Worksheets("Feuil1").Cells(1, 1).Value

traceX(i) = X
traceY(i) = Y

If traceY(i) = traceY(i - 1) Then
erreur = True
GoTo suivant
End If



If erreur = True Then
    If i > 0 Then
        Picture1.Line (7 + traceX(i), 7 - traceY(i))-(7 + traceX(i), 7 - traceY(i))
    End If
erreur = False
GoTo suivant
End If

If i > 0 Then
Picture1.Line (7 + traceX(i - 1), 7 - traceY(i - 1))-(7 + traceX(i), 7 - traceY(i)), couleur
End If

suivant:
i = i + 1
X = X + 0.08

Loop
couleur1(l) = couleur
Panneau(l + 1).Visible = True
Panneau(l + 1) = ""
Supprimer(l).Visible = True
l = l + 1
vir = True
X = -10
i = 0
End Sub

Private Sub ex_Click()

End Sub

Private Sub expo_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "EXP" & "("
Else
Panneau(l) = "EXP" & "("
End If
parenthese = False
End Sub

Private Sub Form_Load()
brack = 0
couleur = vbBlue
loog = False
loog1 = False
initial = True
vir = True
parenthese = False
Panneau(l) = ""
tmp = 0
A = 0
X = -10
m(1) = "+"
m(2) = "-"
m(3) = "*"
m(4) = "/"
Set AppExcel = CreateObject("Excel.Application")
Set wbExcel = AppExcel.Workbooks.Add
Set wsExcel = wbExcel.Worksheets("Feuil1")
End Sub




Private Sub Label2_Click()
If FrameSupport.Left < 13000 Then
Do While FrameSupport.Left < 13000
FrameSupport.Left = FrameSupport.Left + 10
Loop
Else
Do While FrameSupport.Left > 10000
FrameSupport.Left = FrameSupport.Left - 10
Loop
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 1
End Sub

Private Sub Label3_Click()
Label2.BorderStyle = 0
End Sub

Private Sub Label4_Click()
couleur = vbBlack
FrameCouleur.Visible = False
Active.BackColor = &H0&

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Logarithme_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "LN" & "("
Else
Panneau(l) = "LN" & "("
End If
End Sub

Private Sub Logarithme_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Logarithme.BackColor = &H808080
End Sub

Private Sub logDix_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "LOG" & "("
Else
Panneau(l) = "LOG" & "("
End If
parenthese = False
End Sub

Private Sub logyX_Click()
Panneau(l) = "LOG" & "("
loog = True
End Sub

Private Sub logDix_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
logDix.BackColor = &H808080
End Sub

Private Sub Longueur_Click(Index As Integer)
Unload Me
Longue.Show
End Sub


Private Sub Module_Click()
Panneau(l) = "ABS" & "(" & ")"
parenthese = False
End Sub

Private Sub Module_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Module.BackColor = &H808080
End Sub

Private Sub operation_Click(k As Integer)
n = k
m(n) = m(k)
vir = True
numTabPan = numTabPan + 1
tabPanneau(numTabPan) = m(k)
Panneau(l) = Panneau(l) & tabPanneau(numTabPan)
If Panneau(l) = "" Then
Panneau(l) = m(k)
End If
End Sub



Private Sub operation_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
operation(oper).BackColor = &H353535
operation(o).BackColor = &H808080
oper = o
End Sub

Private Sub Panneau_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 0
End Sub

Private Sub Pi_Click()
Panneau(l) = 3.14159265358979
End Sub

Private Sub Pi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pi.BackColor = &H808080
End Sub

Private Sub Picture1_Click()
If FrameSupport.Left < 13000 Then
Do While FrameSupport.Left < 13000
FrameSupport.Left = FrameSupport.Left + 10
Loop
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cschmoins.BackColor = &H353535
cscMoins.BackColor = &H353535
csch.BackColor = &H353535
csc.BackColor = &H353535
cotanhmoins.BackColor = &H353535
cotanMoins.BackColor = &H353535
cotanh.BackColor = &H353535
cotan.BackColor = &H353535
tanhmoins.BackColor = &H353535
tanMoins.BackColor = &H353535
tanh.BackColor = &H353535
tan.BackColor = &H353535
coshmoins.BackColor = &H353535
cosMoins.BackColor = &H353535
cosh.BackColor = &H353535
cos.BackColor = &H353535
sinhmoins.BackColor = &H353535
sinMoins.BackColor = &H353535
sinh.BackColor = &H353535
sin.BackColor = &H353535
seccMoins.BackColor = &H353535
secchmoins.BackColor = &H353535
secch.BackColor = &H353535
secc.BackColor = &H353535
Fonc.BackColor = &H202020
trigo.BackColor = &H202020
hyp.BackColor = &H353535
secondeTrigo.BackColor = &H353535
hyp1.BackColor = &HC0C0FF
secondeTrigo1.BackColor = &HC0C0FF
absol.BackColor = &H353535
color.Appearance = 1
Pi.BackColor = &H353535
e.BackColor = &H353535
clear.BackColor = &H353535
Module.BackColor = &H353535
unSurX.BackColor = &H353535
second.BackColor = &H353535
sec.BackColor = &HC0C0FF
puissanceTrois.BackColor = &H353535
puissance.BackColor = &H353535
bracket(brac).BackColor = &H353535
racineCubique.BackColor = &H353535
racineCarré.BackColor = &H353535
racineSpec.BackColor = &H353535
puissanceSpec.BackColor = &H353535
deuxPuissance.BackColor = &H353535
dixPuissance.BackColor = &H353535
expo.BackColor = &H353535
logDix.BackColor = &H353535
Logarithme.BackColor = &H353535
ixs(ix).BackColor = &H353535
effacer.BackColor = &H353535
égale.BackColor = &HC0C0FF
Virgule.BackColor = &H505050
signes.BackColor = &H505050
operation(oper).BackColor = &H353535
chiffre(chiff).BackColor = &H505050
If Frame1.Visible = True Then
Frame1.Visible = False
End If
If fonct.Visible = True Then
fonct.Visible = False
End If
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
Project1.puissance.Show
End Sub

Private Sub puissance_Click()
Panneau(l) = Panneau(l) & "^" & "2"
End Sub

Private Sub puissance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
puissance.BackColor = &H808080
End Sub

Private Sub puissanceSpec_Click()
Panneau(l) = Panneau(l) & "^"
End Sub

Private Sub puissanceSpec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
puissanceSpec.BackColor = &H808080
End Sub

Private Sub puissanceTrois_Click()
Panneau(l) = Panneau(l) & "^" & "3"
End Sub

Private Sub Purple_Click()


End Sub

Private Sub puissanceTrois_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
puissanceTrois.BackColor = &H808080
End Sub

Private Sub racineCarré_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "sqrt" & "("
Else
Panneau(l) = "sqrt" & "("
End If
parenthese = False
End Sub

Private Sub racineCarré_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
racineCarré.BackColor = &H808080
End Sub

Private Sub racineCubique_Click()
Panneau(l) = "(" & ")" & "^" & 1 / 3
parenthese = False
End Sub

Private Sub racineCubique_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
racineCubique.BackColor = &H808080
End Sub

Private Sub racineSpec_Click()
Panneau(l) = Panneau(l) & "^" & 1 & "/"
End Sub

Private Sub radiant_Click()
rad = False
radiant.Visible = False
deg = True
Degré.Visible = True
End Sub

Private Sub racineSpec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
racineSpec.BackColor = &H808080
End Sub

Private Sub rouge_Click()
couleur = vbRed
FrameCouleur.Visible = False
Active.BackColor = &HFF&
End Sub

Private Sub Scientifique_Click(Index As Integer)
Unload Me
Scient.Show
End Sub

Private Sub sec_Click()
sec.Visible = False
second.Visible = True
puissance.Visible = True
racineCarré.Visible = True
puissanceSpec.Visible = True
dixPuissance.Visible = True
logDix.Visible = True
Logarithme.Visible = True
End Sub

Private Sub sec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sec.BackColor = &H8080FF
End Sub

Private Sub secc_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & 1 & "/" & "cos" & "("
Else
Panneau(l) = 1 & "/" & "cos" & "("
End If
parenthese = False
End Sub

Private Sub secc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secc.BackColor = &H808080
End Sub

Private Sub secch_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & 1 & "/" & "cosh" & "("
Else
Panneau(l) = 1 & "/" & "cosh" & "("
End If
parenthese = False
End Sub

Private Sub secch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secch.BackColor = &H808080
End Sub

Private Sub secchmoins_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & 1 & "/" & "acosh" & "("
Else
Panneau(l) = 1 & "/" & "acosh" & "("
End If
parenthese = False

End Sub

Private Sub secchmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secchmoins.BackColor = &H808080
End Sub

Private Sub seccMoins_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & 1 & "/" & "acos" & "("
Else
Panneau(l) = 1 & "/" & "acos" & "("
End If
parenthese = False
End Sub

Private Sub seccMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
seccMoins.BackColor = &H808080
End Sub

Private Sub second_Click()
second.Visible = False
sec.Visible = True
puissance.Visible = False
racineCarré.Visible = False
puissanceSpec.Visible = False
dixPuissance.Visible = False
Logarithme.Visible = False
End Sub

Private Sub second_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
second.BackColor = &H808080
End Sub

Private Sub secondeTrigo_Click()
If segundo = False Then
sin.Visible = False
cos.Visible = False
tan.Visible = False
secc.Visible = False
csc.Visible = False
cotan.Visible = False
secondeTrigo1.Visible = True
Else
sinhmoins.Visible = True
coshmoins.Visible = True
tanhmoins.Visible = True
secchmoins.Visible = True
cschmoins.Visible = True
cotanhmoins.Visible = True
End If
secondeTrigo1.Visible = True
segundo1 = True
End Sub

Private Sub secondeTrigo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secondeTrigo.BackColor = &H808080
End Sub

Private Sub secondeTrigo1_Click()
secondeTrigo1.Visible = False
If segundo = False Then
sin.Visible = True
cos.Visible = True
tan.Visible = True
secc.Visible = True
csc.Visible = True
cotan.Visible = True
Else
sinh.Visible = True
cosh.Visible = True
tanh.Visible = True
secch.Visible = True
csch.Visible = True
cotanh.Visible = True
End If
sinhmoins.Visible = False
coshmoins.Visible = False
tanhmoins.Visible = False
secchmoins.Visible = False
cschmoins.Visible = False
cotanhmoins.Visible = False

segundo1 = False
End Sub

Private Sub secondeTrigo1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secondeTrigo1.BackColor = &H8080FF
End Sub

Private Sub signes_Click()
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "-"
Else
Panneau(l) = "-"
End If
End Sub

Private Sub signes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
signes.BackColor = &H808080
End Sub

Private Sub sin_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "sin" & "("
Else
Panneau(l) = "sin" & "("
End If
parenthese = False
End Sub

Private Sub sin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sin.BackColor = &H808080
End Sub

Private Sub sinh_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "sinh" & "("
Else
Panneau(l) = "sinh" & "("
End If
parenthese = False
End Sub

Private Sub sinh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sinh.BackColor = &H808080
End Sub

Private Sub sinhmoins_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "asinh" & "("
Else
Panneau(l) = "asinh" & "("
End If
End Sub

Private Sub sinhmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sinhmoins.BackColor = &H808080
End Sub

Private Sub sinMoins_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "asin" & "("
Else
Panneau(l) = "asin" & "("
End If
parenthese = False
End Sub



Private Sub sinMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sinMoins.BackColor = &H808080
End Sub

Private Sub Standard_Click(Index As Integer)
Unload Me
Stand.Show
End Sub

Private Sub Supprimer_Click(n As Integer)
Dim erreur, q
Panneau(n) = ""
Supprimer(n).Visible = False
If l = 4 Then l = 3
If l > 0 Then
    If n >= 0 Then
        If n = l - 1 Then
        If n > 0 Then Panneau(n).Visible = False
        Panneau(n + 1).Visible = False
        ElseIf n < l - 1 Then
        Supprimer(n).Visible = True
            Do While n < l
            Panneau(n).Text = Panneau(n + 1).Text
            memoire(n) = Panneau(n)
            couleur1(n) = couleur1(n + 1)
            Panneau(n + 1).Visible = False
            Supprimer(n + 1).Visible = False
            n = n + 1
            Loop
        End If
    End If
l = l - 1
End If
q = l
Picture1.Refresh
Picture1.DrawWidth = 1
X = -10
i = 0
j = 1
On Error Resume Next
Do While j < 14
Picture1.CurrentX = j
Picture1.CurrentY = 7
Picture1.Print (-7 + j)
j = j + 1
Loop
Do While j > 0
Picture1.CurrentX = 7
Picture1.CurrentY = j
Picture1.Print (7 - j)
j = j - 1
Loop
If l > 0 Then
Picture1.DrawWidth = 3
Do While q > 0
X = -10
i = 0
traceX(i) = 0
traceY(i) = 0
Picture1.PSet (7 + traceX(i), 7 - traceY(i))
Do While X < 10
valExcel = "=" & Replace(memoire(q - 1), "x", X)
valExcel = Replace(valExcel, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = valExcel
Y = Worksheets("Feuil1").Cells(1, 1).Value

traceX(i) = X
traceY(i) = Y

If traceY(i) = traceY(i - 1) Then
erreur = True
GoTo suivant
End If


If erreur = True Then
    If i > 0 Then
        Picture1.Line (7 + traceX(i), 7 - traceY(i))-(7 + traceX(i), 7 - traceY(i))
    End If
erreur = False
GoTo suivant
End If

If i > 0 Then
Picture1.Line (7 + traceX(i - 1), 7 - traceY(i - 1))-(7 + traceX(i), 7 - traceY(i)), couleur1(q - 1)
End If

suivant:
i = i + 1
X = X + 0.08
Loop
q = q - 1
Loop
End If
End Sub

Private Sub surf_Click(Index As Integer)
Unload Me
Surface.Show
End Sub

Private Sub tan_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "tan" & "("
Else
Panneau(l) = "tan" & "("
End If
parenthese = False
End Sub

Private Sub tan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tan.BackColor = &H808080
End Sub

Private Sub tanh_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "tanh" & "("
Else
Panneau(l) = "tanh" & "("
End If
parenthese = False
End Sub

Private Sub tanh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tanh.BackColor = &H808080
End Sub

Private Sub tanhmoins_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "atanh" & "("
Else
Panneau(l) = "atanh" & "("
End If
parenthese = False
End Sub

Private Sub tanhmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tanhmoins.BackColor = &H808080
End Sub

Private Sub tanMoins_Click()
brack = brack + 1
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & "atan" & "("
Else
Panneau(l) = "atan" & "("
End If
parenthese = False

End Sub

Private Sub tanMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tanMoins.BackColor = &H808080
End Sub

Private Sub température_Click(Index As Integer)
Unload Me
Temp.Show
End Sub

Private Sub trigo_Click()
If Frame1.Visible = False Then
Frame1.Visible = True
ElseIf Frame1.Visible = True Then
Frame1.Visible = False
End If
End Sub


Private Sub trigo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Fonc.BackColor = &H202020
trigo.BackColor = &H808080
End Sub

Private Sub unSurX_Click()
If Panneau(l) <> "" Then
If Right(Panneau(l), 1) = ")" And parenthese = False Then
    Panneau(l) = Left(Panneau(l), Len(Panneau(l)) - 1) & 1 & "/" & Right(Panneau(l), 1)
    Else
    Panneau(l) = Panneau(l) & 1 & "/"
    End If
Else
Panneau(l) = 1 & "/"
End If
End Sub

Private Sub unSurX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
unSurX.BackColor = &H808080
End Sub

Private Sub Virgule_Click()
If vir = True Then
If Panneau(l) <> "" Then
Panneau(l) = Panneau(l) & ","
vir = False
ElseIf Panneau(l) = "" Then
Panneau(l) = "0" & ","
vir = False
End If
End If
End Sub


Private Sub virgule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Virgule.BackColor = &H808080
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

Private Sub yellow_Click()
couleur = vbYellow
FrameCouleur.Visible = False
Active.BackColor = &HFFFF&
End Sub

