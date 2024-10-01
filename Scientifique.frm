VERSION 5.00
Begin VB.Form Scient 
   BackColor       =   &H00202020&
   Caption         =   "Calculatrice"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   8778.325
   ScaleMode       =   0  'User
   ScaleWidth      =   17941.23
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameSupport 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   -4033
      TabIndex        =   99
      Top             =   600
      Width           =   2775
      Begin VB.VScrollBar VScroll1 
         Height          =   7215
         LargeChange     =   300
         Left            =   2520
         Max             =   4100
         SmallChange     =   300
         TabIndex        =   100
         Top             =   0
         Width           =   255
      End
      Begin VB.Frame FrameMenu 
         BackColor       =   &H00252525&
         BorderStyle     =   0  'None
         Caption         =   "cc"
         Height          =   10935
         Left            =   0
         TabIndex        =   101
         Top             =   0
         Width           =   2535
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
            TabIndex        =   119
            Top             =   10320
            Width           =   975
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
            TabIndex        =   118
            Top             =   9720
            Width           =   1095
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
            TabIndex        =   117
            Top             =   9120
            Width           =   1335
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
            TabIndex        =   116
            Top             =   8400
            Width           =   1455
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
            TabIndex        =   115
            Top             =   7800
            Width           =   1095
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
            TabIndex        =   114
            Top             =   7200
            Width           =   1095
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
            TabIndex        =   113
            Top             =   600
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
            TabIndex        =   112
            Top             =   1200
            Width           =   2055
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
            TabIndex        =   111
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   2
            Left            =   0
            Picture         =   "Scientifique.frx":0000
            Top             =   1680
            Width           =   555
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   1
            Left            =   120
            Picture         =   "Scientifique.frx":0410
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   0
            Left            =   120
            Picture         =   "Scientifique.frx":07DC
            Top             =   480
            Width           =   390
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
            TabIndex        =   110
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   3
            Left            =   120
            Picture         =   "Scientifique.frx":0C1F
            Top             =   2280
            Width           =   420
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
            TabIndex        =   109
            Top             =   120
            Width           =   1935
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
            TabIndex        =   108
            Top             =   3000
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
            TabIndex        =   107
            Top             =   3600
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
            TabIndex        =   106
            Top             =   4200
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
            TabIndex        =   105
            Top             =   4800
            Width           =   2055
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
            TabIndex        =   103
            Top             =   6000
            Width           =   975
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
            TabIndex        =   102
            Top             =   6600
            Width           =   1335
         End
         Begin VB.Image Logos 
            Height          =   525
            Index           =   4
            Left            =   50
            Picture         =   "Scientifique.frx":1048
            Top             =   3480
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   5
            Left            =   120
            Picture         =   "Scientifique.frx":14AB
            Top             =   4080
            Width           =   300
         End
         Begin VB.Image Logos 
            Height          =   495
            Index           =   6
            Left            =   0
            Picture         =   "Scientifique.frx":1865
            Stretch         =   -1  'True
            Top             =   5280
            Width           =   450
         End
         Begin VB.Image Logos 
            Height          =   375
            Index           =   7
            Left            =   80
            Picture         =   "Scientifique.frx":1C08
            Top             =   4800
            Width           =   330
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   8
            Left            =   0
            Picture         =   "Scientifique.frx":1FF6
            Top             =   5880
            Width           =   480
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   9
            Left            =   0
            Picture         =   "Scientifique.frx":23DC
            Top             =   6480
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   570
            Index           =   10
            Left            =   0
            Picture         =   "Scientifique.frx":277F
            Top             =   8280
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   420
            Index           =   11
            Left            =   0
            Picture         =   "Scientifique.frx":2B6A
            Top             =   7080
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   420
            Index           =   12
            Left            =   0
            Picture         =   "Scientifique.frx":2F42
            Top             =   7680
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   13
            Left            =   0
            Picture         =   "Scientifique.frx":3376
            Top             =   9000
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   14
            Left            =   0
            Picture         =   "Scientifique.frx":3741
            Top             =   9600
            Width           =   480
         End
         Begin VB.Image Logos 
            Height          =   375
            Index           =   15
            Left            =   0
            Picture         =   "Scientifique.frx":3B31
            Stretch         =   -1  'True
            Top             =   10320
            Width           =   375
         End
      End
   End
   Begin VB.Frame fonct 
      BackColor       =   &H00353535&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   2160
      TabIndex        =   74
      Top             =   2880
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton ceill 
         BackColor       =   &H00353535&
         Caption         =   "_|x|_"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton floorr 
         BackColor       =   &H00353535&
         Caption         =   "|_x_|"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton absol 
         BackColor       =   &H00353535&
         Caption         =   "|x|"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton rand 
         BackColor       =   &H00353535&
         Caption         =   "rand"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4920
      ScaleHeight     =   135
      ScaleWidth      =   3735
      TabIndex        =   69
      Top             =   6600
      Width           =   3735
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   8400
      ScaleHeight     =   5415
      ScaleWidth      =   135
      TabIndex        =   68
      Top             =   1440
      Width           =   135
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   4800
      ScaleHeight     =   5415
      ScaleWidth      =   135
      TabIndex        =   70
      Top             =   1440
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4800
      ScaleHeight     =   135
      ScaleWidth      =   3735
      TabIndex        =   71
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox Panneau1 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00353535&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   240
      TabIndex        =   53
      Top             =   2880
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cotanhmoins 
         BackColor       =   &H00353535&
         Caption         =   "coth-1"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   98
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
         TabIndex        =   97
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
         TabIndex        =   96
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton tanhmoins 
         BackColor       =   &H00353535&
         Caption         =   "tanh-1"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   79
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton hyp 
         BackColor       =   &H00353535&
         Caption         =   "hyp"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton secondeTrigo 
         BackColor       =   &H00353535&
         Caption         =   "2nd"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton secc 
         BackColor       =   &H00353535&
         Caption         =   "sec"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton csc 
         BackColor       =   &H00353535&
         Caption         =   "csc"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cotan 
         BackColor       =   &H00353535&
         Caption         =   "cot"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cos 
         BackColor       =   &H00353535&
         Caption         =   "cos"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton tan 
         BackColor       =   &H00353535&
         Caption         =   "tan"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton sin 
         BackColor       =   &H00353535&
         Caption         =   "sin"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cosMoins 
         BackColor       =   &H00353535&
         Caption         =   "cos-1"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton tanMoins 
         BackColor       =   &H00353535&
         Caption         =   "tan-1"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton seccMoins 
         BackColor       =   &H00353535&
         Caption         =   "sec-1"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cscMoins 
         BackColor       =   &H00353535&
         Caption         =   "csc-1"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cotanMoins 
         BackColor       =   &H00353535&
         Caption         =   "cot-1"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton sinMoins 
         BackColor       =   &H00353535&
         Caption         =   "sin-1"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CommandButton second 
      BackColor       =   &H00353535&
      Caption         =   "2nd"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3120
      Width           =   855
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00353535&
      Caption         =   "X"
      Height          =   495
      Index           =   3
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4920
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00353535&
      Caption         =   "+"
      Height          =   495
      Index           =   1
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton égale 
      BackColor       =   &H00C0C0FF&
      Caption         =   "="
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton signes 
      BackColor       =   &H00505050&
      Caption         =   "+/-"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton Chiffre 
      BackColor       =   &H00505050&
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton factoriel 
      BackColor       =   &H00353535&
      Caption         =   "n!"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton bracket 
      BackColor       =   &H00353535&
      Caption         =   ")"
      Height          =   495
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton bracket 
      BackColor       =   &H00353535&
      Caption         =   "("
      Height          =   495
      Index           =   0
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton operationMod 
      BackColor       =   &H00353535&
      Caption         =   "mod"
      Height          =   495
      Index           =   0
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3720
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton ex 
      BackColor       =   &H00353535&
      Caption         =   "exp"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Module 
      BackColor       =   &H00353535&
      Caption         =   "|x|"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton unSurX 
      BackColor       =   &H00353535&
      Caption         =   "1/x"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton e 
      BackColor       =   &H00353535&
      Caption         =   "e"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   735
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton sec 
      BackColor       =   &H00C0C0FF&
      Caption         =   "2nd"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Panneau 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton cmdCE 
      BackColor       =   &H00353535&
      Caption         =   "CE"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton puissanceTrois 
      BackColor       =   &H00353535&
      Caption         =   "x^3"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton racineCubique 
      BackColor       =   &H00353535&
      Caption         =   "3^Vx"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton racineSpec 
      BackColor       =   &H00353535&
      Caption         =   "y^Vx"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton deuxPuissance 
      BackColor       =   &H00353535&
      Caption         =   "2^x"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton logyX 
      BackColor       =   &H00353535&
      Caption         =   "logy X"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton expo 
      BackColor       =   &H00353535&
      Caption         =   "e^x"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton puissance 
      BackColor       =   &H00353535&
      Caption         =   "x²"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton racineCarré 
      BackColor       =   &H00353535&
      Caption         =   "²Vx"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton puissanceSpec 
      BackColor       =   &H00353535&
      Caption         =   "x^y"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton dixPuissance 
      BackColor       =   &H00353535&
      Caption         =   "10^x"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton logDix 
      BackColor       =   &H00353535&
      Caption         =   "log"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Logarithme 
      BackColor       =   &H00353535&
      Caption         =   "ln"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton clear 
      BackColor       =   &H00353535&
      Caption         =   "C"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3120
      Width           =   735
   End
   Begin VB.ListBox historique 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
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
      Height          =   5130
      ItemData        =   "Scientifique.frx":3F07
      Left            =   4920
      List            =   "Scientifique.frx":3F0E
      TabIndex        =   43
      Top             =   1560
      Width           =   3495
   End
   Begin VB.ListBox memoire 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
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
      Height          =   5130
      ItemData        =   "Scientifique.frx":3F33
      Left            =   4920
      List            =   "Scientifique.frx":3F3A
      TabIndex        =   42
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00202020&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      TabIndex        =   66
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
      Left            =   1920
      TabIndex        =   73
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Gradiant 
      BackColor       =   &H00202020&
      Caption         =   "GRAD"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   72
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00202020&
      Caption         =   "Scientifique"
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
      TabIndex        =   67
      Top             =   120
      Width           =   1890
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   476.053
      X2              =   1190.131
      Y1              =   283.744
      Y2              =   283.744
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   476.053
      X2              =   1190.131
      Y1              =   425.616
      Y2              =   425.616
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   476.053
      X2              =   1190.131
      Y1              =   567.488
      Y2              =   567.488
   End
   Begin VB.Label radiant 
      BackColor       =   &H00202020&
      Caption         =   "RAD"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   65
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0FF&
      Visible         =   0   'False
      X1              =   13091.44
      X2              =   14043.55
      Y1              =   1418.719
      Y2              =   1418.719
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0FF&
      X1              =   10473.16
      X2              =   11425.26
      Y1              =   1418.719
      Y2              =   1418.719
   End
   Begin VB.Label cmdMemoire 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      BackStyle       =   0  'Transparent
      Caption         =   "Mémoire"
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
      Left            =   6240
      TabIndex        =   64
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label cmdHistorique 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      BackStyle       =   0  'Transparent
      Caption         =   "Historique"
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
      Left            =   4920
      TabIndex        =   63
      Top             =   840
      Width           =   1215
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
      Left            =   240
      TabIndex        =   62
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label cmdMS 
      BackColor       =   &H00202020&
      Caption         =   "MS"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label memMmoins 
      BackColor       =   &H00202020&
      Caption         =   "M-"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label memMplus 
      BackColor       =   &H00202020&
      Caption         =   "M+"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label cmdMR 
      BackColor       =   &H00202020&
      Caption         =   "MR"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label cmdMC 
      BackColor       =   &H00202020&
      Caption         =   "MC"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Degré 
      BackColor       =   &H00202020&
      Caption         =   "DEG"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Delhistorique 
      Height          =   495
      Left            =   8040
      Picture         =   "Scientifique.frx":3F53
      Stretch         =   -1  'True
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Delmemoire 
      Height          =   495
      Left            =   8040
      Picture         =   "Scientifique.frx":439C
      Stretch         =   -1  'True
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Scient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m(5), tabPanneau(10000)
Dim X, Y, numTabPan
Dim vir, egale, brack, operateur, parenthese, activeExc, initial
Dim modulo, yroot, logBase, puisSpec
Dim memo, hist
Dim mem, tmp
Dim deg, rad, grad
Dim excell, positionPar, panMod
Dim panneauSansPar
Dim AppExcel As Excel.Application
Dim wbExcel As Excel.Workbook
Dim wsExcel As Excel.Worksheet
Dim segundo, segundo1
Dim chiff, oper, ix, brac

Private Sub absol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
absol.BackColor = &H808080
End Sub

Private Sub bracket_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
bracket(brac).BackColor = &H353535
bracket(o).BackColor = &H808080
brac = o
End Sub

Private Sub ceill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ceill.BackColor = &H808080
End Sub

Private Sub Chiffre_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
chiffre(chiff).BackColor = &H505050
chiffre(o).BackColor = &H808080
chiff = o
End Sub

Private Sub clear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clear.BackColor = &H808080
End Sub

Private Sub cmdHistorique_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMemoire.ForeColor = &HFFFFFF
cmdHistorique.ForeColor = &H808080
End Sub

Private Sub cmdMC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdMC.ForeColor = &HFFFFFF Then cmdMC.BackColor = &H808080
End Sub



Private Sub cmdMemoire_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdHistorique.ForeColor = &HFFFFFF
cmdMemoire.ForeColor = &H808080
End Sub

Private Sub cmdMR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdMR.ForeColor = &HFFFFFF Then cmdMR.BackColor = &H808080
End Sub

Private Sub cmdMS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMS.BackColor = &H808080
End Sub

Private Sub cos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cos.BackColor = &H808080
End Sub

Private Sub cosh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cosh.BackColor = &H808080
End Sub

Private Sub coshmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
coshmoins.BackColor = &H808080
End Sub

Private Sub cosMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cosMoins.BackColor = &H808080
End Sub

Private Sub cotan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cotan.BackColor = &H808080
End Sub

Private Sub cotanh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cotanh.BackColor = &H808080
End Sub

Private Sub cotanhmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cotanhmoins.BackColor = &H808080
End Sub

Private Sub cotanMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cotanMoins.BackColor = &H808080
End Sub

Private Sub csc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
csc.BackColor = &H808080
End Sub

Private Sub csch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
csch.BackColor = &H808080
End Sub

Private Sub cschmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cschmoins.BackColor = &H808080
End Sub

Private Sub cscMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cscMoins.BackColor = &H808080
End Sub

Private Sub Degré_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Degré.BackColor = &H808080
End Sub

Private Sub deuxPuissance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
deuxPuissance.BackColor = &H808080
End Sub

Private Sub dixPuissance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
dixPuissance.BackColor = &H808080
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

Private Sub ex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ex.BackColor = &H808080
End Sub

Private Sub expo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
expo.BackColor = &H808080
End Sub

Private Sub factoriel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
factoriel.BackColor = &H808080
End Sub

Private Sub floorr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
floorr.BackColor = &H808080
End Sub

Private Sub Fonc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
trigo.BackColor = &H202020
Fonc.BackColor = &H808080
End Sub

Private Sub fonct_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rand.BackColor = &H353535
ceill.BackColor = &H353535
floorr.BackColor = &H353535
absol.BackColor = &H353535
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

Private Sub Gradiant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Gradiant.BackColor = &H808080
End Sub

Private Sub historique_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMemoire.ForeColor = &HFFFFFF
cmdHistorique.ForeColor = &HFFFFFF
Degré.BackColor = &H202020
radiant.BackColor = &H202020
Gradiant.BackColor = &H202020
cmdMS.BackColor = &H202020
memMmoins.BackColor = &H202020
memMplus.BackColor = &H202020
cmdMR.BackColor = &H202020
cmdMC.BackColor = &H202020
operationMod(0).BackColor = &H353535
factoriel.BackColor = &H353535
ex.BackColor = &H353535
rand.BackColor = &H353535
ceill.BackColor = &H353535
floorr.BackColor = &H353535
logyX.BackColor = &H353535
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

Private Sub hyp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
hyp.BackColor = &H808080
End Sub

Private Sub hyp1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
hyp1.BackColor = &H8080FF
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 0
End Sub

Private Sub Logarithme_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Logarithme.BackColor = &H808080
End Sub

Private Sub logDix_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
logDix.BackColor = &H808080
End Sub

Private Sub logyX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
logyX.BackColor = &H808080
End Sub

Private Sub memMmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
memMmoins.BackColor = &H808080
End Sub

Private Sub memMplus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
memMplus.BackColor = &H808080
End Sub

Private Sub memoire_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMemoire.ForeColor = &HFFFFFF
cmdHistorique.ForeColor = &HFFFFFF
Degré.BackColor = &H202020
radiant.BackColor = &H202020
Gradiant.BackColor = &H202020
cmdMS.BackColor = &H202020
memMmoins.BackColor = &H202020
memMplus.BackColor = &H202020
cmdMR.BackColor = &H202020
cmdMC.BackColor = &H202020
operationMod(0).BackColor = &H353535
factoriel.BackColor = &H353535
ex.BackColor = &H353535
rand.BackColor = &H353535
ceill.BackColor = &H353535
floorr.BackColor = &H353535
logyX.BackColor = &H353535
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

Private Sub Module_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Module.BackColor = &H808080
End Sub

Private Sub operation_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
operation(oper).BackColor = &H353535
operation(o).BackColor = &H808080
oper = o
End Sub

Private Sub operationMod_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
operationMod(0).BackColor = &H808080
End Sub

Private Sub Panneau1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 0
End Sub

Private Sub Pi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pi.BackColor = &H808080
End Sub

Private Sub puissance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
puissance.BackColor = &H808080
End Sub

Private Sub puissanceSpec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
puissanceSpec.BackColor = &H808080
End Sub

Private Sub puissanceTrois_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
puissanceTrois.BackColor = &H808080
End Sub

Private Sub racineCarré_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
racineCarré.BackColor = &H808080
End Sub

Private Sub racineCubique_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
racineCubique.BackColor = &H808080
End Sub

Private Sub racineSpec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
racineSpec.BackColor = &H808080
End Sub



Private Sub radiant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
radiant.BackColor = &H808080
End Sub

Private Sub rand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rand.BackColor = &H808080
End Sub

Private Sub sec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sec.BackColor = &H8080FF
End Sub

Private Sub secc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secc.BackColor = &H808080
End Sub

Private Sub secch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secch.BackColor = &H808080
End Sub

Private Sub secchmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secchmoins.BackColor = &H808080
End Sub


Private Sub seccMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
seccMoins.BackColor = &H808080
End Sub

Private Sub second_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
second.BackColor = &H808080
End Sub

Private Sub secondeTrigo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secondeTrigo.BackColor = &H808080
End Sub

Private Sub secondeTrigo1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secondeTrigo1.BackColor = &H8080FF
End Sub

Private Sub signes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
signes.BackColor = &H808080
End Sub

Private Sub sin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sin.BackColor = &H808080
End Sub

Private Sub sinh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sinh.BackColor = &H808080
End Sub

Private Sub sinhmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sinhmoins.BackColor = &H808080
End Sub

Private Sub sinMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sinMoins.BackColor = &H808080
End Sub

Private Sub tan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tan.BackColor = &H808080
End Sub

Private Sub tanh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tanh.BackColor = &H808080
End Sub

Private Sub tanhmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tanhmoins.BackColor = &H808080
End Sub

Private Sub tanMoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tanMoins.BackColor = &H808080
End Sub

Private Sub trigo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Fonc.BackColor = &H202020
trigo.BackColor = &H808080
End Sub

Private Sub unSurX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
unSurX.BackColor = &H808080
End Sub

Private Sub virgule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Virgule.BackColor = &H808080
End Sub






Private Sub absol_Click()
Panneau1 = "abs" & "(" & Panneau & ")"
Panneau = Abs(Panneau)
End Sub

Private Sub Ang_Click(Index As Integer)
Unload Me
Angle.Show
End Sub

Private Sub bracket_Click(b As Integer)
numTabPan = numTabPan + 1
If b = 0 Then
    parenthese = True
    positionPar = numTabPan
    Panneau1 = Panneau1 & "("
    If Left(Panneau1, Len(Panneau1) - (Len(Panneau1) - 1)) = 0 Then
        Panneau1 = "("
    End If
    brack = brack + 1
End If
If b = 1 And brack > 0 Then
    parenthese = False
    Panneau1 = Panneau1 & ")"
    If Panneau <> 0 Then
    Panneau1 = Replace(Panneau1, ")", Panneau)
    Panneau1 = Panneau1 & ")"
    excell = "=" & panneauSansPar & Panneau
    Worksheets("Feuil1").Cells(1, 1).Value = excell
    Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
    End If
    If Mid(Panneau1, Len(Panneau1) - 1, 1) = "(" Then
    Panneau1 = Left(Panneau1, Len(Panneau1) - 1)
    Panneau1 = Panneau1 & "0" & ")"
    End If
brack = brack - 1
End If
End Sub

Private Sub calculDeLaDate_Click()
Stand.Hide
Scient.Hide
Graph.Hide
calculDate.Show
End Sub

Private Sub CalculDedate_Click(Index As Integer)
Unload Me
CalculdelaDate.Show
End Sub

Private Sub ceill_Click()
Dim ceil() As String
Panneau1 = "floor" & "(" & Panneau & ")"
ceil = Split(Panneau, ",")
Panneau = Val(ceil(0)) + 1
End Sub

Public Sub chiffre_Click(i As Integer)
If FrameSupport.Left > -8000 Then
Do While FrameSupport.Left > -8000
FrameSupport.Left = FrameSupport.Left - 1
Loop
End If
cmdCE.Visible = True
If egale = True Then Panneau1 = ""
egale = False
numTabPan = numTabPan + 1
tabPanneau(numTabPan) = i
If numTabPan = 1 Or operateur = True Then
    Panneau = i
    operateur = False
Else
Panneau = Panneau & tabPanneau(numTabPan)
End If
End Sub

Private Sub clear_Click()
Panneau = 0
Panneau1 = ""
X = 0
panMod = 0
vir = True
vir = True
operateur = False
egale = False
initial = True
modulo = False
End Sub

Private Sub cmdCE_Click()
numTabPan = 0
If egale = True Then
Panneau1 = ""
Panneau = 0
Else
Panneau = 0
End If
X = 0
panMod = 0
modulo = False
vir = True
operateur = False
egale = False
cmdCE.Visible = False
initial = True
End Sub

Private Sub cmdHistorique_Click()
If hist = True Then Delhistorique.Visible = True
    Delmemoire.Visible = False
    memoire.Visible = False
    historique.Visible = True
    Line1.Visible = True
    Line2.Visible = False
End Sub

Private Sub cmdMC_Click()
memoire.clear
Delmemoire.Visible = False
memoire.AddItem ("La mémoire est vide")
cmdMR.ForeColor = &H808080
cmdMC.ForeColor = &H808080
End Sub

Private Sub cmdMemoire_Click()
   If memo = True Then Delmemoire.Visible = True
   Delhistorique.Visible = False
    memoire.Visible = True
    historique.Visible = False
    Line2.Visible = True
    Line1.Visible = False
End Sub

Private Sub cmdMR_Click()
Panneau.Text = tmp
End Sub

Private Sub cmdMS_Click()
memoire.clear
memoire.AddItem (Panneau)
tmp = Panneau
If memoire.Visible = True Then Delmemoire.Visible = True
memo = True
cmdMR.ForeColor = &HFFFFFF
cmdMC.ForeColor = &HFFFFFF
End Sub

Private Sub cosh_Click()
Panneau1 = "cosh" & "(" & Panneau & ")" & "="
excell = "=" & "COSH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)

End Sub

Private Sub coshmoins_Click()
Panneau1 = "cosh-1" & "(" & Panneau & ")" & "="
excell = "=" & "ACOSH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub cosMoins_Click()
Panneau1 = "cos-1" & "(" & Panneau & ")" & "="
If rad = True Then
excell = "=" & "ACOS" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
excell = "=" & "ACOS" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
excell = "=" & "ACOS" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
End If
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub cotan_Click()
Panneau1 = "tan" & "(" & Panneau & ")" & "="
If rad = True Then
Panneau = 1 / Math.tan(Panneau)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
Panneau = 1 / Math.tan(Panneau)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
Panneau = 1 / Math.tan(Panneau)
End If
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub cotanh_Click()
Panneau1 = "cotanh" & "(" & Panneau & ")" & "="
excell = "=" & 1 & "/" & "TANH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub cotanhmoins_Click()
Panneau1 = "cotanh-1" & "(" & Panneau & ")" & "="
excell = "=" & 1 & "/" & "ATANH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub cotanMoins_Click()
Panneau1 = "cot-1" & "(" & Panneau & ")" & "="
If rad = True Then
excell = "=" & 1 & "/" & "ATAN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
excell = "=" & 1 & "/" & "ATAN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
excell = "=" & 1 & "/" & "ATAN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
End If
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub csc_Click()
Panneau1 = "sin" & "(" & Panneau & ")" & "="
If rad = True Then
Panneau = 1 / Math.sin(Panneau)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
Panneau = 1 / Math.sin(Panneau)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
Panneau = 1 / Math.sin(Panneau)
End If
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub csch_Click()
Panneau1 = "csch" & "(" & Panneau & ")" & "="
excell = "=" & 1 & "/" & "SINH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub cschmoins_Click()
Panneau1 = "csch-1" & "(" & Panneau & ")" & "="
excell = "=" & 1 & "/" & "ASINH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub cscMoins_Click()
Panneau1 = "csc-1" & "(" & Panneau & ")" & "="
If rad = True Then
excell = "=" & 1 & "/" & "ASIN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
excell = "=" & 1 & "/" & "ASIN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
excell = "=" & 1 & "/" & "ASIN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
End If
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)

End Sub

Private Sub degg_Click()
Panneau1 = "deg" & "(" & Panneau & ")"
Panneau = Panneau
End Sub

Private Sub Delmemoire_Click()
memoire.clear
memoire.AddItem ("La mémoire est vide")
cmdMR.ForeColor = &H808080
cmdMC.ForeColor = &H808080
Delmemoire.Visible = False
memo = False
End Sub

Private Sub cos_Click()
Panneau1 = "cos" & "(" & Panneau & ")" & "="
If rad = True Then
Panneau = Math.cos(Panneau)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
Panneau = Math.cos(Panneau)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
Panneau = Math.cos(Panneau)
End If
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub Degré_Click()
deg = False
Degré.Visible = False
rad = True
radiant.Visible = True
End Sub

Private Sub DelHistorique_Click()
historique.clear
historique.AddItem ("Aucun historique pour l'instant")
Delhistorique.Visible = False
hist = False
End Sub

Private Sub deuxPuissance_Click()
Panneau1 = "2" & "^" & "(" & Panneau & ")" & "="
Panneau = 2 ^ (Panneau)
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub dixPuissance_Click()
Panneau1 = "10" & "^" & "(" & Panneau & ")" & "="
Panneau = 10 ^ Panneau
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub dms_Click()
Panneau1 = "dms" & "(" & Panneau & ")"
Panneau = Panneau
End Sub

Private Sub don_Click(Index As Integer)
Unload Me
Données.Show
End Sub

Private Sub e_Click()
Panneau = 2.71828182845905
End Sub

Private Sub effacer_Click()
   If Panneau.Text = "" Or Panneau.Text = "0" Then
     Exit Sub
    ElseIf Len(Panneau.Text) = 1 Then
    Panneau.Text = "0"
    numTabPan = 0
    Else
    Panneau.Text = Left(Panneau.Text, Len(Panneau.Text) - 1)
End If
If Len(Panneau) = 1 And Panneau = 0 Then bout = False
If InStr(1, Panneau, ",") = 0 Then
vir = True
X = Panneau
End If
End Sub

Private Sub égale_Click()
If FrameSupport.Left > -8000 Then
Do While FrameSupport.Left > -8000
FrameSupport.Left = FrameSupport.Left - 1
Loop
End If
X = 0
historique.clear
egale = True
If brack > 0 Then
    excell = Panneau1 & Panneau
    Do While brack > 0
    excell = excell & ")"
    Panneau1 = Panneau1 & Panneau & ")" & "="
    brack = brack - 1
    Loop
    excell = "=" & excell
ElseIf brack = 0 And Right(Panneau1, 1) = ")" Then
    Panneau1 = Panneau1 & "="
    excell = Panneau
    excell = "=" & excell
Else
    If modulo = True Then
        Panneau1 = panMod & "Mod" & Panneau & "="
        Panneau = panMod Mod Panneau
        h = Panneau1.Text & Panneau.Text
        historique.AddItem (h)
        modulo = False
        Exit Sub
    ElseIf yroot = True Then
        Panneau1 = panMod & "yroot" & Panneau & "="
        Panneau = panMod ^ (1 / Panneau)
        h = Panneau1.Text & Panneau.Text
        historique.AddItem (h)
        yroot = False
        Exit Sub
    ElseIf logBase = True Then
        Panneau1 = panMod & "log base" & Panneau & "="
        Panneau = Log(panMod) / Log(Panneau)
        h = Panneau1.Text & Panneau.Text
        historique.AddItem (h)
        logBase = False
        Exit Sub
    ElseIf puisSpec = True Then
        Panneau1 = panMod & "^" & Panneau & "="
        Panneau = panMod ^ Panneau
        h = Panneau1.Text & Panneau.Text
        historique.AddItem (h)
        puisSpec = False
    Else
        excell = "=" & Panneau1 & Panneau
        Panneau1 = Panneau1 & Panneau & "="
    End If
End If
If Panneau = 0 And Panneau1 = "" Or Right(Panneau1, 2) = "0=" Then
    Panneau1 = 0 & "="
    h = Panneau1.Text & Panneau.Text
    historique.AddItem (h)
    Exit Sub
End If
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
mem = Panneau
Delhistorique.Visible = True
hist = True
parenthese = False
End Sub

Private Sub exp_Click()
Panneau = Exp(Panneau)
End Sub

Private Sub energ_Click(Index As Integer)
Unload Me
Energie.Show
End Sub

Private Sub expo_Click()
Panneau1 = "e" & "^" & "(" & Panneau & ")" & "="
Panneau = Exp(Panneau)
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub factoriel_Click()
Panneau1 = "fact" & "(" & Panneau & ")"
p = Panneau
If Panneau > 1 Then
Do While p > 1
Panneau = Panneau * (p - 1)
p = p - 1
Loop
ElseIf Panneau = 0 Or Panneau = 1 Then
Panneau = 1
ElseIf Panneau < 0 Then
Panneau = "Non valide"
End If
End Sub

Private Sub floorr_Click()
Dim floor() As String
Panneau1 = "floor" & "(" & Panneau & ")"
floor = Split(Panneau, ",")
Panneau = floor(0)
End Sub

Private Sub Fonc_Click()
If fonct.Visible = False Then
fonct.Visible = True
Else
fonct.Visible = False
End If
End Sub

Private Sub Form_Load()
oper = 1
deg = True
initial = True
vir = True
parenthese = False
Panneau = 0
tmp = 0
A = 0
X = 0
m(1) = "+"
m(2) = "-"
m(3) = "*"
m(4) = "/"
Delmemoire.Visible = False
Delhistorique.Visible = False
historique.Visible = True
Set AppExcel = CreateObject("Excel.Application")
Set wbExcel = AppExcel.Workbooks.Add
Set wsExcel = wbExcel.Worksheets("Feuil1")
End Sub

Private Sub Form_Click()
If FrameSupport.Left > -8000 Then
Do While FrameSupport.Left > -8000
FrameSupport.Left = FrameSupport.Left - 1
Loop
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 0
cmdMemoire.ForeColor = &HFFFFFF
cmdHistorique.ForeColor = &HFFFFFF
Degré.BackColor = &H202020
radiant.BackColor = &H202020
Gradiant.BackColor = &H202020
cmdMS.BackColor = &H202020
memMmoins.BackColor = &H202020
memMplus.BackColor = &H202020
cmdMR.BackColor = &H202020
cmdMC.BackColor = &H202020
operationMod(0).BackColor = &H353535
factoriel.BackColor = &H353535
ex.BackColor = &H353535
rand.BackColor = &H353535
ceill.BackColor = &H353535
floorr.BackColor = &H353535
logyX.BackColor = &H353535
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

Private Sub Gradiant_Click()
grad = False
Gradiant.Visible = False
deg = True
Degré.Visible = True
End Sub



Private Sub Graphique_Click(Index As Integer)
Unload Me
Graph.Show
End Sub

Private Sub heur_Click(Index As Integer)
Unload Me
Heure.Show
End Sub

Private Sub historique_Click()
If FrameSupport.Left > -8000 Then
Do While FrameSupport.Left > -8000
FrameSupport.Left = FrameSupport.Left - 1
Loop
End If
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

Private Sub Label1_Click()
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

Private Sub Label3_Click()

End Sub

Private Sub Logarithme_Click()
Panneau1 = "Ln" & "(" & Panneau & ")" & "="
Panneau = Log(Panneau)
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub logDix_Click()
Panneau1 = "Log" & "(" & Panneau & ")" & "="
Panneau = Log(Panneau) / Log(10)
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub logyX_Click()
Panneau1 = Panneau & "log base"
logBase = True
operateur = True
panMod = Panneau
End Sub

Private Sub Longueur_Click(Index As Integer)
Unload Me
Longue.Show
End Sub

Private Sub memMmoins_Click()
Dim e As Integer
e = Panneau.Text
memoire.clear
tmp = tmp - e
memoire.AddItem (tmp)
If memoire.Visible = True Then Delmemoire.Visible = True
memo = True
cmdMR.ForeColor = &HFFFFFF
cmdMC.ForeColor = &HFFFFFF
End Sub

Private Sub memMplus_Click()
Dim e As Integer
e = Panneau.Text
memoire.clear
tmp = tmp + e
memoire.AddItem (tmp)
cmdMR.ForeColor = &HFFFFFF
cmdMC.ForeColor = &HFFFFFF
If memoire.Visible = True Then Delmemoire.Visible = True
memo = True
End Sub

Private Sub Module_Click()
Panneau1 = "abs" & "(" & Panneau & ")"
Panneau = Abs(Panneau)
End Sub

Private Sub Navigation_Click()
If FrameMenu.Left < 0 Then
Do While FrameMenu.Left < 0
FrameMenu.Left = FrameMenu.Left + 10
Loop
Else
Do While FrameMenu.Left > -5700
FrameMenu.Left = FrameMenu.Left - 10
Loop
End If
End Sub

Private Sub operation_Click(k As Integer)
If FrameSupport.Left > -8000 Then
Do While FrameSupport.Left > -8000
FrameSupport.Left = FrameSupport.Left - 1
Loop
End If
egale = False
vir = True
operateur = True
X = 0
A = numTabPan
If initial = True Then
    Panneau1 = Panneau & m(k)
    panneauSansPar = Panneau1
    initial = False
    If parenthese = True Then
        panneauSansPar = Panneau1
        Panneau1 = "(" & Panneau1
    End If
ElseIf initial = False Then
    excell = "=" & panneauSansPar & Panneau
    Panneau1 = Panneau1 & Panneau & m(k)
    panneauSansPar = panneauSansPar & Panneau & m(k)
    Worksheets("Feuil1").Cells(1, 1).Value = excell
    Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
End If
End Sub

Private Sub operationMod_Click(Index As Integer)
modulo = True
panMod = Panneau
Panneau1 = Panneau1 & Panneau & "Mod"
operateur = True
End Sub

Private Sub Pi_Click()
Panneau = 3.14159265358979
End Sub

Private Sub PoidsEtMasse_Click(Index As Integer)
Unload Me
PoidsMasse.Show
End Sub

Private Sub press_Click(Index As Integer)
Unload Me
Pression.Show
End Sub

Private Sub puissance_Click()
Panneau1 = "sqr" & "(" & Panneau & ")" & "="
Panneau = Panneau ^ 2
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub puissanceSpec_Click()
Panneau1 = Panneau & "^"
panMod = Panneau
puisSpec = True
End Sub

Private Sub puissanceTrois_Click()
Panneau1 = "cube" & "(" & Panneau & ")" & "="
Panneau = Panneau ^ 3
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub racineCarré_Click()
Panneau1 = "V" & "(" & Panneau & ")" & "="
Panneau = Sqr(Panneau)
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub racineCubique_Click()
Panneau1 = "cuberoot" & "(" & Panneau & ")" & "="
Panneau = Panneau ^ (1 / 3)
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub racineSpec_Click()
Panneau1 = Panneau & "yroot"
panMod = Panneau
operateur = True
yroot = True
End Sub

Private Sub radiant_Click()
rad = False
radiant.Visible = False
grad = True
Gradiant.Visible = True
End Sub

Private Sub rand_Click()
Panneau = Rnd()
End Sub


Private Sub signe_Click()
Panneau = -Panneau
End Sub

Private Sub scshmoins_Click()

End Sub

Private Sub sec_Click()
sec.Visible = False
second.Visible = True
puissanceTrois.Visible = False
racineCubique.Visible = False
racineSpec.Visible = False
deuxPuissance.Visible = False
logyX.Visible = False
expo.Visible = False
End Sub

Private Sub secc_Click()
Panneau1 = "cos" & "(" & Panneau & ")" & "="
If rad = True Then
Panneau = 1 / Math.cos(Panneau)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
Panneau = 1 / Math.cos(Panneau)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
Panneau = 1 / Math.cos(Panneau)
End If
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub secch_Click()
Panneau1 = "sech" & "(" & Panneau & ")" & "="
excell = "=" & 1 & "/" & "COSH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub secchmoins_Click()
Panneau1 = "sech-1" & "(" & Panneau & ")" & "="
excell = "=" & 1 & "/" & "ACOSH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub seccMoins_Click()
Panneau1 = "sec-1" & "(" & Panneau & ")" & "="
If rad = True Then
excell = "=" & 1 & "/" & "ACOS" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
excell = "=" & 1 & "/" & "ACOS" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
excell = "=" & 1 & "/" & "ACOS" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
End If
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub second_Click()
second.Visible = False
sec.Visible = True
puissanceTrois.Visible = True
racineCubique.Visible = True
racineSpec.Visible = True
deuxPuissance.Visible = True
logyX.Visible = True
expo.Visible = True
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

Private Sub signes_Click()
Panneau = -Panneau
End Sub

Private Sub sin_Click()
Panneau1 = "sin" & "(" & Panneau & ")" & "="
If rad = True Then
Panneau = Math.sin(Panneau)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
Panneau = Math.sin(Panneau)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
Panneau = Math.sin(Panneau)
End If
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub sinh_Click()
Panneau1 = "sinh" & "(" & Panneau & ")" & "="
excell = "=" & "SINH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)

End Sub

Private Sub sinhmoins_Click()
Panneau1 = "sinh-1" & "(" & Panneau & ")" & "="
excell = "=" & "ASINH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub sinMoins_Click()
Panneau1 = "sin-1" & "(" & Panneau & ")" & "="
If rad = True Then
excell = "=" & "ASIN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
excell = "=" & "ASIN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
excell = "=" & "ASIN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
End If
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub



Private Sub Standard_Click(Index As Integer)
Unload Me
Stand.Show
End Sub

Private Sub surf_Click(Index As Integer)
Unload Me
Surface.Show
End Sub

Private Sub tan_Click()
Panneau1 = "tan" & "(" & Panneau & ")" & "="
If rad = True Then
Panneau = Math.tan(Panneau)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
Panneau = Math.tan(Panneau)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
Panneau = Math.tan(Panneau)
End If
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub tanh_Click()
Panneau1 = "tanh" & "(" & Panneau & ")" & "="
excell = "=" & "TANH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)

End Sub

Private Sub tanhmoins_Click()
Panneau1 = "tanh-1" & "(" & Panneau & ")" & "="
excell = "=" & "ATANH" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
End Sub

Private Sub tanMoins_Click()
Panneau1 = "tan-1" & "(" & Panneau & ")" & "="
If rad = True Then
excell = "=" & "ATAN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf deg = True Then
Panneau = (Panneau * 3.14159265358979) / 180
excell = "=" & "ATAN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
ElseIf grad = True Then
Panneau = (Panneau * 3.14159265358979) / 200
excell = "=" & "ATAN" & "(" & Panneau & ")"
excell = Replace(excell, ",", ".")
Worksheets("Feuil1").Cells(1, 1).Value = excell
Panneau = CDbl(Worksheets("Feuil1").Cells(1, 1).Value)
End If
If Delhistorique.Visible = False Then historique.clear
If historique.Visible = True Then Delhistorique.Visible = True
hist = True
If Delhistorique.Visible = False Then historique.clear
h = Panneau1.Text & Panneau.Text
historique.AddItem (h)
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

Private Sub unSurX_Click()
If Panneau <> 0 Then
Panneau1 = 1 & "/" & Panneau
Panneau = 1 / Panneau
ElseIf Panneau = 0 Then
Panneau1 = 1 & "/" & Panneau
Panneau = "Erreur"
End If
End Sub

Private Sub Virgule_Click()
numTabPan = numTabPan + 1
tabPanneau(numTabPan) = ","
If vir = True Then
If Panneau <> 0 Then
Panneau = Panneau & ","
vir = False
ElseIf Panneau = 0 Then
Panneau = Panneau & ","
vir = False
End If
End If
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
