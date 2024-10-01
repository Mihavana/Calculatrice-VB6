VERSION 5.00
Begin VB.Form Stand 
   BackColor       =   &H00202020&
   Caption         =   "Calculatrice"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9045
   ClipControls    =   0   'False
   DrawMode        =   8  'Xor Pen
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameSupport 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   -8000
      TabIndex        =   41
      Top             =   720
      Width           =   2895
      Begin VB.VScrollBar VScroll1 
         Height          =   7215
         LargeChange     =   300
         Left            =   2640
         Max             =   4500
         SmallChange     =   300
         TabIndex        =   61
         Top             =   0
         Width           =   255
      End
      Begin VB.Frame FrameMenu 
         BackColor       =   &H00252525&
         BorderStyle     =   0  'None
         Caption         =   "cc"
         Height          =   12015
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   2655
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            Height          =   735
            Left            =   840
            Shape           =   3  'Circle
            Top             =   10800
            Width           =   615
         End
         Begin VB.Image Logos 
            Height          =   375
            Index           =   15
            Left            =   0
            Picture         =   "Standard.frx":0000
            Stretch         =   -1  'True
            Top             =   10320
            Width           =   375
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   14
            Left            =   0
            Picture         =   "Standard.frx":03D6
            Top             =   9600
            Width           =   480
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   13
            Left            =   0
            Picture         =   "Standard.frx":07C6
            Top             =   9000
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   420
            Index           =   12
            Left            =   0
            Picture         =   "Standard.frx":0B91
            Top             =   7680
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   420
            Index           =   11
            Left            =   0
            Picture         =   "Standard.frx":0FC5
            Top             =   7080
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   570
            Index           =   10
            Left            =   0
            Picture         =   "Standard.frx":139D
            Top             =   8280
            Width           =   405
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   9
            Left            =   0
            Picture         =   "Standard.frx":1788
            Top             =   6480
            Width           =   465
         End
         Begin VB.Image Logos 
            Height          =   480
            Index           =   8
            Left            =   0
            Picture         =   "Standard.frx":1B2B
            Top             =   5880
            Width           =   480
         End
         Begin VB.Image Logos 
            Height          =   375
            Index           =   7
            Left            =   80
            Picture         =   "Standard.frx":1F11
            Top             =   4800
            Width           =   330
         End
         Begin VB.Image Logos 
            Height          =   495
            Index           =   6
            Left            =   0
            Picture         =   "Standard.frx":22FF
            Stretch         =   -1  'True
            Top             =   5280
            Width           =   450
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   5
            Left            =   120
            Picture         =   "Standard.frx":26A2
            Top             =   4080
            Width           =   300
         End
         Begin VB.Image Logos 
            Height          =   525
            Index           =   4
            Left            =   50
            Picture         =   "Standard.frx":2A5C
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
            TabIndex        =   60
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
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
            Left            =   720
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
            Top             =   120
            Width           =   1935
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   3
            Left            =   120
            Picture         =   "Standard.frx":2EBF
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
            TabIndex        =   52
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   0
            Left            =   120
            Picture         =   "Standard.frx":32E8
            Top             =   480
            Width           =   390
         End
         Begin VB.Image Logos 
            Height          =   450
            Index           =   1
            Left            =   120
            Picture         =   "Standard.frx":372B
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image Logos 
            Height          =   465
            Index           =   2
            Left            =   0
            Picture         =   "Standard.frx":3AF7
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            Height          =   495
            Index           =   15
            Left            =   600
            TabIndex        =   43
            Top             =   10320
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   615
            Left            =   720
            TabIndex        =   62
            Top             =   10920
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   5040
      ScaleHeight     =   15
      ScaleWidth      =   3615
      TabIndex        =   40
      Top             =   1080
      Width           =   3615
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   4920
      ScaleHeight     =   5655
      ScaleWidth      =   135
      TabIndex        =   39
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   5040
      ScaleHeight     =   135
      ScaleWidth      =   3495
      TabIndex        =   38
      Top             =   6720
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   8520
      ScaleHeight     =   5775
      ScaleWidth      =   120
      TabIndex        =   37
      Top             =   1080
      Width           =   120
   End
   Begin VB.TextBox Panneau1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   35
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox Panneau 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   24
      Top             =   1560
      Width           =   4575
   End
   Begin VB.CommandButton Bouton3 
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
      Height          =   735
      Index           =   5
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Bouton1 
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
      Height          =   735
      Index           =   4
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Bouton1 
      BackColor       =   &H00353535&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Bouton1 
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
      Height          =   735
      Index           =   2
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Bouton1 
      BackColor       =   &H00353535&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Bouton5 
      BackColor       =   &H00353535&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Bouton5 
      BackColor       =   &H00353535&
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Bouton10 
      BackColor       =   &H00353535&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Bouton8 
      BackColor       =   &H00353535&
      Caption         =   "²Vx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Bouton7 
      BackColor       =   &H00353535&
      Caption         =   "x²"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Bouton6 
      BackColor       =   &H00353535&
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton bouton 
      BackColor       =   &H00505050&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton bouton 
      BackColor       =   &H00505050&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton bouton 
      BackColor       =   &H00505050&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton bouton 
      BackColor       =   &H00505050&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton bouton 
      BackColor       =   &H00505050&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton bouton 
      BackColor       =   &H00505050&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton bouton 
      BackColor       =   &H00505050&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton bouton 
      BackColor       =   &H00505050&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton bouton 
      BackColor       =   &H00505050&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Bouton9 
      BackColor       =   &H00505050&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton bouton 
      BackColor       =   &H00505050&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Bouton4 
      BackColor       =   &H00505050&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Bouton2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   1095
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
      Height          =   5730
      ItemData        =   "Standard.frx":3F07
      Left            =   5040
      List            =   "Standard.frx":3F0E
      TabIndex        =   26
      Top             =   1080
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
      Height          =   5730
      ItemData        =   "Standard.frx":3F33
      Left            =   5040
      List            =   "Standard.frx":3F3A
      TabIndex        =   25
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00202020&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Delhistorique 
      Height          =   495
      Left            =   8160
      Picture         =   "Standard.frx":3F53
      Stretch         =   -1  'True
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Delmemoire 
      Height          =   495
      Left            =   8160
      Picture         =   "Standard.frx":439C
      Stretch         =   -1  'True
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00202020&
      Caption         =   "Standard"
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
      Left            =   720
      TabIndex        =   36
      Top             =   146
      Width           =   1455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   600
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0FF&
      Visible         =   0   'False
      X1              =   6840
      X2              =   7320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0FF&
      X1              =   5640
      X2              =   6120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label cmdMemoire 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
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
      Left            =   6480
      TabIndex        =   33
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label cmdHistorique 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
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
      Left            =   5280
      TabIndex        =   32
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label cmdMS 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      Caption         =   "MS"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label memMmoins 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      Caption         =   "M-"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   31
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label memMplus 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   30
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label cmdMR 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1080
      TabIndex        =   29
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label cmdMC 
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "Stand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim operateurEgale(5), m(5), tabPanneau(10000)
Dim X, Y, A, p, l, numTabPan
Dim q, r, s, t, u, v, w, f, g, h, h2
Dim vir, egale, bout, op, operateur, signe, erreur, memo
Dim operateur1
Dim pan As Double
Dim mem, tmp


Private Sub memMmoins_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
memMmoins.BackColor = &H808080
End Sub

Private Sub memMplus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
memMplus.BackColor = &H808080
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

Function Operations()
    If signe = "+" Then
    Panneau1 = pan + p
    pan = Panneau1
    Panneau = Panneau1
    Panneau1 = Panneau1 & m(l)
    ElseIf signe = "-" Then
    Panneau1 = pan - p
    pan = Panneau1
    Panneau = Panneau1
    Panneau1 = Panneau1 & m(l)
    ElseIf signe = "x" Then
    Panneau1 = pan * p
    pan = Panneau1
    Panneau = Panneau1
    Panneau1 = Panneau1 & m(l)
    ElseIf signe = "÷" Then
    Panneau1 = pan / p
    pan = Panneau1
    Panneau = Panneau1
    Panneau1 = Panneau1 & m(l)
        End If
End Function

Private Sub Ang_Click(Index As Integer)
Unload Me
Angle.Show
End Sub

Private Sub bouton_Click(i As Integer)
If FrameSupport.Left > -8000 Then
Do While FrameSupport.Left > -8000
FrameSupport.Left = FrameSupport.Left - 1
Loop
End If
If egale = True Then Panneau1 = ""
egale = False
numTabPan = numTabPan + 1
tabPanneau(numTabPan) = i
If numTabPan = 1 Or operateur1 = True Then
Panneau = i
operateur1 = False
Else
Panneau = Panneau & tabPanneau(numTabPan)
End If
If operateur = True Then op = True
bout = False

End Sub

Private Sub bouton_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton8(h2).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton10(f).BackColor = &H353535
Bouton5(w).BackColor = &H353535
Bouton9(u).BackColor = &H505050
Bouton3(v).BackColor = &H353535
Bouton4(t).BackColor = &H505050
Bouton2(s).BackColor = &HC0C0FF
Bouton1(r).BackColor = &H353535
bouton(q).BackColor = &H505050
bouton(o).BackColor = &H808080
q = o
End Sub

Private Sub Bouton1_Click(k As Integer)
If FrameSupport.Left > -8000 Then
Do While FrameSupport.Left > -8000
FrameSupport.Left = FrameSupport.Left - 1
Loop
End If
operateur1 = True
egale = False
vir = True
bout = True
X = 0
p = Panneau
If operateur = False Or op = False Then
If Panneau1 <> "" Then
Panneau1.Text = Panneau1 & Panneau.Text & m(k)
Else
Panneau1 = Panneau & m(k)
End If
pan = Panneau
operateurEgale(k) = True
operateur = True
signe = m(k)
ElseIf operateur = True Then
    For j = 1 To 4 Step 1
        operateurEgale(j) = False
    Next j
    operateurEgale(k) = True
    l = k
    If k = 1 Then
       Call Operations
    End If
    If k = 2 Then
        Call Operations
    End If
    If k = 3 Then
        Call Operations
    End If
    If k = 4 Then
       Call Operations
    End If
    signe = m(k)
End If

End Sub


Private Sub Bouton1_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton8(h2).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton6(g).BackColor = &H353535
Bouton5(w).BackColor = &H353535
Bouton3(v).BackColor = &H353535
Bouton9(u).BackColor = &H505050
Bouton2(s).BackColor = &HC0C0FF
Bouton4(t).BackColor = &H505050
bouton(q).BackColor = &H505050
Bouton1(r).BackColor = &H353535
Bouton1(o).BackColor = &H808080
r = o
End Sub




Private Sub Bouton10_Click(Index As Integer)
 Panneau.Text = (mem * Panneau.Text) / 100
 bout = False
End Sub

Private Sub Bouton10_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton8(h2).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton6(g).BackColor = &H353535
Bouton10(f).BackColor = &H353535
Bouton5(w).BackColor = &H353535
bouton(q).BackColor = &H505050
Bouton1(r).BackColor = &H353535
Bouton2(s).BackColor = &HC0C0FF
Bouton4(t).BackColor = &H505050
Bouton9(u).BackColor = &H505050
Bouton3(v).BackColor = &H353535
Bouton10(o).BackColor = &H808080
f = o
End Sub

Private Sub Bouton2_Click(Index As Integer)
If FrameSupport.Left > -8000 Then
Do While FrameSupport.Left > -8000
FrameSupport.Left = FrameSupport.Left - 1
Loop
End If
egale = True
bout = False
If operateur = True Then
    If operateurEgale(1) = True Then
    p = Panneau
    Panneau1 = pan & "+" & p & "="
    Panneau = p + pan
    operateurEgale(1) = False
    ElseIf operateurEgale(2) = True Then
    p = Panneau
    Panneau1 = pan & "-" & p & "="
    Panneau = pan - p
    operateurEgale(2) = False
    ElseIf operateurEgale(3) = True Then
    p = Panneau
    Panneau1 = pan & "x" & p & "="
    Panneau = p * pan
    operateurEgale(3) = False
    ElseIf operateurEgale(4) = True Then
        p = Panneau
        If p = 0 Then
        Panneau = "Nous ne pouvons pas diviser par zéro"
        Panneau1 = pan & "÷"
        Else
        Panneau1 = pan & "÷" & p & "="
        Panneau = pan / p
        End If
    operateurEgale(4) = False
    Else
    Panneau1 = Panneau & "+" & p & "="
    Panneau = Val(Panneau.Text) + p
    End If
ElseIf operateur = False Then
If Panneau1 = "" And Panneau = "0" Then
Panneau1 = 0 & "="
h = Panneau1 & Panneau
ElseIf Panneau1 = "" And Panneau <> "0" Then
Panneau1 = Panneau & "="
h = Panneau1 & Panneau
End If
End If
   If p <> 0 Then h = Panneau1.Text & Panneau.Text
    historique.clear
    historique.AddItem (h)
    mem = Panneau
    Delhistorique.Visible = True
End Sub

Private Sub Bouton2_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton8(h2).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton5(w).BackColor = &H353535
Bouton3(v).BackColor = &H353535
Bouton9(u).BackColor = &H505050
bouton(q).BackColor = &H505050
Bouton1(r).BackColor = &H353535
Bouton4(t).BackColor = &H505050
Bouton2(s).BackColor = &HC0C0FF
Bouton2(o).BackColor = &H8080FF
Bouton10(f).BackColor = &H353535
s = o
End Sub

Private Sub Bouton3_Click(Index As Integer)
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

Private Sub Bouton3_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton8(h2).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton6(g).BackColor = &H353535
Bouton5(w).BackColor = &H353535
Bouton10(f).BackColor = &H353535
bouton(q).BackColor = &H505050
Bouton1(r).BackColor = &H353535
Bouton2(s).BackColor = &HC0C0FF
Bouton4(t).BackColor = &H505050
Bouton9(u).BackColor = &H505050
Bouton3(o).BackColor = &H808080
v = o
End Sub

Private Sub Bouton4_Click(Index As Integer)
numTabPan = numTabPan + 1
tabPanneau(numTabPan) = ","
bout = True
If vir = True Then
If X <> 0 Then
Panneau = X & ","
X = Panneau
vir = False
ElseIf X = 0 Then
Panneau = X & ","
X = Panneau
vir = False
End If
End If
End Sub

Private Sub Bouton4_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton8(h2).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton6(g).BackColor = &H353535
Bouton5(w).BackColor = &H353535
Bouton10(f).BackColor = &H353535
Bouton9(u).BackColor = &H505050
Bouton3(v).BackColor = &H353535
Bouton2(s).BackColor = &HC0C0FF
Bouton1(r).BackColor = &H353535
bouton(q).BackColor = &H505050
Bouton4(o).BackColor = &H808080
t = o
End Sub


Private Sub Bouton5_Click(c As Integer)
If c = 1 Then
Panneau = 0
ElseIf c = 2 Then
Panneau = 0
Panneau1 = ""
End If
X = 0
p = 0
vir = True
bout = False
vir = True
Panneau = 0
operateur = False
egale = False
op = False
tabPanneau(10000) = 0
numTabPan = 0
End Sub

Private Sub Bouton5_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton8(h2).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton6(g).BackColor = &H353535
bouton(q).BackColor = &H505050
Bouton1(r).BackColor = &H353535
Bouton2(s).BackColor = &HC0C0FF
Bouton4(t).BackColor = &H505050
Bouton9(u).BackColor = &H505050
Bouton3(v).BackColor = &H353535
Bouton5(w).BackColor = &H353535
Bouton5(o).BackColor = &H808080
w = o
End Sub

Private Sub Bouton6_Click(Index As Integer)
If Panneau <> 0 Then
Panneau1 = 1 & "/" & Panneau
Panneau = 1 / Panneau
ElseIf Panneau = 0 Then
Panneau1 = 1 & "/" & Panneau
Panneau = "Erreur"
End If
End Sub

Private Sub Bouton6_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton8(h2).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton5(w).BackColor = &H353535
bouton(q).BackColor = &H505050
Bouton1(r).BackColor = &H353535
Bouton2(s).BackColor = &HC0C0FF
Bouton4(t).BackColor = &H505050
Bouton9(u).BackColor = &H505050
Bouton3(v).BackColor = &H353535
Bouton10(f).BackColor = &H353535
Bouton6(o).BackColor = &H808080
g = o

End Sub

Private Sub Bouton7_Click(Index As Integer)
Panneau1 = "sqr" & "(" & Panneau & ")"
Panneau = Panneau ^ 2
End Sub

Private Sub Bouton7_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton8(h2).BackColor = &H353535
Bouton5(w).BackColor = &H353535
bouton(q).BackColor = &H505050
Bouton1(r).BackColor = &H353535
Bouton2(s).BackColor = &HC0C0FF
Bouton4(t).BackColor = &H505050
Bouton9(u).BackColor = &H505050
Bouton3(v).BackColor = &H353535
Bouton10(f).BackColor = &H353535
Bouton6(g).BackColor = &H353535
Bouton7(o).BackColor = &H808080
h = o
End Sub

Private Sub Bouton8_Click(Index As Integer)
Panneau1 = "V" & Panneau
Panneau = Panneau ^ (1 / 2)
End Sub

Private Sub Bouton8_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton5(w).BackColor = &H353535
bouton(q).BackColor = &H505050
Bouton1(r).BackColor = &H353535
Bouton2(s).BackColor = &HC0C0FF
Bouton4(t).BackColor = &H505050
Bouton9(u).BackColor = &H505050
Bouton3(v).BackColor = &H353535
Bouton10(f).BackColor = &H353535
Bouton6(g).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton8(o).BackColor = &H808080
h2 = o
End Sub

Private Sub Bouton9_Click(Index As Integer)
Panneau = -Panneau
End Sub


Private Sub ChangePage_Click()
FrameMenu.Visible = False
End Sub

Private Sub ChangePage2_Click()
FrameMenu.Visible = True
End Sub

Private Sub Bouton9_MouseMove(o As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Bouton8(h2).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton6(g).BackColor = &H353535
Bouton10(f).BackColor = &H353535
Bouton3(v).BackColor = &H353535
Bouton5(w).BackColor = &H353535
bouton(q).BackColor = &H505050
Bouton1(r).BackColor = &H353535
Bouton2(s).BackColor = &HC0C0FF
Bouton4(t).BackColor = &H505050
Bouton9(o).BackColor = &H808080
u = o
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



Private Sub DelHistorique_Click()
historique.clear
historique.AddItem ("Aucun historique pour l'instant")
Delhistorique.Visible = False
hist = False
End Sub

Private Sub Delmemoire_Click()
memoire.clear
memoire.AddItem ("La mémoire est vide")
cmdMR.ForeColor = &H808080
cmdMC.ForeColor = &H808080
Delmemoire.Visible = False
memo = False
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
f = 15
r = 1
w = 1
u = 11
v = 5
g = 12
h = 13
h2 = 14
vir = True
Panneau = 0
tmp = 0
X = 0
p = 0
operateurEgale(1) = False
operateurEgale(2) = False
operateurEgale(3) = False
operateurEgale(4) = False
egale = False
m(1) = "+"
m(2) = "-"
m(3) = "x"
m(4) = "÷"
memoire.Visible = False
historique.Visible = True
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMS.BackColor = &H202020
memMmoins.BackColor = &H202020
memMplus.BackColor = &H202020
cmdMR.BackColor = &H202020
cmdMC.BackColor = &H202020
cmdMemoire.ForeColor = &HFFFFFF
cmdHistorique.ForeColor = &HFFFFFF
Label1.BorderStyle = 0
Bouton5(w).BackColor = &H353535
bouton(q).BackColor = &H505050
Bouton1(r).BackColor = &H353535
Bouton2(s).BackColor = &HC0C0FF
Bouton4(t).BackColor = &H505050
Bouton9(u).BackColor = &H505050
Bouton3(v).BackColor = &H353535
Bouton10(f).BackColor = &H353535
Bouton6(g).BackColor = &H353535
Bouton7(13).BackColor = &H353535
Bouton8(h2).BackColor = &H353535
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


Private Sub Label4_Click()

End Sub



Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 0
End Sub

Private Sub Label3_Click()
copyright.Show
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
Private Sub Text1_Change()

End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub

Private Sub Panneau1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 0
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


Private Sub Stand_Click()

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
