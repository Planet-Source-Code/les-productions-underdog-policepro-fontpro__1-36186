VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPrinc 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Police Pro"
   ClientHeight    =   6030
   ClientLeft      =   1440
   ClientTop       =   630
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   402
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   526
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   873
      ButtonWidth     =   847
      ButtonHeight    =   714
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Info"
            Object.ToolTipText     =   "Info sur la police"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grid"
            Object.ToolTipText     =   "Caractères disponibles"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Compare"
            Object.ToolTipText     =   "Pour comparer deux polices"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Gras"
            Object.ToolTipText     =   "Afficher en caractère gras"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Italique"
            Object.ToolTipText     =   "Afficher en caractère Italique"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Souligne"
            Object.ToolTipText     =   "Afficher en caractère souligné"
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Imprime la sélection"
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprime"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   60
      TabIndex        =   32
      Top             =   5190
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   9870
      Left            =   5490
      Picture         =   "Principal.frx":030A
      ScaleHeight     =   654
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   540
      TabIndex        =   31
      Top             =   930
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.PictureBox P_Clip 
      BorderStyle     =   0  'None
      Height          =   4110
      Index           =   3
      Left            =   30
      ScaleHeight     =   274
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   503
      TabIndex        =   25
      Top             =   885
      Visible         =   0   'False
      Width           =   7545
      Begin ComctlLib.Toolbar Toolbar2 
         Height          =   465
         Left            =   3660
         TabIndex        =   26
         Top             =   1815
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   820
         ButtonWidth     =   847
         ButtonHeight    =   714
         ImageList       =   "ImageList1"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   4
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Gras2"
               Object.Tag             =   ""
               ImageIndex      =   4
               Style           =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Italique2"
               Object.Tag             =   ""
               ImageIndex      =   5
               Style           =   1
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Souligne2"
               Object.Tag             =   ""
               ImageIndex      =   6
               Style           =   1
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtAfficheHaut 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1725
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   30
         Text            =   "Principal.frx":56AD4
         Top             =   45
         Width           =   7425
      End
      Begin VB.TextBox txtAfficheBas 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1725
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   29
         Text            =   "Principal.frx":56AE1
         Top             =   2295
         Width           =   7425
      End
      Begin VB.ComboBox cmbPoliceBas 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   45
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1890
         Width           =   2595
      End
      Begin VB.ComboBox TailleBas 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2685
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1890
         Width           =   855
      End
   End
   Begin ComctlLib.StatusBar stBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   5655
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   10848
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox TailleHaut 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3495
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   570
      Width           =   855
   End
   Begin VB.ComboBox cmbPoliceHaut 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   60
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   570
      Width           =   3420
   End
   Begin VB.ComboBox cmbToutePolice 
      Height          =   315
      Left            =   4575
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   570
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox P_Clip 
      BorderStyle     =   0  'None
      Height          =   4140
      Index           =   2
      Left            =   45
      ScaleHeight     =   4140
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   945
      Visible         =   0   'False
      Width           =   7575
      Begin Threed.SSPanel S_Char 
         Height          =   885
         Left            =   1170
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1561
         _StockProps     =   15
         Caption         =   "A"
         ForeColor       =   8388736
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         BorderWidth     =   2
         BevelInner      =   2
      End
      Begin MSGrid.Grid GridFont 
         Height          =   3990
         Left            =   30
         TabIndex        =   1
         Top             =   45
         Width           =   7470
         _Version        =   65536
         _ExtentX        =   13176
         _ExtentY        =   7038
         _StockProps     =   77
         ForeColor       =   16711680
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Rows            =   16
         Cols            =   14
         HighLight       =   0   'False
         MouseIcon       =   "Principal.frx":56AEE
      End
   End
   Begin VB.PictureBox P_Clip 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   1
      Left            =   75
      ScaleHeight     =   4080
      ScaleWidth      =   7650
      TabIndex        =   8
      Top             =   930
      Visible         =   0   'False
      Width           =   7650
      Begin VB.Line Line10 
         X1              =   3450
         X2              =   4605
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line9 
         X1              =   3450
         X2              =   4605
         Y1              =   1425
         Y2              =   1425
      End
      Begin VB.Line Line8 
         X1              =   3435
         X2              =   4605
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Line Line2 
         X1              =   3435
         X2              =   4605
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         X1              =   3435
         X2              =   4620
         Y1              =   3540
         Y2              =   3540
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Type de police:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   45
         TabIndex        =   22
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label lblScale 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "00"
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   0
         Left            =   4590
         TabIndex        =   20
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label lblScale 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "00"
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   1
         Left            =   4590
         TabIndex        =   19
         Top             =   1335
         Width           =   495
      End
      Begin VB.Label lblScale 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "00"
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   2
         Left            =   4590
         TabIndex        =   18
         Top             =   1995
         Width           =   495
      End
      Begin VB.Label lblScale 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "00"
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   3
         Left            =   4605
         TabIndex        =   17
         Top             =   2655
         Width           =   495
      End
      Begin VB.Label lblScale 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "00"
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   4
         Left            =   4620
         TabIndex        =   16
         Top             =   3435
         Width           =   495
      End
      Begin VB.Label S_ScaleName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(Leading Externe)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   0
         Left            =   5145
         TabIndex        =   15
         Top             =   990
         Width           =   2775
      End
      Begin VB.Label S_ScaleName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(Leading Interne)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   1
         Left            =   5130
         TabIndex        =   14
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label S_ScaleName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(Hauteur)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   2
         Left            =   5145
         TabIndex        =   13
         Top             =   1980
         Width           =   2775
      End
      Begin VB.Label S_ScaleName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(Montée)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   3
         Left            =   5145
         TabIndex        =   12
         Top             =   2685
         Width           =   2775
      End
      Begin VB.Label S_ScaleName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(Descente)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   4
         Left            =   5145
         TabIndex        =   11
         Top             =   3435
         Width           =   2775
      End
      Begin VB.Image I_Scale 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3150
         Left            =   75
         Picture         =   "Principal.frx":56B0A
         Top             =   810
         Width           =   3420
      End
      Begin VB.Label S_Family 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   2070
         TabIndex        =   2
         Top             =   -15
         Width           =   45
      End
      Begin VB.Label S_FamilyTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nom de la Famille:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   -30
         Width           =   2070
      End
      Begin VB.Label S_FontType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   2070
         TabIndex        =   4
         Top             =   255
         Width           =   45
      End
      Begin VB.Image I_ttf 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   7185
         Picture         =   "Principal.frx":6288C
         Top             =   270
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label S_DefaultCharTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caractère par Défaut:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   510
         Width           =   1920
      End
      Begin VB.Label S_DefaultChar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   2070
         TabIndex        =   6
         Top             =   525
         Width           =   45
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7020
      Top             =   315
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Principal.frx":62DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Principal.frx":6345C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Principal.frx":63AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Principal.frx":64198
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Principal.frx":64836
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Principal.frx":64ED4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Principal.frx":65572
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Principal.frx":65C10
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Principal.frx":719A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPolice 
      Caption         =   "Police"
      Begin VB.Menu mnuAffichage 
         Caption         =   "Info"
         Index           =   1
      End
      Begin VB.Menu mnuAffichage 
         Caption         =   "Grille"
         Index           =   2
      End
      Begin VB.Menu mnuAffichage 
         Caption         =   "Comparer"
         Index           =   3
      End
      Begin VB.Menu z00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImprime 
         Caption         =   "Imprimer la Sélection"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuImprimeTout 
         Caption         =   "Tout Imprimer"
      End
      Begin VB.Menu mnuCatalogue 
         Caption         =   "Imprimer un Catalogue"
      End
      Begin VB.Menu z1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuEdition 
      Caption         =   "Edition"
      Enabled         =   0   'False
      Begin VB.Menu mnuCopie 
         Caption         =   "Copier"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Options"
      Begin VB.Menu mnuFiltre 
         Caption         =   "Filtre"
      End
      Begin VB.Menu z20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEchelle 
         Caption         =   "Echelle"
      End
      Begin VB.Menu z30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLanguage 
         Caption         =   "Language"
         Begin VB.Menu mnuFrancais 
            Caption         =   "Français"
         End
         Begin VB.Menu mnuAnglais 
            Caption         =   "English"
         End
      End
   End
End
Attribute VB_Name = "frmPrinc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AfficheEchelle As Integer
Dim m0020() As Integer
Dim m0052() As Integer
Dim EcritDansIni As String * 255

Sub PrépareAffichage(p00B4 As Integer)
    If p00B4 = Choix Then
        Exit Sub
    End If
    P_Clip(Choix).Visible = False
    mnuAffichage(Choix).Checked = False
    If cmbPoliceHaut.ListCount <> 0 Then
        P_Clip(p00B4).Visible = True
    End If
    mnuAffichage(p00B4).Checked = True
    Toolbar1.Buttons(p00B4).Value = tbrPressed
    Choix = p00B4
    Select Case Choix
        Case 1:
            TailleHaut.Visible = False
            mnuEdition.Enabled = False
            TrouveInfoDePolice
        Case 2:
            PrépareGrille
            TailleHaut.Visible = False
        Case 3:
            TailleHaut.Visible = True
            mnuEdition.Enabled = False
            If cmbPoliceHaut.ListCount = 0 Then
                Exit Sub
            End If
            stBar1.Panels.Item(1).Text = " " & cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " / " & cmbPoliceBas.List(cmbPoliceBas.ListIndex)
            txtAfficheHaut.FontName = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
            txtAfficheHaut.FontSize = TailleHaut.List(TailleHaut.ListIndex)
            txtAfficheBas.Text = txtAfficheHaut.Text
    End Select
End Sub
Sub PrépareGrille()
    Dim TraitDeLigne As Integer
    If cmbPoliceHaut.ListCount = 0 Then
        Exit Sub
    End If
    mnuEdition.Enabled = True
    stBar1.Panels.Item(1).Text = " " & cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
    GridFont.FontName = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
    GridFont.FontBold = Toolbar1.Buttons(5).Value
    GridFont.FontItalic = Toolbar1.Buttons(6).Value
    GridFont.FontUnderline = Toolbar1.Buttons(7).Value
    If AfficheEchelle = True Then
        GridFont.Col = 0
        GridFont.Row = 0
        GridFont.Text = "+"
        For TraitDeLigne = 1 To 14
            GridFont.Col = TraitDeLigne
            GridFont.Text = TraitDeLigne - 1
        Next TraitDeLigne
        GridFont.Col = 0
        For TraitDeLigne = 1 To 16
            GridFont.Row = TraitDeLigne
            GridFont.Text = 32 + (TraitDeLigne - 1) * 14
        Next TraitDeLigne
        For TraitDeLigne = 0 To 223
            GridFont.Col = TraitDeLigne Mod 14 + 1
            GridFont.Row = TraitDeLigne \ 14 + 1
            GridFont.Text = Chr$(TraitDeLigne + Asc(" "))
        Next TraitDeLigne
    Else
        For TraitDeLigne = 0 To 223
            GridFont.Col = TraitDeLigne Mod 14
            GridFont.Row = TraitDeLigne \ 14
            GridFont.Text = Chr$(TraitDeLigne + Asc(" "))
        Next TraitDeLigne
    End If
    S_Char.FontName = GridFont.FontName
    S_Char.FontItalic = False
    S_Char.FontUnderline = False
    S_Char.FontBold = True
    GridFont.Col = 0
    GridFont.Row = 0
End Sub

Private Sub Form_Initialize()
Language = GetPrivateProfileString(Titre, "Language", "Francais", EcritDansIni, 30, App.Path & "\" & App.EXEName & ".INI")
Language = Left$(EcritDansIni, 8)
Charger Language
End Sub

Sub Form_Load()
    Dim l0108 As TEXTMETRIC
    Dim NombreDeCarac As Integer
    Dim l01AA As Integer
    Dim l01B4 As Variant
    Dim l01B8 As Integer
    frmPrinc.MousePointer = vbHourglass
    If Language = "Francais" Then
    stBar1.Panels.Item(1).Text = "Chargement des Polices"
    stBar1.Panels.Item(2).Text = "Un instant..."
    Else
    stBar1.Panels.Item(1).Text = "Fonts loading"
    stBar1.Panels.Item(2).Text = "One moment..."
    End If
    Me.Show vbModeless
    Me.Refresh
    NombreDeCarac = GetPrivateProfileString(Titre, "Texte", "Bienvenue", EcritDansIni, 30, App.Path & "\" & App.EXEName & ".INI")
    txtAfficheBas.Text = Left$(EcritDansIni, NombreDeCarac)
    txtAfficheHaut.Text = txtAfficheBas.Text
    NombreDeCarac = GetPrivateProfileString(Titre, "Echelle de Grille", "0", EcritDansIni, 30, App.Path & "\" & App.EXEName & ".INI")
    AfficheEchelle = Val(Left$(EcritDansIni, NombreDeCarac))
    If AfficheEchelle <> -1 Then
        AfficheEchelle = 0
    End If
    mnuEchelle.Checked = AfficheEchelle
    If AfficheEchelle = True Then
        GridFont.Rows = 17
        GridFont.Cols = 15
    Else
        GridFont.FixedRows = 0
        GridFont.FixedCols = 0
    End If
    NombreDeCarac = GetPrivateProfileString(Titre, "Filtre", "63", EcritDansIni, 30, App.Path & "\" & App.EXEName & ".INI")
    gv010C = Val(Left$(EcritDansIni, NombreDeCarac))
    If gv010C < 0 Or gv010C > 63 Then
        gv010C = 63
    End If
    NombreDeCarac = GetPrivateProfileString(Titre, "Dernier Choix", "0", EcritDansIni, 30, App.Path & "\" & App.EXEName & ".INI")
    Choix = Val(Left$(EcritDansIni, NombreDeCarac))
    If Choix < 1 Or Choix > 3 Then
        Choix = 1
    End If
    P_Clip(Choix).Visible = True
    mnuAffichage(Choix).Checked = True
    Toolbar1.Buttons(Choix).Value = tbrPressed
    ReDim m0020(0 To Screen.FontCount - 1)
    ReDim m0052(0 To Screen.FontCount - 1)
    ProgressBar1.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Max = Screen.FontCount - 1
    For l01AA = 0 To Screen.FontCount - 1
        Me.Refresh
        ProgressBar1.Value = l01AA
        cmbToutePolice.AddItem Screen.Fonts(l01AA)
    Next l01AA
    If Language = "Francais" Then
    stBar1.Panels.Item(1).Text = "Classement par Type"
    stBar1.Panels.Item(2).Text = cmbToutePolice.ListCount & " Police(s) trouvée(s)"
    Else
    stBar1.Panels.Item(1).Text = "Classification by Type"
    stBar1.Panels.Item(2).Text = cmbToutePolice.ListCount & " Found font(s)"
    End If
    For l01AA = 0 To cmbToutePolice.ListCount - 1
        ProgressBar1.Value = l01AA
        frmPrinc.FontName = cmbToutePolice.List(l01AA)
        l01B4 = GetTextMetrics(frmPrinc.hdc, l0108)
        m0020(l01AA) = (l0108.tmPitchAndFamily) And &HF  ' 15
        m0052(l01AA) = (l0108.tmPitchAndFamily) And &HF0 ' 240
    Next l01AA
    If Language = "Francais" Then
    stBar1.Panels.Item(1).Text = "Préparation pour l'affichage"
    stBar1.Panels.Item(2).Text = "Patientez..."
    Else
    stBar1.Panels.Item(1).Text = "Preparation for display"
    stBar1.Panels.Item(2).Text = "Wait..."
    End If
    For l01AA = 0 To Screen.FontCount - 1
        ProgressBar1.Value = l01AA
        Select Case m0052(l01AA)
            Case 0
                l01B8 = ((gv010C And &H1) = 1) ' 1
            Case 16
                l01B8 = ((gv010C And &H2) = 2)
            Case 32
                l01B8 = ((gv010C And &H4) = 4)
            Case 48
                l01B8 = ((gv010C And &H8) = 8)
            Case 64
                l01B8 = ((gv010C And &H10) = 16)
            Case 80
                l01B8 = ((gv010C And &H20) = 32)
            Case Else
                l01B8 = False
        End Select
        If l01B8 = True Then
            cmbPoliceHaut.AddItem Screen.Fonts(l01AA)
            cmbPoliceBas.AddItem Screen.Fonts(l01AA)
        End If
    Next l01AA
    ProgressBar1.Visible = False
    If cmbPoliceHaut.ListCount <> 0 Then
        mnuImprime.Enabled = True
        mnuImprimeTout.Enabled = True
        cmbPoliceBas.ListIndex = 0
        cmbPoliceHaut.ListIndex = 0
        P_Clip(Choix).Visible = True
    Else
        mnuImprime.Enabled = False
        mnuImprimeTout.Enabled = False
        P_Clip(1).Visible = False
        P_Clip(2).Visible = False
        P_Clip(3).Visible = False
        stBar1.Panels.Item(1).Text = ""
    End If
    stBar1.Panels.Item(2).Text = NombreDePoliceDansFamille()
    For l01AA = 4 To 80
        TailleHaut.AddItem Str$(l01AA)
        TailleBas.AddItem Str$(l01AA)
    Next l01AA
    TailleHaut.ListIndex = 39
    txtAfficheHaut.FontSize = TailleHaut.List(TailleHaut.ListIndex)
    TailleBas.ListIndex = 39
    txtAfficheBas.FontSize = TailleBas.List(TailleHaut.ListIndex)
    txtAfficheBas.Text = txtAfficheHaut.Text
    frmPrinc.MousePointer = vbArrow
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
Exit Sub
Else
Me.Height = 6420
Me.Width = 7875
End If
End Sub

Sub GridFont_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
    If Asc(GridFont.Text) <> 38 Then
        S_Char.Caption = GridFont.Text
    Else
        S_Char.Caption = "&" & GridFont.Text
    End If
    If Y < 75 Then
        Y = 75
    End If
    If Y > 3255 Then
        Y = 3255
    End If
    If X > 6465 Then
        X = 6465
    End If
    S_Char.Move X, Y
    S_Char.Visible = True
    S_Char.ZOrder 0
    If Language = "Francais" Then
    stBar1.Panels.Item(1).Text = "Hex =  " & Hex$(Asc(GridFont.Text)) & "  ----  Touche = (" & Format$(Asc(GridFont.Text), "000") & ")"
    Else
    stBar1.Panels.Item(1).Text = "Hex =  " & Hex$(Asc(GridFont.Text)) & "  ----  Key = (" & Format$(Asc(GridFont.Text), "000") & ")"
    End If
    Case 2
    PopupMenu mnuEdition
    End Select
End Sub
Sub GridFont_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    S_Char.Visible = False
    stBar1.Panels.Item(1).Text = " " & cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
End Sub
Sub cmbPoliceHaut_Click()
    Select Case Choix
        Case 1:
            TrouveInfoDePolice
        Case 2:
            PrépareGrille
        Case 3:
            stBar1.Panels.Item(1).Text = " " & cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " / " & cmbPoliceBas.List(cmbPoliceBas.ListIndex)
            txtAfficheHaut.FontName = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
            txtAfficheHaut.FontBold = Toolbar1.Buttons(5).Value
            txtAfficheHaut.FontItalic = Toolbar1.Buttons(6).Value
            txtAfficheHaut.FontUnderline = Toolbar1.Buttons(7).Value
    End Select
End Sub
Sub cmbPoliceBas_Click()
    stBar1.Panels.Item(1).Text = " " & cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " / " & cmbPoliceBas.List(cmbPoliceBas.ListIndex)
    txtAfficheBas.FontName = cmbPoliceBas.List(cmbPoliceBas.ListIndex)
    txtAfficheBas.FontBold = Toolbar2.Buttons(1).Value
    txtAfficheBas.FontItalic = Toolbar2.Buttons(2).Value
    txtAfficheBas.FontUnderline = Toolbar2.Buttons(3).Value
End Sub

Private Sub mnuAnglais_Click()
Charger "Anglais"
End Sub

Private Sub mnuCatalogue_Click()
    Dim NuméroDeLaPolice As Variant
    Dim NombreDePolice As Integer
    Dim NoDePage As Byte
    Dim LaDate As String
    Me.Refresh
    If Language = "Francais" Then
    LaDate = Format(Date, "dddd d mmm yyyy")
    stBar1.Panels.Item(1).Text = "Préparation du Catalogue en cours..."
    stBar1.Panels.Item(2).Text = "Patientez..."
    Else
    LaDate = Format(Date, "yyyy mmm d")
    stBar1.Panels.Item(1).Text = "Préparation du Catalogue en cours..."
    stBar1.Panels.Item(2).Text = "Patientez..."
    End If
    ProgressBar1.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Max = frmPrinc.cmbToutePolice.ListCount
    Screen.MousePointer = 11
    Printer.FontName = "Arial"
    Printer.ForeColor = vbBlue
    Printer.FontSize = 28
    Printer.FontItalic = False
    Printer.FontBold = True
    Printer.FontUnderline = False
    Printer.FontStrikethru = False
    NombreDePolice = frmPrinc.cmbToutePolice.ListCount
    Printer.DrawMode = 1
    Printer.ScaleMode = 3
    If Language = "Francais" Then
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Len("Catalogue de Polices de Caractères") * Printer.FontSize)
    Printer.CurrentY = 200
    Printer.Print "Catalogue de Polices de Caractères"
    Else
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Len("Fonts Catalog") * Printer.FontSize)
    Printer.CurrentY = 200
    Printer.Print "Fonts Catalog"
    End If
    Printer.ScaleMode = 6
    frmPrinc.Picture1.ScaleMode = 6
    Printer.PaintPicture frmPrinc.Picture1.Picture, (Printer.ScaleWidth / 2 - frmPrinc.Picture1.ScaleWidth / 2), (Printer.ScaleHeight / 2 - frmPrinc.Picture1.ScaleHeight / 2)
    Printer.ScaleMode = 3
    Printer.FontName = "Arial"
    Printer.FontSize = 7
    Printer.FontBold = False
    Printer.CurrentY = 2900
    If Language = "Francais" Then
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Len("Imprimé le " & LaDate) * Printer.FontSize)
    Printer.Print "Imprimer le "; LaDate
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Len("Par Police Pro") * Printer.FontSize)
    Printer.CurrentY = 2950
    Printer.Print "Par Police Pro"
    Else
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Len("Printed on " & LaDate) * Printer.FontSize)
    Printer.Print "Imprimer le "; LaDate
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Len("With Police Pro") * Printer.FontSize)
    Printer.CurrentY = 2950
    Printer.Print "With Police Pro"
    End If
    Printer.NewPage
    Printer.ForeColor = vbBlack
    Printer.CurrentY = 100
    For NuméroDeLaPolice = 0 To NombreDePolice
        On Error Resume Next
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.CurrentX = 100
        Printer.Print frmPrinc.cmbToutePolice.List(NuméroDeLaPolice);
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.Print "    Taille: 12"
        Printer.FontSize = 12
        Printer.FontName = frmPrinc.cmbToutePolice.List(NuméroDeLaPolice)
        Printer.CurrentX = 100
        Printer.Print "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z"
        Printer.CurrentX = 100
        Printer.Print "a b c d e f g h i j k l m n o p q r s t u v w x y z"
        Printer.CurrentX = 100
        Printer.Print "1 2 3 4 5 6 7 8 9 0" & " É é È è Ê ê Ë ë Ç ç À à"
        Printer.Print
        If Printer.CurrentY > 2710 Then
            Printer.Print
            Printer.FontName = "Arial"
            Printer.FontBold = True
            Printer.FontSize = 8
            Printer.CurrentX = Printer.ScaleWidth / 2 - (Len("Page " & NoDePage) * Printer.FontSize)
            NoDePage = NoDePage + 1
            Printer.Print "Page "; NoDePage
            Printer.NewPage
            Printer.CurrentY = 100
        End If
        ProgressBar1.Value = NuméroDeLaPolice
        Me.Refresh
    Next NuméroDeLaPolice
    Printer.FontName = "Arial"
    Printer.Print
    Printer.EndDoc
    Screen.MousePointer = 0
    ProgressBar1.Visible = False
    stBar1.Panels.Item(1).Text = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
    stBar1.Panels.Item(2).Text = NombreDePoliceDansFamille()
End Sub


Private Sub mnuFrancais_Click()
Charger "Francais"
End Sub


Sub TailleHaut_Click()
    Select Case Choix
        Case 1:
        Case 2:
        Case 3:
            txtAfficheHaut.FontSize = TailleHaut.List(TailleHaut.ListIndex)
            txtAfficheHaut.FontName = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
    End Select
End Sub
Sub TailleBas_Click()
    txtAfficheBas.FontSize = TailleBas.List(TailleBas.ListIndex)
End Sub
Sub TrouveInfoDePolice()
    Dim l022C As TEXTMETRIC
    Dim l022E As Variant
    If cmbPoliceHaut.ListCount = 0 Then
        Exit Sub
    End If
    stBar1.Panels.Item(1).Text = " " & cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
    frmPrinc.FontName = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
    l022E = GetTextMetrics(frmPrinc.hdc, l022C)
    S_Family.Caption = familleDePolice(l022C)
    lblScale(0).Caption = l022C.tmExternalLeading
    lblScale(1).Caption = l022C.tmInternalLeading
    lblScale(2).Caption = l022C.tmHeight
    lblScale(3).Caption = l022C.tmAscent
    lblScale(4).Caption = l022C.tmDescent
    S_DefaultChar.Caption = l022C.tmFirstChar
    I_ttf.Visible = ((m0020(cmbPoliceHaut.ListIndex) And &H4) = 4)
    Select Case m0020(cmbPoliceHaut.ListIndex)
        Case 0
        If Language = "Francais" Then
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " est une Police à Points Fixes."
        Else
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " is a font with fixed points."
        End If
        Case 1
        If Language = "Francais" Then
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " est une Police Proportionnelle."
            Else
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " is a proportional font."
            End If
        Case 2
        If Language = "Francais" Then
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " est une Police Vectorielle à points fixes."
            Else
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " is a vectorial font with fixed points."
            End If
        Case 3
        If Language = "Francais" Then
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " est une Police Vectorielle proportionnelle."
            Else
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " is a vectorial proportional font."
            End If
        Case 4
        If Language = "Francais" Then
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " est une Police True Type."
            Else
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " is a True Type font."
            End If
        Case 5
        If Language = "Francais" Then
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " est une Police True Type à points fixes."
            Else
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " is a True Type font with fixed points."
            End If
        Case 6
        If Language = "Francais" Then
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " est une Police à Points Fixes."
            Else
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " is a fixed points font."
        End If
        Case 7
        If Language = "Francais" Then
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " est une Police Proportionnelle True Type."
            Else
            S_FontType.Caption = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex) & " is a proportional True Type font."
        End If
    End Select
End Sub
Sub mnuAffichage_Click(Index As Integer)
    PrépareAffichage Index
End Sub
Sub mnuCopie_Click()
    Clipboard.Clear
    Clipboard.SetText GridFont.Text
End Sub
Sub mnuQuitter_Click()
    Dim ÉcrireIni As Integer
    ÉcrireIni = WritePrivateProfileString(Titre, "Dernier Choix", LTrim$(Str$(Choix)), App.Path & "\" & App.EXEName & ".INI")
    ÉcrireIni = WritePrivateProfileString(Titre, "Echelle de Grille", LTrim$(Str$(AfficheEchelle)), App.Path & "\" & App.EXEName & ".INI")
    ÉcrireIni = WritePrivateProfileString(Titre, "Texte", txtAfficheBas.Text, App.Path & "\" & App.EXEName & ".INI")
    ÉcrireIni = WritePrivateProfileString(Titre, "Language", Language, App.Path & "\" & App.EXEName & ".INI")
    End
End Sub
Sub mnuFiltre_Click()
    Dim l0260 As Integer
    Dim l0266 As Integer
    Dim l0268 As Integer
    l0260 = gv010C
    frmFiltre.Show (vbModal)
    Me.Refresh
    If gv010C = -1 Then
        gv010C = l0260
        Exit Sub
    End If
    If gv010C <> l0260 Then
        cmbPoliceHaut.Clear
        cmbPoliceBas.Clear
        ProgressBar1.Visible = True
        ProgressBar1.Min = 0
        ProgressBar1.Max = Screen.FontCount - 1
        stBar1.Panels.Item(1).Text = "Préparation pour l'affichage"
        stBar1.Panels.Item(2).Text = "Patientez..."
        For l0266 = 0 To Screen.FontCount - 1
            ProgressBar1.Value = l0266
            Select Case m0052(l0266)
                Case 0
                    l0268 = ((gv010C And &H1) = 1)
                Case 16
                    l0268 = ((gv010C And &H2) = 2) ' Roman
                Case 32
                    l0268 = ((gv010C And &H4) = 4) ' Swiss
                Case 48
                    l0268 = ((gv010C And &H8) = 8) ' Modern
                Case 64
                    l0268 = ((gv010C And &H10) = 16) ' Script
                Case 80
                    l0268 = ((gv010C And &H20) = 32) ' Décorative
                Case Else
                    l0268 = False ' Toute autre police
            End Select
            If l0268 = True Then
                cmbPoliceHaut.AddItem Screen.Fonts(l0266)
                cmbPoliceBas.AddItem Screen.Fonts(l0266)
            End If
        Next l0266
        If cmbPoliceHaut.ListCount <> 0 Then
            mnuImprime.Enabled = True
            mnuImprimeTout.Enabled = True
            cmbPoliceBas.ListIndex = 0
            cmbPoliceHaut.ListIndex = 0
            P_Clip(Choix).Visible = True
        Else
            mnuImprime.Enabled = False
            mnuImprimeTout.Enabled = False
            P_Clip(1).Visible = False
            P_Clip(2).Visible = False
            P_Clip(3).Visible = False
            stBar1.Panels.Item(1).Text = ""
        End If
        stBar1.Panels.Item(2).Text = NombreDePoliceDansFamille()
        Select Case Choix
            Case 0:
                TrouveInfoDePolice
            Case 1:
                PrépareGrille
            Case 2:
                If cmbPoliceHaut.ListCount = 0 Then
                    Exit Sub
                End If
                txtAfficheHaut.FontName = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
                txtAfficheHaut.FontSize = TailleHaut.List(TailleHaut.ListIndex)
                txtAfficheBas.Text = txtAfficheHaut.Text
        End Select
    End If
    ProgressBar1.Visible = False
End Sub
Sub mnuEchelle_Click()
    AfficheEchelle = Not AfficheEchelle
    mnuEchelle.Checked = AfficheEchelle
    If AfficheEchelle = True Then
        GridFont.Rows = 17
        GridFont.Cols = 15
        GridFont.FixedRows = 1
        GridFont.FixedCols = 1
    Else
        GridFont.Rows = 16
        GridFont.Cols = 14
        GridFont.FixedRows = 0
        GridFont.FixedCols = 0
    End If
    If Choix = 2 Then
        PrépareGrille
    End If
End Sub
Sub mnuImprime_Click()
    Dim UneSélection As String
    Dim l020C As TEXTMETRIC
    Dim l0212 As Variant
    Dim Distance As Integer
    Dim l021C As Integer
    Dim l021E As Integer
    Dim l0220 As Integer
    Dim l0222 As Integer
    Dim l0224 As Integer
    If Language = "Francais" Then
    stBar1.Panels.Item(1).Text = "Préparation pour l'impression"
    stBar1.Panels.Item(2).Text = "1 Page à Imprimer"
    Else
    stBar1.Panels.Item(1).Text = "Preparation for printing"
    stBar1.Panels.Item(2).Text = "1 Page to print"
    End If
    Me.Refresh
    UneSélection = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
    If cmbPoliceHaut.ListCount = 0 Then
        Exit Sub
    End If
    frmPrinc.MousePointer = 11
    frmPrinc.FontName = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
    l0212 = GetTextMetrics(frmPrinc.hdc, l020C)
    Distance = 32
    Printer.FillStyle = 0
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.FontItalic = False
    Printer.FontUnderline = False
    Printer.FontSize = 8
    Printer.ScaleMode = 3
    If Language = "Francais" Then
    Printer.Print UneSélection;
    Printer.CurrentX = Printer.ScaleWidth - (Len("Famille: " & familleDePolice(l020C)) + 50 * Printer.FontSize)
    Printer.Print "Famille: " & familleDePolice(l020C)
    Else
    Printer.Print UneSélection;
    Printer.CurrentX = Printer.ScaleWidth - (Len("Familly: " & familleDePolice(l020C)) + 50 * Printer.FontSize)
    Printer.Print "Familly: " & familleDePolice(l020C)
    End If
    Printer.Print
    Printer.ScaleMode = 1
    Printer.FontName = UneSélection
    Printer.FontBold = False
    l021C = (Printer.ScaleWidth - 700) / 14
    l021E = (Printer.ScaleHeight - 1000) / 16
    Printer.Line (Distance, Printer.CurrentY)-(Distance + 14 * l021C, Printer.CurrentY)
    ProgressBar1.Min = 0
    ProgressBar1.Max = 15
    ProgressBar1.Visible = True
    For l0220 = 0 To 15
        ProgressBar1.Value = l0220
        For l0222 = 0 To 13
            Printer.Line (Distance + l0222 * l021C, Printer.CurrentY)-(Distance + l0222 * l021C, Printer.CurrentY + l021E)
            Printer.CurrentY = Printer.CurrentY - l021E
            Printer.FontName = "Arial"
            Printer.FontSize = 5
            Printer.Print l0220 * 14 + l0222 + Asc(" ");
            Printer.FontName = UneSélection
            Printer.FontSize = 25
            l0224 = l0220 * 14 + l0222 + Asc(" ")
            Printer.Print Chr$(l0224);
        Next l0222
        Printer.Line (Distance + 14 * l021C, Printer.CurrentY)-(Distance + 14 * l021C, Printer.CurrentY + l021E)
        Printer.Line (Distance, Printer.CurrentY)-(Distance + 14 * l021C, Printer.CurrentY)
        Printer.FontSize = 25
    Next l0220
    ProgressBar1.Visible = False
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.FontSize = 8
    Printer.Print
    Printer.ScaleMode = 3
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Len("Imprimer par Police Pro") * Printer.FontSize)
    If Language = "Francais" Then
    Printer.Print "Imprimer par Police Pro"
    Else
    Printer.Print "Printed by Police Pro"
    End If
    Printer.Print
    Printer.FontSize = 5
    Printer.ScaleMode = 1
    Printer.EndDoc
    frmPrinc.MousePointer = 1 ' Flèche
    stBar1.Panels.Item(1).Text = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
    stBar1.Panels.Item(2).Text = NombreDePoliceDansFamille()
End Sub
Sub mnuImprimeTout_Click()
    Dim l0270 As Integer
    Dim LesSélections As String
    Dim l020C As TEXTMETRIC
    Dim l0212 As Variant
    Dim Distance As Integer
    Dim l021C As Integer
    Dim l021E As Integer
    Dim l0220 As Integer
    Dim l0222 As Integer
    Dim l0224 As Integer
    Dim NumeroDePage As Integer
    Dim LaDate As String
    Dim Réponse As Byte
    If Language = "Francais" Then
    Réponse = MsgBox(cmbPoliceHaut.ListCount & " Page(s) à Imprimer voulez-vous poursuivre?", vbOKCancel, "Information")
    If Réponse = vbCancel Then
        Exit Sub
    End If
    Else
    Réponse = MsgBox(cmbPoliceHaut.ListCount & " Page(s) to print, do you want to continue?", vbOKCancel, "Information")
    If Réponse = vbCancel Then
        Exit Sub
    End If
    End If
    
    If Language = "Francais" Then
    LaDate = Format(Date, "dd mmm yyyy")
    stBar1.Panels.Item(1).Text = "Sélection des Polices à Imprimer"
    stBar1.Panels.Item(2).Text = cmbPoliceHaut.ListCount & " Page(s) à Imprimer"
    Else
    LaDate = Format(Date, "yyyy mmm d")
    stBar1.Panels.Item(1).Text = "Selection of fonts to be print"
    stBar1.Panels.Item(2).Text = cmbPoliceHaut.ListCount & " Page(s) to print"
    End If
    ProgressBar1.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Max = cmbPoliceHaut.ListCount - 1
    For l0270 = 0 To cmbPoliceHaut.ListCount - 1
        Me.Refresh
        ProgressBar1.Value = l0270
        LesSélections$ = cmbPoliceHaut.List(l0270)
        If cmbPoliceHaut.ListCount = 0 Then
            Exit Sub
        End If
        NumeroDePage = NumeroDePage + 1
        frmPrinc.MousePointer = 11 ' Sablier
        frmPrinc.FontName = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
        l0212 = GetTextMetrics(frmPrinc.hdc, l020C)
        Distance = 32
        Printer.FillStyle = 0
        Printer.FontName = "Arial"
        Printer.FontBold = True
        Printer.FontItalic = False
        Printer.FontUnderline = False
        Printer.FontSize = 8
        Printer.CurrentX = Distance
        If Language = "Francais" Then
        Printer.Print LesSélections; Tab; Tab; "Famille: " & familleDePolice(l020C)
        Else
        Printer.Print LesSélections; Tab; Tab; "Familly: " & familleDePolice(l020C)
        End If
        Printer.FontName = LesSélections
        Printer.FontBold = False
        l021C = (Printer.ScaleWidth - 700) / 14
        l021E = (Printer.ScaleHeight - 1000) / 16
        Printer.Line (Distance, Printer.CurrentY)-(Distance + 14 * l021C, Printer.CurrentY)
        For l0220 = 0 To 15
            For l0222 = 0 To 13
                Printer.Line (Distance + l0222 * l021C, Printer.CurrentY)-(Distance + l0222 * l021C, Printer.CurrentY + l021E)
                Printer.CurrentY = Printer.CurrentY - l021E
                Printer.FontName = "Arial"
                Printer.FontSize = 5
                Printer.Print l0220 * 14 + l0222 + Asc(" ");
                Printer.FontName = LesSélections
                Printer.FontSize = 25
                l0224 = l0220 * 14 + l0222 + Asc(" ")
                Printer.Print Chr$(l0224);
            Next l0222
            Printer.Line (Distance + 14 * l021C, Printer.CurrentY)-(Distance + 14 * l021C, Printer.CurrentY + l021E)
            Printer.Line (Distance, Printer.CurrentY)-(Distance + 14 * l021C, Printer.CurrentY)
            Printer.FontSize = 25
        Next l0220
        Printer.FontName = "Arial"
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.Print
        Printer.CurrentX = Printer.ScaleWidth - 700
        Printer.Print "Page "; NumeroDePage
        Printer.FontSize = 5
        Printer.CurrentX = Distance
    Next l0270
    ProgressBar1.Visible = False
    Printer.EndDoc
    frmPrinc.MousePointer = 1 ' Flèche
    stBar1.Panels.Item(1).Text = cmbPoliceHaut.List(cmbPoliceHaut.ListIndex)
    stBar1.Panels.Item(2).Text = NombreDePoliceDansFamille()
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "Grid"
            PrépareAffichage Button.Index
        Case "Info"
            PrépareAffichage Button.Index
        Case "Compare"
            PrépareAffichage Button.Index
        Case "Gras"
            txtAfficheHaut.FontBold = Button.Value
            GridFont.FontBold = Button.Value
        Case "Souligne"
            txtAfficheHaut.FontUnderline = Button.Value
            GridFont.FontUnderline = Button.Value
        Case "Italique"
            txtAfficheHaut.FontItalic = Button.Value
            GridFont.FontItalic = Button.Value
        Case "Imprime"
            mnuImprime_Click
    End Select
End Sub
Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "Gras2"
            txtAfficheBas.FontBold = Button.Value
        Case "Italique2"
            txtAfficheBas.FontItalic = Button.Value
        Case "Souligne2"
            txtAfficheBas.FontUnderline = Button.Value
    End Select
End Sub
Sub txtAfficheHaut_Change()
    txtAfficheBas.Text = txtAfficheHaut.Text
End Sub
Sub txtAfficheBas_Change()
    txtAfficheHaut.Text = txtAfficheBas.Text
End Sub

Private Sub Charger(Langue As String)
If Langue = "Francais" Then
Language = "Francais"
mnuFrancais.Checked = True
mnuAnglais.Checked = False
mnuPolice.Caption = "Police"
mnuAffichage(2).Caption = "Grille"
mnuAffichage(3).Caption = "Comparer"
mnuImprime.Caption = "Imprimer la sélection"
mnuImprimeTout.Caption = "Imprimer tout"
mnuCatalogue.Caption = "Imprimer un Catalogue"
mnuQuitter.Caption = "Quitter"
mnuCopie.Caption = "Copier"
mnuFiltre.Caption = "Filtre"
mnuEchelle.Caption = "Échelle"
mnuLanguage.Caption = "Langue"
Toolbar1.Buttons(1).ToolTipText = "Info sur la police sélectionnée"
Toolbar1.Buttons(2).ToolTipText = "Caractères disponibles"
Toolbar1.Buttons(3).ToolTipText = "Pour comparer deux polices"
Toolbar1.Buttons(5).ToolTipText = "Afficher en caractère gras"
Toolbar1.Buttons(6).ToolTipText = "Afficher en caractère Italique"
Toolbar1.Buttons(7).ToolTipText = "Afficher en caractère Souligné"
Toolbar1.Buttons(9).ToolTipText = "Imprime la police sélectionnée"
frmFiltre.Frame1.Caption = "Type de police"
frmFiltre.Ch3D_Ext(1).Caption = "Divers"
frmFiltre.Ch3D_Ext(6).Caption = "Décorative"
frmFiltre.cmdAnnule.Caption = "Annuler"
Label1.Caption = "Type de police:"
S_DefaultCharTitle.Caption = "Caractère par Défaut:"
S_FamilyTitle.Caption = "Nom de la Famille:"
S_ScaleName(0).Caption = "(Leading Externe )"
S_ScaleName(1).Caption = "(Leading Interne)"
S_ScaleName(2).Caption = "(Hauteur)"
S_ScaleName(3).Caption = "(Montée)"
S_ScaleName(4).Caption = "(Descente)"
If S_Family.Caption = "Not classified!..." Then
S_Family.Caption = "Non Classée!..."
ElseIf S_Family.Caption = "Decorative" Then
S_Family.Caption = "Décorative"
End If
TrouveInfoDePolice
stBar1.Panels.Item(2).Text = NombreDePoliceDansFamille()
Else
Language = "Anglais"
mnuFrancais.Checked = False
mnuAnglais.Checked = True
mnuPolice.Caption = "Font"
mnuAffichage(2).Caption = "Grid"
mnuAffichage(3).Caption = "Compare"
mnuImprime.Caption = "Print selection"
mnuImprimeTout.Caption = "Print all"
mnuCatalogue.Caption = "Print a catalog"
mnuQuitter.Caption = "Quit"
mnuCopie.Caption = "Copy"
mnuFiltre.Caption = "Filter"
mnuEchelle.Caption = "Scale"
mnuLanguage.Caption = "Language"
Toolbar1.Buttons(1).ToolTipText = "Information on the selected font"
Toolbar1.Buttons(2).ToolTipText = "Available character"
Toolbar1.Buttons(3).ToolTipText = "To compare two fonts"
Toolbar1.Buttons(5).ToolTipText = "Bold"
Toolbar1.Buttons(6).ToolTipText = "Italic"
Toolbar1.Buttons(7).ToolTipText = "Underline"
Toolbar1.Buttons(9).ToolTipText = "Print the selected font"
frmFiltre.Frame1.Caption = "Type of font"
frmFiltre.Ch3D_Ext(1).Caption = "Various"
frmFiltre.Ch3D_Ext(6).Caption = "Decorative"
frmFiltre.cmdAnnule.Caption = "Cancel"
Label1.Caption = "Type of font:"
S_DefaultCharTitle.Caption = "Character by default:"
S_FamilyTitle.Caption = "Familly name:"
S_ScaleName(0).Caption = "(External Leading)"
S_ScaleName(1).Caption = "(Internal Leading)"
S_ScaleName(2).Caption = "(Height)"
S_ScaleName(3).Caption = "(Ascent)"
S_ScaleName(4).Caption = "(Descent)"
If S_Family.Caption = "Non Classée!..." Then
S_Family.Caption = "Not classified!..."
ElseIf S_Family.Caption = "Décorative" Then
S_Family.Caption = "Decorative"
TrouveInfoDePolice
End If
TrouveInfoDePolice
stBar1.Panels.Item(2).Text = NombreDePoliceDansFamille()
End If
End Sub

