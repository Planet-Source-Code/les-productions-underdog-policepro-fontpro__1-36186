VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmFiltre 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3105
   ClientLeft      =   7305
   ClientTop       =   3615
   ClientWidth     =   1620
   ControlBox      =   0   'False
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
   Icon            =   "Filtre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3105
   ScaleWidth      =   1620
   Begin VB.CommandButton cmdAnnule 
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   870
      TabIndex        =   8
      Top             =   2670
      Width           =   615
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   7
      Top             =   2670
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type de Police"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2430
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1395
      Begin Threed.SSCheck Ch3D_Ext 
         Height          =   240
         Index           =   6
         Left            =   135
         TabIndex        =   6
         Top             =   2025
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Décorative"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck Ch3D_Ext 
         Height          =   240
         Index           =   5
         Left            =   135
         TabIndex        =   5
         Top             =   1674
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Script"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck Ch3D_Ext 
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   4
         Top             =   1323
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Modern"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck Ch3D_Ext 
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   3
         Top             =   972
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Swiss"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck Ch3D_Ext 
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   621
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Roman"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck Ch3D_Ext 
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   1
         Top             =   300
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "Divers"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
   End
End
Attribute VB_Name = "frmFiltre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub cmdAnnule_Click()
    gv010C = -1
    Unload frmFiltre
End Sub
Sub cmdOk_Click()
    Dim l0020 As Integer
    Dim l0032 As Integer
    gv010C = 0
    For l0020 = 1 To 6
        If Ch3D_Ext(l0020).Value = True Then
            gv010C = gv010C + (2 ^ (l0020 - 1))
        End If
    Next l0020
    l0032 = WritePrivateProfileString(Titre, "Filtre", Str$(gv010C), App.Path & "\" & App.EXEName & ".INI")
    Unload frmFiltre
End Sub
Sub Form_Load()
    Dim l0034 As Integer
    For l0034 = 1 To 6
        Ch3D_Ext(l0034).Value = gv010C And (2 ^ (l0034 - 1))
    Next l0034
End Sub

