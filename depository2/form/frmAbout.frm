VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3630
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   5655
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   435
         Left            =   2025
         TabIndex        =   3
         Top             =   3075
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PskSoft™ Inc."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2115
         TabIndex        =   6
         Top             =   930
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Mobile: (+63)922-374-2323"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   2595
         Width           =   5310
      End
      Begin VB.Label Label1 
         Caption         =   "Email: hex@rt.nl"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   2325
         Width           =   5310
      End
      Begin VB.Label Label1 
         Caption         =   "Coded by: Arnel Labarda"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   2055
         Width           =   5310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code Depository"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1500
         TabIndex        =   1
         Top             =   555
         Width           =   2925
      End
      Begin VB.Image Image1 
         Height          =   810
         Left            =   195
         Picture         =   "frmAbout.frx":08CA
         Stretch         =   -1  'True
         Top             =   270
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub
