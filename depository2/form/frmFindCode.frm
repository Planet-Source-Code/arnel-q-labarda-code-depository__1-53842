VERSION 5.00
Begin VB.Form frmFindCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Code"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2580
      Left            =   90
      TabIndex        =   10
      Top             =   -30
      Width           =   6900
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   5355
         TabIndex        =   9
         Top             =   2040
         Width           =   1395
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   435
         Left            =   3870
         TabIndex        =   8
         Top             =   2040
         Width           =   1395
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Title"
         Height          =   495
         Left            =   195
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Author"
         Height          =   495
         Left            =   195
         TabIndex        =   6
         Top             =   1515
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   1395
         TabIndex        =   7
         Top             =   1590
         Width           =   5355
      End
      Begin VB.TextBox txtTitle 
         Height          =   360
         Left            =   1395
         TabIndex        =   5
         Top             =   1140
         Width           =   5355
      End
      Begin VB.ComboBox cboCat 
         Height          =   360
         Left            =   1395
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   690
         Width           =   2745
      End
      Begin VB.ComboBox cboLang 
         Height          =   360
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2745
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Category"
         Height          =   240
         Index           =   2
         Left            =   195
         TabIndex        =   2
         Top             =   735
         Width           =   960
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Language"
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   0
         Top             =   285
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmFindCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub
