VERSION 5.00
Begin VB.Form frmNewLang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Language"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "frmNewLang.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   4770
      Begin VB.TextBox txtLang 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   0
         Top             =   300
         Width           =   4395
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2265
         TabIndex        =   1
         Top             =   750
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3465
         TabIndex        =   2
         Top             =   750
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmNewLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bEditing As Boolean
Public strOldLang As String

Private Sub cmdCancel_Click()
  bEditing = False
  Unload Me
End Sub

Private Sub cmdUpdate_Click()
Dim rs As New ADODB.Recordset

  If txtLang = "" Then
    MsgBox "Please enter language...", vbExclamation, appName
    txtLang.SetFocus
  End If
  
  
  If strOldLang = txtLang Then Exit Sub
     
  rs.Open "select * from lang where lang_name ='" & IIf(bEditing = True, strOldLang, txtLang) & "'", adoCon, 1, adLockOptimistic
  If rs.RecordCount >= 1 And bEditing = False Then
    MsgBox "Language already is the list", , appName
    txtLang.SetFocus
    Exit Sub
  End If
  
  If bEditing = False Then rs.AddNew
  rs!lang_name = txtLang
  rs.Update
  rs.Close
  Set rs = Nothing
  
  MsgBox IIf(bEditing = True, "Language Updated", "New Language added"), , appName
  bEditing = False
  Call loadLanguage
  Call frmMain.loadLanguage
  
  Unload Me
End Sub


