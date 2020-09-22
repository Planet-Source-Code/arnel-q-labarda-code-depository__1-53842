VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Language"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmLang.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3450
      Left            =   45
      TabIndex        =   0
      Top             =   -60
      Width           =   5235
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   4
         Top             =   675
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Hide"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   3
         Top             =   2955
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   2
         Top             =   1125
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
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
         Height          =   375
         Left            =   3885
         TabIndex        =   1
         Top             =   210
         Width           =   1215
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5310
         Top             =   1980
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLang.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLang.frx":0C66
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLang.frx":1540
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstLang 
         Height          =   3150
         Left            =   60
         TabIndex        =   5
         Top             =   195
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   5556
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   12582912
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lang"
            Object.Width           =   8114
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  frmNewLang.Show 1, Me
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim rs As New ADODB.Recordset


  If lstLang.ListItems.Count = 0 Then Exit Sub
  
  adoCon.Execute "delete * from lang where lang_name ='" & lstLang.ListItems(lstLang.SelectedItem.Index) & "'"
  Call loadLanguage
  Call frmMain.loadLanguage
  
End Sub

Private Sub cmdEdit_Click()
  With frmNewLang
    .bEditing = True
    .strOldLang = lstLang.ListItems(lstLang.SelectedItem.Index)
    .txtLang = lstLang.ListItems(lstLang.SelectedItem.Index)
    .cmdUpdate.Caption = "&Update"
    .Show 1, Me
  End With
End Sub

Private Sub Form_Load()
  Call loadLanguage
End Sub

