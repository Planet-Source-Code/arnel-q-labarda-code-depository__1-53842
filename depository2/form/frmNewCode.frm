VERSION 5.00
Begin VB.Form frmNewCode 
   Caption         =   "New Code"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtMail 
      Height          =   360
      Left            =   960
      TabIndex        =   5
      Top             =   990
      Width           =   6795
   End
   Begin VB.TextBox txtAuth 
      Height          =   360
      Left            =   960
      TabIndex        =   3
      Top             =   540
      Width           =   6795
   End
   Begin VB.ComboBox cboCat 
      Height          =   360
      Left            =   5835
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1455
      Width           =   3105
   End
   Begin VB.ComboBox cboLang 
      Height          =   360
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1425
      Width           =   3105
   End
   Begin VB.CommandButton cmdSAve 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7815
      TabIndex        =   11
      Top             =   75
      Width           =   1125
   End
   Begin VB.TextBox txtCode 
      Height          =   4845
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   1980
      Width           =   8865
   End
   Begin VB.TextBox txtTitle 
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   90
      Width           =   6795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&E-Mail:"
      Height          =   240
      Index           =   4
      Left            =   105
      TabIndex        =   4
      Top             =   1035
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Author:"
      Height          =   240
      Index           =   3
      Left            =   105
      TabIndex        =   2
      Top             =   585
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Category:"
      Height          =   240
      Index           =   2
      Left            =   4710
      TabIndex        =   8
      Top             =   1545
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Language:"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1515
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Title:"
      Height          =   240
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   720
   End
End
Attribute VB_Name = "frmNewCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bEditing As Boolean
Public strOldTitle As String
Public strCategory As String

Private Sub cmdSAve_Click()
Dim rs As New ADODB.Recordset
Dim ctl As Control
Dim cat As String
Dim lang As String

If Len(Trim(txtTitle)) < 1 Then
  MsgBox "Please enter a title for your code...", vbExclamation, ""
  txtTitle.SetFocus
  Exit Sub
End If

If Len(Trim(txtCode)) < 1 Then
  MsgBox "Please type your code...", vbExclamation, ""
  txtCode.SetFocus
  Exit Sub
End If


'//this is the best way i know to do this
openRecordSet rs, "select lang_id from lang where lang_name ='" & cboLang.Text & "'", adoCon
lang = rs!lang_id

openRecordSet rs, "select cat_id from category where cat_desc ='" & cboCat.Text & "'", adoCon
cat = rs!cat_id

If bEditing = True Then
  adoCon.Execute "UPDATE codes SET CODE_CONTENT = '" & txtCode & "'," & _
  "CODE_AUTH ='" & txtAuth & "'," & _
  "CODE_MAIL ='" & txtMail & "'," & _
  "CODE_TITLE ='" & txtTitle & "', " & _
  "LANG_ID ='" & lang & "', " & _
  "cat_ID ='" & cat & "' " & _
  "WHERE CODE_TITLE='" & strOldTitle & "' AND " & _
  "LANG_ID = (select lang_id from lang where lang_name ='" & strlang & "') AND " & _
  "CAT_ID = (select cat_id from category where cat_desc ='" & strCategory & "')"
  
Else
  adoCon.Execute "Insert into codes values('" & _
  txtCode & "','" & _
  IIf(txtAuth = "", "Anonymous", txtAuth) & "','" & _
  IIf(txtMail = "", "Not Specified", txtMail) & "','" & _
  txtTitle & "'," & _
  lang & "," & _
  cat & ")"

End If
  
frmMain.loadCodes
MsgBox IIf(bEditing = True, "Code Updated", "New code added to database"), , appName

If bEditing = True Then GoTo jmp
For Each ctl In Me
  If TypeOf ctl Is TextBox Then ctl.Text = ""
Next ctl

jmp:
  Unload Me
End Sub

Private Sub Form_Load()
  Call loadRsToCbo("select * from category", "cat_desc", cboCat)
  Call loadRsToCbo("select * from lang", "lang_name", cboLang)
  cboLang.Text = IIf(strlang = "", cboLang.List(0), strlang)
End Sub

Private Sub Form_Resize()
On Error Resume Next
  txtCode.Move 10, 1980, Me.Width - 150, Me.Height - 2500
End Sub

Private Sub txtCode_Change()
  If Len(Trim(txtCode)) >= 1 Then
    cmdSAve.Enabled = True
  Else
    cmdSAve.Enabled = False
  End If
End Sub
