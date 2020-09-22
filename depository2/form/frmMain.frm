VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Code Depository"
   ClientHeight    =   6645
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgTree 
      Left            =   1035
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTlbr 
      Left            =   1650
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":116E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2708
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E74
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3718
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   8655
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   810
      Width           =   8655
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Language:"
         Height          =   270
         Index           =   0
         Left            =   15
         TabIndex        =   7
         Tag             =   " TreeView:"
         Top             =   15
         Width           =   2010
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codes:"
         Height          =   270
         Index           =   1
         Left            =   2670
         TabIndex        =   6
         Tag             =   " ListView:"
         Top             =   15
         Width           =   3210
      End
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4995
      Left            =   0
      TabIndex        =   3
      Top             =   1140
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   8811
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imgTree"
      Appearance      =   1
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   8610
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   2
      Top             =   1110
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   6255
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7355
            Text            =   "Language"
            TextSave        =   "Language"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7355
            Text            =   "No. Of Codes"
            TextSave        =   "No. Of Codes"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2700
      Top             =   5295
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   1429
      ButtonWidth     =   1482
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgTlbr"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Code"
            Key             =   "newcode"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "open"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Key             =   "edit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "delete"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Key             =   "find"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Language"
            Key             =   "language"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Category"
            Key             =   "category"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Password"
            Key             =   "password"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "exit"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   5025
      Left            =   2640
      TabIndex        =   4
      Top             =   1125
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   8864
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      Icons           =   "imgTlbr"
      SmallIcons      =   "imgTlbr"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   10584
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Author"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mail"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Category"
         Object.Width           =   2540
         ImageIndex      =   1
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   2550
      MousePointer    =   9  'Size W E
      Top             =   1125
      Width           =   150
   End
   Begin VB.Menu mnuProgram 
      Caption         =   "Program"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDb 
      Caption         =   "Database"
      Begin VB.Menu mnuComRep 
         Caption         =   "Compact/Repair"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuCode 
      Caption         =   "Code"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbMoving As Boolean

Const sglSplitLimit = 1500

Private Sub Form_Load()
  Call loadLanguage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer

  adoCon.Close
  Set adoCon = Nothing
  For i = Forms.Count - 1 To 1 Step -1
    Unload Forms(i)
  Next
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If Me.Width < 3000 Then Me.Width = 3000
  SizeControls imgSplitter.Left
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  With imgSplitter
    picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
  End With
  picSplitter.Visible = True
  mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim sglPos As Single
  
  If mbMoving Then
    sglPos = x + imgSplitter.Left
    If sglPos < sglSplitLimit Then
      picSplitter.Left = sglSplitLimit
    ElseIf sglPos > Me.Width - sglSplitLimit Then
      picSplitter.Left = Me.Width - sglSplitLimit
    Else
      picSplitter.Left = sglPos
    End If
  End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  SizeControls picSplitter.Left
  picSplitter.Visible = False
  mbMoving = False
End Sub

Sub SizeControls(x As Single)
  On Error Resume Next
  
  'set the width
  If x < 1500 Then x = 1500
  If x > (Me.Width - 1500) Then x = Me.Width - 1500
  tvTreeView.Width = x
  imgSplitter.Left = x
  lvListView.Left = x + 40
  lvListView.Width = Me.Width - (tvTreeView.Width + 140)
  lblTitle(0).Width = tvTreeView.Width
  lblTitle(1).Left = lvListView.Left + 20
  lblTitle(1).Width = lvListView.Width - 40

  'set the top
  
  If tbToolBar.Visible Then
    tvTreeView.Top = tbToolBar.Height + picTitles.Height
  Else
    tvTreeView.Top = picTitles.Height
  End If

  lvListView.Top = tvTreeView.Top
  
  'set the height
  If sbStatusBar.Visible Then
    tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
  Else
    tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
  End If
  
  lvListView.Height = tvTreeView.Height
  imgSplitter.Top = tvTreeView.Top
  imgSplitter.Height = tvTreeView.Height
End Sub

Sub loadLanguage()
Dim rs As New ADODB.Recordset
Dim lst As ListItem
Dim i As Byte

  tvTreeView.Nodes.Clear
  lvListView.ListItems.Clear
  openRecordSet rs, "select * from lang order by lang_name", adoCon
  With rs
    If .RecordCount < 1 Then Exit Sub
    .MoveFirst
    Do While Not .EOF
        tvTreeView.Nodes.Add , , , rs!lang_name, 1, 2
        .MoveNext
    Loop
    .MoveFirst
    tvTreeView.Nodes.Item(.AbsolutePosition).EnsureVisible
    tvTreeView.Nodes.Item(.AbsolutePosition).Selected = True
  End With
End Sub

Private Sub lvListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case ColumnHeader.Index
  Case 1
    If lvListView.ColumnHeaders(ColumnHeader.Index).Icon = 2 Then
      SortColumn lvListView, ColumnHeader.Index, sortAlphanumeric, sortDescending
      lvListView.ColumnHeaders(ColumnHeader.Index).Icon = 1
    Else
     SortColumn lvListView, ColumnHeader.Index, sortAlphanumeric, sortAscending
      lvListView.ColumnHeaders(ColumnHeader.Index).Icon = 2
    End If
 
    
  Case 2
    SortColumn lvListView, 2, 0, 3
  Case 3
    SortColumn lvListView, 3, 0, 3
  Case 4
    If lvListView.ColumnHeaders(ColumnHeader.Index).Icon = 2 Then
      SortColumn lvListView, ColumnHeader.Index, sortAlphanumeric, sortDescending
      lvListView.ColumnHeaders(ColumnHeader.Index).Icon = 1
    Else
     SortColumn lvListView, ColumnHeader.Index, sortAlphanumeric, sortAscending
      lvListView.ColumnHeaders(ColumnHeader.Index).Icon = 2
    End If
End Select
 
End Sub

Private Sub lvListView_DblClick()
If lvListView.ListItems.Count < 1 Then Exit Sub
  Call loadContent(lvListView.ListItems(lvListView.SelectedItem.Index))
End Sub

Private Sub lvListView_KeyDown(KeyCode As Integer, Shift As Integer)
Dim msg As VbMsgBoxResult

  If lvListView.ListItems.Count = 0 Then Exit Sub
  If lvListView.ListItems.Item(lvListView.SelectedItem.Index) = "[No Codes]" Then Exit Sub
  If KeyCode = vbKeyDelete Then
    msg = MsgBox("Are you sure you want to delete this code? " & vbCrLf & "Title: " & lvListView.ListItems(lvListView.SelectedItem.Index), _
    vbQuestion + vbYesNo, appName)
    If msg = vbYes Then
      adoCon.Execute " delete * from codes where code_title = '" & lvListView.ListItems(lvListView.SelectedItem.Index) & "'"
      Call loadCodes
    End If
  End If
  If KeyCode = 13 Then
    Call lvListView_DblClick
  End If
End Sub

Private Sub lvListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 And lvListView.ListItems.Count >= 1 Then PopupMenu mnuCode
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show 1
End Sub

Private Sub mnuDelete_Click()
  lvListView_KeyDown vbKeyDelete, 0
End Sub

Private Sub mnuEdit_Click()
  If lvListView.ListItems.Count < 1 Then Exit Sub
  Call editCodes(lvListView.ListItems(lvListView.SelectedItem.Index).Text)
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuOpen_Click()
  Call lvListView_DblClick
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim frm As New frmNewCode

Select Case Button.Key
  Case "newcode":   With frm:   .Show:    End With
  Case "open": Call lvListView_DblClick
  Case "edit": Call mnuEdit_Click
  Case "delete": Call lvListView_KeyDown(vbKeyDelete, 0)
  Case "language": frmLang.Show 1, Me
  Case "category": MsgBox "In progress", , appName
  Case "password": MsgBox "In progress", , appName
  Case "find": frmFindCode.Show 1, Me
  Case "exit": Unload Me
End Select
End Sub

Private Sub tvTreeView_Click()
  strlang = tvTreeView.SelectedItem
  sbStatusBar.Panels(1).Text = "Language: " & strlang
  Call loadCodes
End Sub

Sub loadCodes()
Dim rs As New ADODB.Recordset
Dim strSql As String
Dim lst As ListItem
Dim i As Byte
  
  
  lvListView.ListItems.Clear
  
  strSql = "SELECT DISTINCT codes.CODE_TITLE, codes.CODE_AUTH, codes.CODE_MAIL, category.CAT_DESC, lang.LANG_NAME " & _
         "FROM category INNER JOIN (lang INNER JOIN codes ON lang.LANG_ID = codes.LANG_ID) ON category.CAT_ID = codes.CAT_ID " & _
         "WHERE lang.LANG_NAME='" & strlang & "' order by code_title"
  
  rs.Open strSql, adoCon, 1, 3
  
  If rs.RecordCount < 1 Then lvListView.ListItems.Add , , "[No Codes]", , 2: GoTo jmp
  rs.MoveFirst
  Do While Not rs.EOF
    With lvListView
      Set lst = .ListItems.Add(, , rs!CODE_TITLE, , 2)
      With lst
        .SubItems(1) = rs!code_auth
        .SubItems(2) = rs!code_mail
        .SubItems(3) = rs!cat_desc
      End With
    End With
    rs.MoveNext
  Loop
  
jmp:
  lvListView.ListItems.Item(1).Selected = True
  lvListView.SetFocus
  'lvListView.SelectedItem.Selected = True
  sbStatusBar.Panels(2).Text = "No of Codes: " & rs.RecordCount & " Code(s)"
End Sub

Sub editCodes(title As String)
Dim rs As New ADODB.Recordset
Dim frm As New frmNewCode
Dim sCodes As String
Dim sTmp As String
Dim cchChunkReceived As Long
Dim cchChunkRequested As Long

  If title = "[No Codes]" Then Exit Sub
  
  rs.Open "SELECT * FROM codes where code_title='" & title & "'", adoCon, adOpenKeyset, adLockOptimistic

  cchChunkRequested = 16
  
  Do
    sTmp = rs.Fields("code_content").GetChunk(cchChunkRequested)
    cchChunkReceived = Len(sTmp)
    If cchChunkReceived > 0 Then
      sCodes = sCodes & sTmp
    End If
  Loop While cchChunkReceived = cchChunkRequested
  
  With frm
    .bEditing = True
    .Caption = title
    .strOldTitle = title
    .strCategory = lvListView.ListItems(lvListView.SelectedItem.Index).SubItems(3)
    .txtCode = sCodes
    .txtAuth = rs!code_auth
    .txtMail = rs!code_mail
    .txtTitle = title
    .cboCat.Text = lvListView.ListItems(lvListView.SelectedItem.Index).SubItems(3)
    .cboLang.Text = strlang
    .cmdSAve.Caption = "&Update"
    .Show
  End With
  rs.Close

End Sub
