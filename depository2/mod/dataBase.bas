Attribute VB_Name = "dataBase"
Option Explicit

Public Const appName = "Code Depository"
Private Const dbPass = "paske"
Public adoCon As New ADODB.Connection
Public strlang As String


Sub Main()
  Call openConnection
  frmSplash.Show 1
  frmMain.Show
End Sub

Sub openConnection()
  adoCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb" & ";Persist Security Info=False ;Jet OLEDB:Database Password=" & dbPass
End Sub

Public Sub openRecordSet(rs As ADODB.Recordset, strSource As String, cn As ADODB.Connection)
On Error GoTo err

  Set rs = Nothing: Set rs = New ADODB.Recordset
  rs.Open strSource, cn, adOpenKeyset, adLockOptimistic
  Exit Sub
  
err:
  MsgBox err.Description, , ""
End Sub

Sub loadLanguage()
Dim rs As New ADODB.Recordset
Dim lst As ListItem
Dim i As Byte
  
  frmLang.lstLang.ListItems.Clear
  openRecordSet rs, "select * from lang order by lang_name", adoCon
  
  If rs.RecordCount > 1 Then
    For i = 1 To rs.RecordCount
      With frmLang.lstLang
        Set lst = .ListItems.Add(, , rs!lang_name, , 3)
      End With
      rs.MoveNext
    Next i
  End If
End Sub

Sub loadContent(title As String)
Dim rs As New ADODB.Recordset
Dim sCodes As String
Dim sTmp As String
Dim cchChunkReceived As Long
Dim cchChunkRequested As Long
Dim frm As New frmCode

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
  rs.Close
  With frm
    .load sCodes
    .Caption = title
    .Show
  End With
End Sub

Sub loadRsToCbo(strRsSource As String, _
                        intFldIndx As String, cbo As ComboBox)
Dim rs As New ADODB.Recordset
Dim i As Integer

  Call openRecordSet(rs, strRsSource, adoCon)
  cbo.Clear
  If rs.RecordCount < 1 Then Set rs = Nothing: Exit Sub
  For i = 1 To rs.RecordCount
    cbo.AddItem rs.Fields(intFldIndx).Value ',   i - 1
    rs.MoveNext
  Next i
  cbo.Text = cbo.List(0)
  Set rs = Nothing
  
End Sub

