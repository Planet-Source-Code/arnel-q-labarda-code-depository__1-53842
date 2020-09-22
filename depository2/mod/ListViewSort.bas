Attribute VB_Name = "ListViewSort"
Option Explicit

Public Const sortAlphanumeric = 0
Public Const sortNumeric = 1
Public Const sortDate = 2
Public Const sortAscending = 3
Public Const sortDescending = 4


Function SortColumn(ByVal ListViewControl As MSComctlLib.ListView, ColumnIndex As Integer, SortType As Integer, SortOrder As Integer) As Boolean
Dim x As Integer, y As Integer
On Error GoTo ErrHandler
    
Select Case SortType
  Case sortAlphanumeric
      DoSort ListViewControl, SortOrder, ColumnIndex - 1
  Case sortNumeric
    Dim strMax As String, strNew As String
    If ColumnIndex > 1 Then
      For x = 1 To ListViewControl.ListItems.Count
        If Len(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)) <> 0 Then
          If Len(CStr(Int(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)))) > Len(strMax) Then
            strMax = CStr(Int(ListViewControl.ListItems(x).SubItems(ColumnIndex - 1)))
          End If
        End If
      Next
    Else
      For x = 1 To ListViewControl.ListItems.Count
        If Len(ListViewControl.ListItems(x)) <> 0 Then
          If Len(CStr(Int(ListViewControl.ListItems(x)))) > Len(strMax) Then
            strMax = CStr(Int(ListViewControl.ListItems(x)))
          End If
        End If
      Next
    End If
    ListViewControl.Visible = False
    If ColumnIndex > 1 Then
      For x = 1 To ListViewControl.ListItems.Count
        If Len(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)) = 0 Then
        ElseIf Len(CStr(Int(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)))) < Len(strMax) Then
          strNew = ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)
          For y = 1 To Len(strMax) - Len(CStr(Int(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1))))
            strNew = "0" & strNew
          Next
          ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = strNew
        End If
      Next
    Else
      For x = 1 To ListViewControl.ListItems.Count
        If Len(ListViewControl.ListItems(x).Text) = 0 Then
          ElseIf Len(CStr(Int(ListViewControl.ListItems(x)))) < Len(strMax) Then
            strNew = ListViewControl.ListItems(x).Text
            For y = 1 To Len(strMax) - Len(CStr(Int(ListViewControl.ListItems(x))))
              strNew = "0" & strNew
            Next
            ListViewControl.ListItems(x).Text = strNew
          End If
      Next
    End If
            
    DoSort ListViewControl, SortOrder, ColumnIndex - 1
    If ColumnIndex > 1 Then
      For x = 1 To ListViewControl.ListItems.Count
        ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = CDbl(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1))
        If ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = 0 Then ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = ""
      Next
    Else
      For x = 1 To ListViewControl.ListItems.Count
        ListViewControl.ListItems(x).Text = CDbl(ListViewControl.ListItems(x).Text)
        If ListViewControl.ListItems(x).Text = 0 Then ListViewControl.ListItems(x).Text = ""
      Next
    End If
    ListViewControl.Visible = True
                
  Case sortDate
    ListViewControl.Visible = False
    If ColumnIndex > 1 Then
      For x = 1 To ListViewControl.ListItems.Count
        ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = Format(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1), "YYYY MM DD hh:mm:ss")
      Next
      DoSort ListViewControl, SortOrder, ColumnIndex - 1
      For x = 1 To ListViewControl.ListItems.Count
        ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = Format(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1), "General Date")
      Next
    Else
      For x = 1 To ListViewControl.ListItems.Count
        ListViewControl.ListItems(x).Text = Format(ListViewControl.ListItems(x).Text, "YYYY MM DD hh:mm:ss")
      Next
      DoSort ListViewControl, SortOrder, ColumnIndex - 1
      For x = 1 To ListViewControl.ListItems.Count
        ListViewControl.ListItems(x).Text = Format(ListViewControl.ListItems(x).Text, "General Date")
      Next
    End If
    ListViewControl.Visible = True
End Select
    SortColumn = True
                
Exit_Function:
    Exit Function
                
ErrHandler:
    MsgBox err.Description & " (" & err.Number & ")", vbOKOnly + vbCritical, "ListView Sort module Error"
    SortColumn = False
    Resume Exit_Function
End Function


Private Sub DoSort(ByVal ListViewControl As MSComctlLib.ListView, SortOrder As Integer, SortKey As Integer)


    If SortOrder = sortAscending Then
        ListViewControl.SortOrder = lvwAscending
    ElseIf SortOrder = sortDescending Then
        ListViewControl.SortOrder = lvwDescending
    End If
    ListViewControl.SortKey = SortKey
    ListViewControl.Sorted = True
End Sub
